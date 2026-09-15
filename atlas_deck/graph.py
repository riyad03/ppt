"""LangGraph subgraph « regulatory note generation ».

    plan ──▶ render ──▶ qa ──┬── (warnings & iterations < MAX) ──▶ plan
                             └── ok ──▶ END

Only the `plan` node calls the model. `render` and `qa` are deterministic,
which makes the pipeline testable and reproducible — essential for a
deliverable that ends up in an audit folder.

The `plan` node proceeds in two steps rather than one giant call:
1. an outline (title, entity, ordered list of {type, topic}) — a small flat
   schema;
2. one slide at a time, each with its own simple schema (no discriminated
   union for the model to juggle).
A modest local model regularly fails to correctly fill the large `Deck`
schema (8 slide types) in a single function-calling call — the same task
broken down into small calls, each with a trivial schema, succeeds far more
reliably.
"""

from __future__ import annotations

import os
from pathlib import Path
from typing import Annotated, Any, Literal, TypedDict

from langchain_core.exceptions import OutputParserException
from langgraph.graph import END, START, StateGraph
from pydantic import BaseModel, Field, ValidationError

from .renderer import render_deck
from .schema import (
    BulletsSlide,
    Deck,
    ImpactsSlide,
    KpiSlide,
    ReferenceSlide,
    SectionSlide,
    Slide,
    SummarySlide,
    TimelineSlide,
    TitleSlide,
)
from .theme import Theme

MAX_ITERATIONS = 2
MAX_ATTEMPTS = 3

SlideType = Literal[
    "title", "section", "bullets", "reference", "impacts", "kpi", "timeline", "summary"
]

_SLIDE_SCHEMAS: dict[str, type[Slide]] = {
    "title": TitleSlide,
    "section": SectionSlide,
    "bullets": BulletsSlide,
    "reference": ReferenceSlide,
    "impacts": ImpactsSlide,
    "kpi": KpiSlide,
    "timeline": TimelineSlide,
    "summary": SummarySlide,
}

# Kept in French on purpose: this is the actual content specification sent to
# the LLM — it instructs it to write a French regulatory note for a French
# financial institution. Translating it would change the deck's output
# language, not just the code.
SYSTEM_PROMPT = """Tu rédiges une note réglementaire destinée aux instances de \
gouvernance d'une institution financière (comité d'audit, comité des risques, \
direction générale).

Règles éditoriales, non négociables :
- Chaque titre de slide est un MESSAGE, pas une étiquette. « Trois obligations \
nouvelles pèsent sur le dispositif LCB-FT » et non « Obligations ».
- Une idée par slide. Si deux idées, deux slides.
- Toute affirmation reposant sur un texte réglementaire cite sa source \
(référence, article, date).
- Tu ne paraphrases jamais un article de manière approximative : soit tu le \
cites verbatim dans une slide `reference`, soit tu en donnes la portée.
- Pas de superlatif, pas de formulation commerciale. Registre de note interne."""

OUTLINE_PROMPT = """Structure attendue : couverture (title), puis section(s), \
contexte (bullets/reference), analyse d'impact (impacts/kpi), échéancier \
(timeline), recommandations (summary). Le premier élément est toujours de \
type "title". Découpe en 3 à 12 slides. Pour chaque slide, donne uniquement \
son type et, en une phrase, ce qu'elle doit démontrer — le contenu détaillé \
sera rédigé séparément pour chacune."""


class _PlanItem(BaseModel):
    type: SlideType
    topic: str = Field(max_length=200, description="Ce que cette slide doit démontrer")


class _Outline(BaseModel):
    title: str = Field(max_length=80, description="Titre du deck")
    entity: str = Field(max_length=60, description="Institution destinataire")
    plan: list[_PlanItem] = Field(min_length=3, max_length=12)


class DeckState(TypedDict, total=False):
    # Inputs
    request: str
    context: str
    tenant: str
    output_dir: str
    # Internal state
    deck: Deck
    warnings: Annotated[list[str], lambda a, b: b]
    iterations: int
    # Output
    pptx_path: str


def _repair_envelope(args: Any, schema_name: str) -> Any:
    """Unwraps the case where Qwen nests the arguments under the schema's name
    (e.g. `{"BulletsSlide": {...}}` instead of `{...}` directly)."""
    if isinstance(args, dict) and set(args.keys()) == {schema_name} and isinstance(args[schema_name], dict):
        return args[schema_name]
    return args


def _max_length(prop: dict) -> int | None:
    if "maxLength" in prop:
        return prop["maxLength"]
    for sub in prop.get("anyOf", []) or []:
        if "maxLength" in sub:
            return sub["maxLength"]
    return None


def _truncate(value: str, limit: int) -> str:
    return value if len(value) <= limit else value[: limit - 1].rstrip() + "…"


def _repair_values(args: Any, model_cls: type[BaseModel]) -> Any:
    """Fixes two recurring Qwen artifacts, even on a simple schema:
    - an optional field rendered as a copied schema fragment
      (`{"description": "..."}`) or an empty object instead of a string or an
      omission;
    - a string that exceeds the schema's `max_length` constraint — truncated
      rather than rejecting an otherwise usable response, which is exactly
      what these constraints are for (see schema.py).
    """
    if not isinstance(args, dict):
        return args
    properties = model_cls.model_json_schema().get("properties", {})
    args = dict(args)

    for name, prop in properties.items():
        possible_values = prop.get("enum") or ([prop["const"]] if "const" in prop else None)
        if possible_values and len(possible_values) == 1:
            # Single-value field already known before the call (a slide's
            # `type` discriminator, e.g.) — no need to trust the model to
            # reproduce it exactly rather than just imposing the one valid
            # value (observed: 'impact' instead of 'impacts').
            args[name] = possible_values[0]

    for name, value in list(args.items()):
        if isinstance(value, dict):
            if not value:
                del args[name]
                continue
            for key in ("description", "value", "title", "text"):
                if isinstance(value.get(key), str):
                    value = value[key]
                    args[name] = value
                    break
        if isinstance(value, str):
            limit = _max_length(properties.get(name, {}))
            if limit:
                args[name] = _truncate(value, limit)
        elif isinstance(value, list):
            limit = _max_length(properties.get(name, {}).get("items", {}))
            if limit:
                args[name] = [_truncate(v, limit) if isinstance(v, str) else v for v in value]
    return args


def _extract(raw: dict, model_cls: type[BaseModel]) -> BaseModel:
    """Extracts an instance of `model_cls` from a
    `with_structured_output(..., include_raw=True)` result."""
    if raw.get("parsed") is not None:
        return raw["parsed"]

    error = raw.get("parsing_error")
    tool_calls = getattr(raw.get("raw"), "tool_calls", None) or []
    if tool_calls:
        args = _repair_envelope(tool_calls[0].get("args", {}), model_cls.__name__)
        args = _repair_values(args, model_cls)
        try:
            return model_cls.model_validate(args)
        except ValidationError as e:
            error = e
    raise error or RuntimeError(f"No valid tool call for {model_cls.__name__}.")


def _invoke_with_retries(planner: Any, messages: list, model_cls: type[BaseModel]) -> BaseModel:
    """A modest local model sometimes fails the expected schema. We feed the
    validation error back rather than giving up — the same corrective-loop
    logic as for layout warnings."""
    error: str | None = None
    for _ in range(MAX_ATTEMPTS):
        attempt = list(messages)
        if error:
            attempt.append(
                (
                    "human",
                    "Ta réponse précédente ne respectait pas le schéma attendu :\n"
                    f"{error}\nCorrige et renvoie une réponse complète et valide.",
                )
            )
        try:
            return _extract(planner.invoke(attempt), model_cls)
        except (ValidationError, OutputParserException, RuntimeError) as e:
            error = str(e)[:1200]
    raise RuntimeError(
        f"The model failed to produce a valid {model_cls.__name__} after "
        f"{MAX_ATTEMPTS} attempts: {error}"
    )


def _move_title_to_front(plan: list[_PlanItem]) -> list[_PlanItem]:
    """`Deck` requires a `title` slide in first position. We fix the order
    rather than re-asking the model for a whole new outline over an
    ordering detail."""
    idx = next((i for i, e in enumerate(plan) if e.type == "title"), None)
    if idx in (None, 0):
        return plan
    plan = list(plan)
    plan.insert(0, plan.pop(idx))
    return plan


def build_graph(llm: Any, themes_dir: Path | None = None):
    """`llm`: any LangChain model supporting `with_structured_output`."""

    # method="function_calling": the only reliable method on Qwen/Ollama —
    # Ollama's native "json_schema" mode ignores the `type` discriminator of
    # unions. include_raw=True: needed to unwrap Qwen's output, see
    # `_extract`.
    outline_planner = llm.with_structured_output(_Outline, method="function_calling", include_raw=True)
    slide_planners = {
        t: llm.with_structured_output(cls, method="function_calling", include_raw=True)
        for t, cls in _SLIDE_SCHEMAS.items()
    }

    def plan(state: DeckState) -> DeckState:
        header = [
            ("system", SYSTEM_PROMPT),
            ("human", f"Demande :\n{state['request']}\n\nContexte :\n{state.get('context', '')}"),
        ]
        if state.get("warnings"):
            issues = "\n".join(f"- {a}" for a in state["warnings"])
            header.append(
                (
                    "human",
                    "La version précédente présentait ces défauts de mise en page. "
                    "Reprends le plan en les corrigeant (scinder une slide, réduire "
                    "le nombre de points prévu) sans appauvrir le fond :\n" + issues,
                )
            )

        outline: _Outline = _invoke_with_retries(
            outline_planner, [*header, ("human", OUTLINE_PROMPT)], _Outline
        )
        items = _move_title_to_front(outline.plan)

        slides = []
        for i, item in enumerate(items, start=1):
            slide_messages = [
                *header,
                (
                    "human",
                    f"Rédige uniquement la slide {i}/{len(items)}, de type "
                    f"« {item.type} ». Sujet de cette slide : {item.topic}\n"
                    f"Titre du deck : {outline.title}. Entité destinataire : {outline.entity}.",
                ),
            ]
            slide = _invoke_with_retries(
                slide_planners[item.type], slide_messages, _SLIDE_SCHEMAS[item.type]
            )
            slides.append(slide)

        deck = Deck(title=outline.title, entity=outline.entity, slides=slides)
        return {"deck": deck, "iterations": state.get("iterations", 0) + 1}

    def render(state: DeckState) -> DeckState:
        theme = Theme.load(state.get("tenant", "default"), base_dir=themes_dir)
        out_dir = Path(state.get("output_dir", "/tmp/atlas-decks"))
        path = out_dir / f"note-reglementaire-{state['deck'].entity[:20]}.pptx"
        result = render_deck(state["deck"], theme, path)
        return {
            "pptx_path": str(result.path),
            "warnings": result.warnings,
        }

    def qa(state: DeckState) -> DeckState:
        # Extension point: PNG rendering via LibreOffice + visual review by a
        # vision model. Warnings produced here simply get added to the
        # renderer's and feed the same loop.
        return {}

    def next_step(state: DeckState) -> str:
        if state.get("warnings") and state.get("iterations", 0) < MAX_ITERATIONS:
            return "plan"
        return END

    g = StateGraph(DeckState)
    g.add_node("plan", plan)
    g.add_node("render", render)
    g.add_node("qa", qa)
    g.add_edge(START, "plan")
    g.add_edge("plan", "render")
    g.add_edge("render", "qa")
    g.add_conditional_edges("qa", next_step, {"plan": "plan", END: END})
    return g.compile()


# --------------------------------------------------------------------------
# Exposed as a tool for a higher-level agent
# --------------------------------------------------------------------------


def build_note_tool(llm: Any):
    """Returns a LangChain tool the Atlas Guardian agent can call."""
    from langchain_core.tools import tool

    graph = build_graph(llm)

    @tool
    def generate_regulatory_note(request: str, context: str, tenant: str) -> str:
        """Génère une note réglementaire au format PowerPoint et retourne son chemin.

        request : ce que la note doit démontrer.
        context : extraits réglementaires et éléments de situation de l'entité.
        tenant : identifiant du client, détermine la charte graphique.
        """
        state = graph.invoke(
            {
                "request": request,
                "context": context,
                "tenant": tenant,
                "output_dir": os.getenv("ATLAS_DECK_DIR", "/tmp/atlas-decks"),
            }
        )
        return state["pptx_path"]

    return generate_regulatory_note
