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

import json
import os
from pathlib import Path
from typing import Annotated, Any, Literal, TypedDict

from langchain_core.exceptions import OutputParserException
from langgraph.graph import END, START, StateGraph
from pydantic import BaseModel, Field, ValidationError, model_validator

from pptx import Presentation
from pptx.util import Emu

from .renderer import render_deck
from .slide_reuse import (
    allocate_by_type,
    budget_note,
    capacities,
    catalog,
    items_for,
    items_wanted,
)
from .schema import (
    BulletsSlide,
    Deck,
    ImpactRow,
    ImpactsSlide,
    Kpi,
    KpiSlide,
    Milestone,
    Recommendation,
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

OUTLINE_PROMPT = """Structure attendue : couverture (title), puis contexte \
(bullets/reference), analyse d'impact (impacts/kpi), échéancier (timeline), \
recommandations (summary). Le premier élément est toujours de type "title", \
le dernier de type "summary". Découpe en 6 à 12 slides.

Le type "section" n'est qu'un intercalaire : il ne porte AUCUN contenu, \
seulement un titre de partie. Au plus un "section" par grande partie, jamais \
deux "section" consécutifs, et la majorité des slides doit porter du contenu \
réel (bullets, reference, impacts, kpi, timeline, summary). Une note \
composée d'intercalaires est inutilisable.

Pour chaque slide, donne uniquement son type et, en une phrase, ce qu'elle \
doit démontrer — le contenu détaillé sera rédigé séparément pour chacune."""

# Slide types that actually carry content, as opposed to `title`/`section`
# which only carry a heading.
CONTENT_TYPES = frozenset({"bullets", "reference", "impacts", "kpi", "timeline", "summary"})


class _PlanItem(BaseModel):
    type: SlideType
    topic: str = Field(max_length=200, description="Ce que cette slide doit démontrer")


class _Item(BaseModel):
    """One piece of content: an optional lead-in, and the substance.

    Everything a slide carries has this shape — a figure and what it counts,
    a date and what happens then, a recommendation and its justification, or
    simply a point with no lead-in. One schema the model fills for every kind
    of slide, instead of eight different ones.
    """

    lead: str = Field(
        default="",
        max_length=40,
        description="Chiffre, date ou titre court. Vide si le point n'en a pas.",
    )
    text: str = Field(max_length=220, description="Le fond du point")


class _Block(BaseModel):
    """A section of the note, before it is mapped onto any slide design."""

    heading: str = Field(max_length=140, description="Le message de cette partie")
    items: list[_Item] = Field(min_length=1, max_length=6)


class _Outline(BaseModel):
    # Optional because the model routinely omits it, and it is recoverable:
    # the cover slide's topic, or failing that the request itself, says what
    # the deck is called. Rejecting the whole outline over it is wasteful.
    title: str = Field(default="", max_length=80, description="Titre du deck")
    entity: str = Field(max_length=60, description="Institution destinataire")
    plan: list[_PlanItem] = Field(min_length=3, max_length=12)

    @model_validator(mode="after")
    def _enough_content(self) -> "_Outline":
        """Reject an outline made mostly of dividers.

        A local model left to itself picks `section` for nearly every slide —
        it is the cheapest type to produce, being just a heading. Rejecting
        here routes the problem back through `_invoke_with_retries`, which
        hands the model this message and asks again.
        """
        content = sum(1 for item in self.plan if item.type in CONTENT_TYPES)
        if content < len(self.plan) - content:
            raise ValueError(
                f"Seulement {content} slides de contenu sur {len(self.plan)} : le plan est "
                "majoritairement composé d'intercalaires « section », qui ne portent aucun "
                "contenu. Remplace-les par des slides bullets / reference / impacts / kpi / "
                "timeline / summary."
            )
        return self


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
    design_indices: list[int | None]
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


def _flatten_to_str(value: dict) -> str | None:
    """Squash an object the model returned where a string was expected.

    Seen in the wild: `source` came back as
    `{"article": "12", "number": "BAM n° 5/W/2022"}`. The information is
    correct, only the shape is wrong, so the string values are joined rather
    than the answer thrown away.
    """
    for key in ("description", "value", "title", "text", "label", "name"):
        if isinstance(value.get(key), str):
            return value[key]
    parts = [v for v in value.values() if isinstance(v, str) and v.strip()]
    return ", ".join(parts) if parts else None


def _hoist_nested_payload(args: dict, field_names: set[str]) -> dict:
    """Undo an extra level of nesting under an arbitrary field.

    The model sometimes puts the whole object inside one of its own fields —
    `{"action_title": {"kpis": [...], "type": "kpi"}}`. If a field's value is
    an object holding *other* fields of this schema, it is the real payload.
    """
    for name, value in args.items():
        if isinstance(value, dict) and (set(value) & field_names) - {name}:
            merged = {k: v for k, v in args.items() if k != name}
            merged.update(value)
            return merged
    return args


def _repair_values(args: Any, model_cls: type[BaseModel], defaults: dict | None = None) -> Any:
    """Fixes the recurring Qwen artifacts, even on a simple schema:
    - the whole payload nested one level down under one of its own fields;
    - a field rendered as an object (a copied schema fragment, or structured
      data) where a string belongs;
    - a string over the schema's `max_length`, or a list over its `maxItems` —
      trimmed rather than rejecting an otherwise usable answer, which is
      exactly what those constraints are for (see schema.py);
    - a required field simply missing, where the caller can supply a sane
      default (the slide's topic makes a serviceable title).
    """
    if not isinstance(args, dict):
        return args
    properties = model_cls.model_json_schema().get("properties", {})
    args = _hoist_nested_payload(dict(args), set(properties))

    for name, prop in properties.items():
        possible_values = prop.get("enum") or ([prop["const"]] if "const" in prop else None)
        if possible_values and len(possible_values) == 1:
            # Single-value field already known before the call (a slide's
            # `type` discriminator, e.g.) — no need to trust the model to
            # reproduce it exactly rather than just imposing the one valid
            # value (observed: 'impact' instead of 'impacts').
            args[name] = possible_values[0]

    for name, default in (defaults or {}).items():
        if name in properties and not args.get(name):
            args[name] = default

    for name, value in list(args.items()):
        if isinstance(value, dict):
            flattened = _flatten_to_str(value)
            if flattened is None:
                del args[name]
                continue
            value = flattened
            args[name] = value
        if isinstance(value, str):
            limit = _max_length(properties.get(name, {}))
            if limit:
                args[name] = _truncate(value, limit)
        elif isinstance(value, list):
            limit = _max_length(properties.get(name, {}).get("items", {}))
            if limit:
                value = [_truncate(v, limit) if isinstance(v, str) else v for v in value]
            # Asked for "6 to 12 slides" the model cheerfully returns 14.
            # Keeping the first N beats rejecting an otherwise good answer.
            max_items = properties.get(name, {}).get("maxItems")
            if max_items and len(value) > max_items:
                value = value[:max_items]
            args[name] = value
    return args


def _json_from_text(content: Any) -> dict | None:
    """Pull a JSON object out of message text.

    Faced with the busiest schemas the model sometimes answers in prose with
    the JSON inline instead of calling the tool. The content is right there —
    no reason to throw the answer away over how it was delivered.
    """
    if not isinstance(content, str):
        return None
    start, end = content.find("{"), content.rfind("}")
    if start == -1 or end <= start:
        return None
    try:
        parsed = json.loads(content[start : end + 1])
    except json.JSONDecodeError:
        return None
    return parsed if isinstance(parsed, dict) else None


def _extract(raw: dict, model_cls: type[BaseModel], defaults: dict | None = None) -> BaseModel:
    """Extracts an instance of `model_cls` from a
    `with_structured_output(..., include_raw=True)` result."""
    if raw.get("parsed") is not None:
        return raw["parsed"]

    error = raw.get("parsing_error")
    message = raw.get("raw")
    tool_calls = getattr(message, "tool_calls", None) or []

    candidates = [call.get("args", {}) for call in tool_calls[:1]]
    from_text = _json_from_text(getattr(message, "content", None))
    if from_text is not None:
        candidates.append(from_text)

    for args in candidates:
        args = _repair_values(_repair_envelope(args, model_cls.__name__), model_cls, defaults)
        try:
            return model_cls.model_validate(args)
        except ValidationError as e:
            error = e
    raise error or RuntimeError(f"No valid tool call for {model_cls.__name__}.")


def _invoke_with_retries(
    planner: Any, messages: list, model_cls: type[BaseModel], defaults: dict | None = None
) -> BaseModel:
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
            return _extract(planner.invoke(attempt), model_cls, defaults)
        except (ValidationError, OutputParserException, RuntimeError) as e:
            error = str(e)[:1200]
    raise RuntimeError(
        f"The model failed to produce a valid {model_cls.__name__} after "
        f"{MAX_ATTEMPTS} attempts: {error}"
    )


def _fill_ratio(slide, caps: dict | None) -> float:
    """How much of the design's room the written text actually occupies.

    A slide whose text uses a third of the space it was given reads as empty
    — a band of words across the top of an otherwise bare board.
    """
    if not caps:
        return 1.0
    room = caps.get("title", 0) + max(1, caps.get("columns", 1)) * sum(caps.get("slots") or [0])
    if room <= 0:
        return 1.0
    written = len(getattr(slide, "action_title", None) or getattr(slide, "title", "") or "")
    written += sum(len(part) for pair in items_for(slide) for part in pair)
    return written / room


# Below this the slide is mostly empty space, and the model is asked to
# develop its content rather than the design being abandoned.
_MIN_FILL = 0.55


def _block_to_slide(slide_type: str, block: _Block, number: int, topic: str):
    """Map a written block onto the slide type that will present it.

    The model writes content; this decides presentation. Where a type's
    minimum can't be met — two figures for a `kpi`, two milestones for a
    `timeline` — the block is presented as bullets rather than discarded.
    """
    heading = _truncate(block.heading or topic, 120)
    items = block.items

    def leads_and_texts(lead_max: int, text_max: int, count: int):
        return [
            (_truncate(i.lead or "—", lead_max), _truncate(i.text, text_max))
            for i in items[:count]
        ]

    if slide_type == "kpi" and len(items) >= 2:
        pairs = leads_and_texts(12, 60, 4)
        return KpiSlide(
            action_title=heading, kpis=[Kpi(value=v, label=l) for v, l in pairs]
        )

    if slide_type == "timeline" and len(items) >= 2:
        pairs = leads_and_texts(24, 90, 5)
        return TimelineSlide(
            action_title=heading,
            milestones=[Milestone(date=d, label=l) for d, l in pairs],
        )

    if slide_type == "summary" and len(items) >= 2:
        pairs = leads_and_texts(70, 200, 4)
        return SummarySlide(
            action_title=heading,
            recommendations=[
                Recommendation(title=t if t != "—" else _truncate(d, 70), detail=d)
                for t, d in pairs
            ],
        )

    if slide_type == "reference":
        return ReferenceSlide(
            action_title=heading,
            reference=_truncate(items[0].lead or topic, 110),
            excerpt=_truncate(items[0].text, 600),
            interpretation=[_truncate(i.text, 180) for i in items[1:4]]
            or [_truncate(items[0].text, 180)],
        )

    if slide_type == "impacts" and len(items) >= 2:
        return ImpactsSlide(
            action_title=heading,
            rows=[
                ImpactRow(
                    requirement=_truncate(i.lead or i.text, 90),
                    impact=_truncate(i.text, 140),
                    entity="—",
                    criticality="Moyenne",
                )
                for i in items[:6]
            ],
        )

    # Bullets is the universal fallback: any block can be shown as points.
    points = [_truncate(i.text, 200) for i in items[:5]]
    while len(points) < 2:
        points.append(_truncate(topic, 200))
    return BulletsSlide(action_title=heading, points=points)


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


def _drop_consecutive_sections(plan: list[_PlanItem]) -> list[_PlanItem]:
    """Two dividers in a row means an empty slide announcing another empty
    slide. Dropping the redundant one is unambiguous, so it's fixed here
    instead of costing a whole re-generation."""
    kept: list[_PlanItem] = []
    for item in plan:
        if item.type == "section" and kept and kept[-1].type == "section":
            continue
        kept.append(item)
    # A divider is a promise of content after it. Ending on one leaves the
    # deck finishing on an empty slide announcing nothing.
    while len(kept) > 3 and kept[-1].type == "section":
        kept.pop()
    return kept


def _renumber_sections(slides: list) -> None:
    """Number the dividers in order.

    Their position in the deck decides their number, so there is no reason to
    let the model guess it — left alone it repeats itself (two slides both
    numbered 09)."""
    counter = 0
    for slide in slides:
        if slide.type == "section":
            counter += 1
            slide.number = min(counter, 9)  # the schema allows 1-9


def build_graph(
    llm: Any,
    themes_dir: Path | None = None,
    theme: Theme | None = None,
    template_path: Path | None = None,
):
    """`llm`: any LangChain model supporting `with_structured_output`.

    `theme`: overrides the per-tenant YAML lookup — used to pass in a style
    measured from a client's own template.
    `template_path`: that same template, so slides can be built on its own
    designs rather than only in its colours.
    """

    # method="function_calling": the only reliable method on Qwen/Ollama —
    # Ollama's native "json_schema" mode ignores the `type` discriminator of
    # unions. include_raw=True: needed to unwrap Qwen's output, see
    # `_extract`.
    outline_planner = llm.with_structured_output(_Outline, method="function_calling", include_raw=True)
    slide_planners = {
        t: llm.with_structured_output(cls, method="function_calling", include_raw=True)
        for t, cls in _SLIDE_SCHEMAS.items()
    }

    # The template is read once here, not per run: which of its designs each
    # slide will use has to be decided *before* the text is written, so the
    # text can be written to the size of the box it lands in.
    source_prs, designs = None, []
    probe_w = probe_h = 0.0
    if template_path:
        source_prs = Presentation(str(template_path))
        probe_w = Emu(source_prs.slide_width).inches
        probe_h = Emu(source_prs.slide_height).inches
        designs = catalog(source_prs, probe_w, probe_h)

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
        items = _drop_consecutive_sections(_move_title_to_front(outline.plan))
        deck_title = _truncate(
            outline.title or next((i.topic for i in items if i.type == "title"), "") or state["request"],
            80,
        )

        # Each slide is written by its own call, which otherwise sees nothing
        # but its own one-line topic — so slides repeat each other and the note
        # reads as a pile of fragments. Handing every call the whole plan, and
        # what has already been written, is what makes them cohere. The plan
        # sits before the per-slide instruction so the shared prefix stays
        # identical across calls and the model's prompt cache still applies.
        plan_overview = "\n".join(
            f"{n}. [{it.type}] {it.topic}" for n, it in enumerate(items, start=1)
        )
        plan_message = (
            "human",
            f"Plan complet de la note (ne traite QUE la slide qu'on te demande, "
            f"mais situe-la dans cet ensemble) :\n{plan_overview}",
        )
        written: list[str] = []

        # Decide which of the template's designs each slide will use *now*, so
        # the text can be written to the size of the box it will occupy. Their
        # boxes hold as little as 20 characters; our schema allows 120, which
        # is why text was landing on the slide as « Incident… ».
        allocation = (
            allocate_by_type(designs, [it.type for it in items], [None] * len(items))
            if designs
            else [None] * len(items)
        )
        caps_for = [capacities(source_prs, d, probe_w, probe_h) if d else None for d in allocation]
        budgets = [budget_note(c) if c else "" for c in caps_for]

        # How many items the chosen design shows, so the writing matches the
        # design rather than being stretched or padded to fit it afterwards.
        # A repeated row wants one item per column; any other design wants as
        # many as its boxes divide into, which `items_wanted` works out from
        # the design alone.
        wanted_items = [
            items_wanted(d, it.type) if d else 3 for d, it in zip(allocation, items)
        ]

        def write_slide(index: int, slide_type: str, topic: str, expand: str = ""):
            count = wanted_items[index - 1]
            # Where the chosen design gives each element a short line above a
            # longer one, a point written as « intitulé : explication » fills
            # both. Without it the lead boxes stay empty and every element
            # shows a paragraph under a blank caption. Driven by what the
            # design actually has, so a template whose elements are a single
            # box is never asked for a form it has nowhere to put.
            caps = caps_for[index - 1]
            paired = (
                "\n\nÉcris chaque point sous la forme « intitulé court : "
                "explication », l'intitulé nommant l'objet en quelques mots."
                if slide_type == "bullets" and caps and len(caps.get("slots") or []) > 1
                else ""
            )
            messages = [
                *header,
                plan_message,
                (
                    "human",
                    f"Rédige uniquement la slide {index}/{len(items)}, de type "
                    f"« {slide_type} ». Sujet de cette slide : {topic}\n"
                    f"Titre du deck : {deck_title}. Entité destinataire : {outline.entity}.\n"
                    f"Prévois environ {count} élément(s), pour tenir dans la maquette."
                    + (f"\n\n{budgets[index - 1]}" if budgets[index - 1] else "")
                    + paired
                    # Without this the top-up pass re-sent a prompt identical
                    # to the first one and simply drew another sample, so a
                    # slide judged too empty had no reason to come back fuller.
                    + (f"\n\n{expand}" if expand else "")
                    + (
                        "\n\nTitres déjà rédigés dans cette note — n'y reviens pas, "
                        "n'en répète ni le message ni les exemples :\n"
                        + "\n".join(f"- {t}" for t in written)
                        if written
                        else ""
                    ),
                ),
            ]
            return _invoke_with_retries(
                slide_planners[slide_type],
                messages,
                _SLIDE_SCHEMAS[slide_type],
                defaults={
                    "action_title": _truncate(topic, 120),
                    "title": _truncate(topic, 60),
                    "number": index,
                },
            )

        slides = []
        for i, item in enumerate(items, start=1):
            # Progress on stdout. A run takes twenty minutes on CPU and used to
            # print nothing between the measurement and the finished file, which
            # leaves anyone watching — at the terminal or through the web
            # interface — unable to tell a slow slide from a hung one.
            print(f"[slide {i}/{len(items)}] {item.type} — {item.topic[:70]}", flush=True)
            if item.type == "title":
                # The cover's topic is usually a restatement of the deck title,
                # which would print the same sentence twice — once large, once
                # truncated in the subtitle slot above it.
                topic = _truncate(item.topic, 120)
                duplicate = topic[:40].lower() in deck_title.lower() or (
                    deck_title[:40].lower() in topic.lower()
                )
                slides.append(
                    TitleSlide(
                        title=deck_title,
                        subtitle=None if duplicate else topic,
                        reference=None,
                    )
                )
                continue
            if item.type == "section":
                slides.append(SectionSlide(number=1, title=_truncate(item.topic, 60)))
                continue

            try:
                slide = write_slide(i, item.type, item.topic)
            except RuntimeError:
                # The busiest schemas are the ones a small model chokes on.
                # Falling back to bullets keeps the topic in the deck.
                if item.type == "bullets":
                    raise
                slide = write_slide(i, "bullets", item.topic)
            caps = caps_for[i - 1]
            if _fill_ratio(slide, caps) < _MIN_FILL:
                # The text doesn't fill the design it was written for. Rather
                # than leave the slide bare or abandon the design, the model is
                # asked to develop the same points using the source material.
                try:
                    fuller = write_slide(
                        i, item.type, item.topic,
                        expand=(
                            "Ta version précédente ne remplit qu'une fraction de la "
                            "place disponible et la slide paraît vide. Reprends-la en "
                            "développant chaque point à partir du contexte fourni — "
                            "précise les obligations, les entités concernées, les "
                            "échéances — sans inventer de faits absents du contexte, "
                            "et sans ajouter de slide."
                        ),
                    )
                    if _fill_ratio(fuller, caps) > _fill_ratio(slide, caps):
                        slide = fuller
                except RuntimeError:
                    pass  # the first version stands
            slides.append(slide)
            heading = getattr(slide, "action_title", None) or getattr(slide, "title", None)
            if heading:
                written.append(heading)

        _renumber_sections(slides)
        deck = Deck(title=deck_title, entity=outline.entity, slides=slides)
        return {
            "deck": deck,
            "iterations": state.get("iterations", 0) + 1,
            # The renderer must use the designs the text was written for, not
            # re-decide and land it in differently-sized boxes.
            "design_indices": [d.index if d else None for d in allocation],
        }

    def render(state: DeckState) -> DeckState:
        print(f"[render] {len(state['deck'].slides)} slides", flush=True)
        style = theme or Theme.load(state.get("tenant", "default"), base_dir=themes_dir)
        out_dir = Path(state.get("output_dir", "/tmp/atlas-decks"))
        path = out_dir / f"note-reglementaire-{state['deck'].entity[:20]}.pptx"
        result = render_deck(
            state["deck"], style, path,
            template_path=template_path,
            design_indices=state.get("design_indices"),
        )
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
            print(
                f"[qa] {len(state['warnings'])} avertissement(s) — nouvelle passe",
                flush=True,
            )
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
