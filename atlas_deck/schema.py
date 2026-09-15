"""Deck contract — the only structure the LLM is allowed to produce.

The length constraints aren't cosmetic: they guarantee the renderer will
never have to handle an overflow. All the layout discipline lives here, not
in the prompt.
"""

from typing import Annotated, Literal, Union

from pydantic import BaseModel, Field

# Values stay in French: they are the exact strings rendered on the slide
# (the deck's content is a French regulatory note) and the exact keys the
# theme's `criticality` map is looked up by.
Criticality = Literal["Faible", "Moyenne", "Élevée"]


# --------------------------------------------------------------------------
# Slide types (closed inventory — regulatory note)
# --------------------------------------------------------------------------


class TitleSlide(BaseModel):
    """Cover. Always the first slide, exactly one per deck."""

    type: Literal["title"] = "title"
    title: str = Field(max_length=80)
    subtitle: str | None = Field(default=None, max_length=120)
    reference: str | None = Field(
        default=None, max_length=90, description="Regulatory text this note relates to"
    )


class SectionSlide(BaseModel):
    """Numbered section divider."""

    type: Literal["section"] = "section"
    number: int = Field(ge=1, le=9)
    title: str = Field(max_length=60)


class BulletsSlide(BaseModel):
    """Substance slide: a message-title, 2 to 5 supporting points."""

    type: Literal["bullets"] = "bullets"
    action_title: str = Field(
        max_length=120,
        description="The title carries the message, not a label. "
        "« Trois obligations nouvelles pèsent sur le dispositif LCB-FT » "
        "and not « Obligations ».",
    )
    points: list[Annotated[str, Field(max_length=200)]] = Field(
        min_length=2, max_length=5
    )
    source: str | None = Field(default=None, max_length=120)


class ReferenceSlide(BaseModel):
    """Regulatory text excerpt on the left, interpretation on the right."""

    type: Literal["reference"] = "reference"
    action_title: str = Field(max_length=120)
    reference: str = Field(
        max_length=110, description="E.g. « Circulaire BAM n° 5/W/2022, article 12 »"
    )
    excerpt: str = Field(max_length=600, description="Verbatim quote of the text")
    interpretation: list[Annotated[str, Field(max_length=180)]] = Field(
        min_length=1, max_length=3, description="What the text concretely implies"
    )


class ImpactRow(BaseModel):
    requirement: str = Field(max_length=90)
    impact: str = Field(max_length=140)
    entity: str = Field(max_length=40, description="Responsible entity / function")
    criticality: Criticality


class ImpactsSlide(BaseModel):
    """Impact-analysis table — the core of a regulatory note."""

    type: Literal["impacts"] = "impacts"
    action_title: str = Field(max_length=120)
    rows: list[ImpactRow] = Field(min_length=2, max_length=6)
    source: str | None = Field(default=None, max_length=120)


class Kpi(BaseModel):
    value: str = Field(max_length=12, description="E.g. « 18 mois », « 3 », « 12 % »")
    label: str = Field(max_length=60)


class KpiSlide(BaseModel):
    """Key figures (deadlines, volumes, regulatory thresholds)."""

    type: Literal["kpi"] = "kpi"
    action_title: str = Field(max_length=120)
    kpis: list[Kpi] = Field(min_length=2, max_length=4)
    source: str | None = Field(default=None, max_length=120)


class Milestone(BaseModel):
    date: str = Field(max_length=24, description="E.g. « 31/12/2026 », « T2 2027 »")
    label: str = Field(max_length=90)


class TimelineSlide(BaseModel):
    """Compliance rollout timeline."""

    type: Literal["timeline"] = "timeline"
    action_title: str = Field(max_length=120)
    milestones: list[Milestone] = Field(min_length=2, max_length=5)


class Recommendation(BaseModel):
    title: str = Field(max_length=70)
    detail: str = Field(max_length=200)


class SummarySlide(BaseModel):
    """Numbered recommendations. Closes the deck."""

    type: Literal["summary"] = "summary"
    action_title: str = Field(max_length=120)
    recommendations: list[Recommendation] = Field(min_length=2, max_length=4)


Slide = Union[
    TitleSlide,
    SectionSlide,
    BulletsSlide,
    ReferenceSlide,
    ImpactsSlide,
    KpiSlide,
    TimelineSlide,
    SummarySlide,
]


class Deck(BaseModel):
    """The full deck. This is the object produced by `llm.with_structured_output(Deck)`."""

    title: str = Field(max_length=80)
    entity: str = Field(max_length=60, description="Recipient institution")
    slides: list[Annotated[Slide, Field(discriminator="type")]] = Field(
        min_length=3, max_length=25
    )

    def model_post_init(self, __context) -> None:  # noqa: D105
        if self.slides and self.slides[0].type != "title":
            raise ValueError("The first slide must be of type 'title'.")
