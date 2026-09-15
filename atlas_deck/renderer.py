"""Deterministic rendering engine. No LLM call here.

One helper per slide type, a single grid, all colors and fonts sourced from
`Theme`. Each helper returns the list of layout warnings (estimated
overflow), later surfaced in the graph's state.
"""

from __future__ import annotations

import math
from dataclasses import dataclass, field
from pathlib import Path

from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import MSO_ANCHOR, PP_ALIGN
from pptx.util import Inches, Pt

from .schema import (
    BulletsSlide,
    Deck,
    ImpactsSlide,
    KpiSlide,
    ReferenceSlide,
    SectionSlide,
    SummarySlide,
    TimelineSlide,
    TitleSlide,
)
from .theme import Theme

# --------------------------------------------------------------------------
# Grid — 16:9, 13.333" x 7.5". All values in inches.
# --------------------------------------------------------------------------
SLIDE_W, SLIDE_H = 13.333, 7.5
MARGIN = 0.75
USABLE_W = SLIDE_W - 2 * MARGIN
TITLE_Y = 0.62
TITLE_H = 0.95
CONTENT_Y = 1.85
FOOTER_Y = 6.95
CONTENT_H = FOOTER_Y - 0.25 - CONTENT_Y

COVER_TITLE_PT = 34
TITLE_PT = 21
BODY_PT = 14
SMALL_PT = 11
SOURCE_PT = 9


@dataclass
class RenderResult:
    path: Path
    warnings: list[str] = field(default_factory=list)

    @property
    def ok(self) -> bool:
        return not self.warnings


# --------------------------------------------------------------------------
# Primitives
# --------------------------------------------------------------------------


def _estimate_lines(text: str, width_in: float, pt: float) -> int:
    """Estimated number of lines after word-wrap.

    Deliberately pessimistic approximation (0.55 em average width in Arial):
    a false warning is better than an overflow in production.
    """
    if not text:
        return 0
    chars_per_line = max(1, int(width_in * 72 / (pt * 0.55)))
    return max(1, math.ceil(len(text) / chars_per_line))


def _estimate_height(text: str, width_in: float, pt: float) -> float:
    return _estimate_lines(text, width_in, pt) * (pt * 1.28 / 72)


def _text(
    slide,
    text: str,
    x: float,
    y: float,
    w: float,
    h: float,
    *,
    pt: int,
    color: RGBColor,
    font: str,
    bold: bool = False,
    align=PP_ALIGN.LEFT,
    anchor=MSO_ANCHOR.TOP,
    line_spacing: float = 1.15,
):
    box = slide.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h))
    frame = box.text_frame
    frame.word_wrap = True
    frame.vertical_anchor = anchor
    frame.margin_left = frame.margin_right = 0
    frame.margin_top = frame.margin_bottom = 0
    p = frame.paragraphs[0]
    p.alignment = align
    p.line_spacing = line_spacing
    run = p.add_run()
    run.text = text
    run.font.name = font
    run.font.size = Pt(pt)
    run.font.bold = bold
    run.font.color.rgb = color
    return box


def _rect(slide, x, y, w, h, color: RGBColor, shape=MSO_SHAPE.RECTANGLE):
    s = slide.shapes.add_shape(shape, Inches(x), Inches(y), Inches(w), Inches(h))
    s.fill.solid()
    s.fill.fore_color.rgb = color
    s.line.fill.background()
    s.shadow.inherit = False
    return s


def _blank_slide(prs: Presentation, theme: Theme, background: str = "background"):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    background_shape = slide.background.fill
    background_shape.solid()
    background_shape.fore_color.rgb = theme.color(background)
    return slide


def _action_title(slide, text: str, theme: Theme) -> list[str]:
    _text(
        slide,
        text,
        MARGIN,
        TITLE_Y,
        USABLE_W,
        TITLE_H,
        pt=TITLE_PT,
        color=theme.color("primary"),
        font=theme.heading_font,
        bold=True,
        anchor=MSO_ANCHOR.TOP,
    )
    if _estimate_height(text, USABLE_W, TITLE_PT) > TITLE_H:
        return [f"Title too long, overflows into the content area: « {text[:60]}… »"]
    return []


def _footer(slide, theme: Theme, number: int):
    if theme.footer:
        _text(
            slide,
            theme.footer,
            MARGIN,
            FOOTER_Y,
            USABLE_W - 1,
            0.28,
            pt=SOURCE_PT,
            color=theme.color("muted"),
            font=theme.body_font,
        )
    _text(
        slide,
        str(number),
        SLIDE_W - MARGIN - 0.6,
        FOOTER_Y,
        0.6,
        0.28,
        pt=SOURCE_PT,
        color=theme.color("muted"),
        font=theme.body_font,
        align=PP_ALIGN.RIGHT,
    )


def _source_line(slide, text: str | None, theme: Theme):
    if not text:
        return
    _text(
        slide,
        f"Source : {text}",
        MARGIN,
        FOOTER_Y - 0.42,
        USABLE_W,
        0.3,
        pt=SOURCE_PT,
        color=theme.color("muted"),
        font=theme.body_font,
    )


# --------------------------------------------------------------------------
# Per-slide-type helpers
# --------------------------------------------------------------------------


def _render_title(prs, s: TitleSlide, theme: Theme, deck: Deck, number: int) -> list[str]:
    slide = _blank_slide(prs, theme)
    y = 2.6
    _text(
        slide,
        s.title,
        MARGIN,
        y,
        USABLE_W - 2,
        1.6,
        pt=COVER_TITLE_PT,
        color=theme.color("primary"),
        font=theme.heading_font,
        bold=True,
        line_spacing=1.1,
    )
    y += _estimate_height(s.title, USABLE_W - 2, COVER_TITLE_PT) + 0.25
    if s.subtitle:
        _text(
            slide, s.subtitle, MARGIN, y, USABLE_W - 2, 0.5,
            pt=16, color=theme.color("text"), font=theme.body_font,
        )
        y += 0.55
    if s.reference:
        _text(
            slide, s.reference, MARGIN, y, USABLE_W - 2, 0.4,
            pt=SMALL_PT, color=theme.color("muted"), font=theme.body_font,
        )
    _text(
        slide, deck.entity, MARGIN, FOOTER_Y - 0.1, USABLE_W, 0.35,
        pt=SMALL_PT, color=theme.color("muted"), font=theme.body_font,
    )
    if theme.logo and theme.logo.exists():
        slide.shapes.add_picture(
            str(theme.logo), Inches(SLIDE_W - MARGIN - 1.6), Inches(MARGIN), height=Inches(0.55)
        )
    return []


def _render_section(prs, s: SectionSlide, theme: Theme, deck: Deck, number: int) -> list[str]:
    slide = _blank_slide(prs, theme, background="section_background")
    _text(
        slide, f"{s.number:02d}", MARGIN, 2.7, 2.0, 1.0,
        pt=40, color=theme.color("background"), font=theme.heading_font, bold=True,
    )
    _text(
        slide, s.title, MARGIN, 3.65, USABLE_W - 1.5, 1.2,
        pt=28, color=theme.color("background"), font=theme.heading_font, bold=True,
    )
    return []


def _render_bullets(prs, s: BulletsSlide, theme: Theme, deck: Deck, number: int) -> list[str]:
    slide = _blank_slide(prs, theme)
    warnings = _action_title(slide, s.action_title, theme)
    _footer(slide, theme, number)
    _source_line(slide, s.source, theme)

    x_bullet, x_text = MARGIN, MARGIN + 0.32
    w = USABLE_W - 0.32
    y = CONTENT_Y
    for point in s.points:
        h = _estimate_height(point, w, BODY_PT)
        _rect(slide, x_bullet, y + 0.09, 0.1, 0.1, theme.color("primary"))
        _text(
            slide, point, x_text, y, w, h + 0.1,
            pt=BODY_PT, color=theme.color("text"), font=theme.body_font,
        )
        y += h + 0.30
    if y > CONTENT_Y + CONTENT_H:
        warnings.append(
            f"Content too dense on « {s.action_title[:50]}… » : "
            f"{len(s.points)} points, split the slide."
        )
    return warnings


def _render_reference(prs, s: ReferenceSlide, theme: Theme, deck: Deck, number: int) -> list[str]:
    slide = _blank_slide(prs, theme)
    warnings = _action_title(slide, s.action_title, theme)
    _footer(slide, theme, number)

    col_w = (USABLE_W - 0.6) / 2
    x_right = MARGIN + col_w + 0.6

    # Left column: the regulatory text, in an indented block.
    block_h = min(CONTENT_H, _estimate_height(s.excerpt, col_w - 0.6, SMALL_PT) + 1.25)
    _rect(slide, MARGIN, CONTENT_Y, col_w, block_h, theme.color("block_background"))
    _text(
        slide, s.reference, MARGIN + 0.3, CONTENT_Y + 0.28, col_w - 0.6, 0.5,
        pt=SMALL_PT, color=theme.color("primary"), font=theme.body_font, bold=True,
    )
    _text(
        slide, f"« {s.excerpt} »", MARGIN + 0.3, CONTENT_Y + 0.85,
        col_w - 0.6, block_h - 1.1,
        pt=SMALL_PT, color=theme.color("text"), font=theme.body_font,
    )
    if block_h >= CONTENT_H:
        warnings.append("Regulatory excerpt too long for the left column.")

    # Right column: the interpretation.
    _text(
        slide, "Lecture", x_right, CONTENT_Y, col_w, 0.35,
        pt=SMALL_PT, color=theme.color("muted"), font=theme.body_font, bold=True,
    )
    y = CONTENT_Y + 0.5
    for point in s.interpretation:
        h = _estimate_height(point, col_w - 0.32, BODY_PT)
        _rect(slide, x_right, y + 0.09, 0.1, 0.1, theme.color("primary"))
        _text(
            slide, point, x_right + 0.32, y, col_w - 0.32, h + 0.1,
            pt=BODY_PT, color=theme.color("text"), font=theme.body_font,
        )
        y += h + 0.30
    return warnings


def _render_impacts(prs, s: ImpactsSlide, theme: Theme, deck: Deck, number: int) -> list[str]:
    slide = _blank_slide(prs, theme)
    warnings = _action_title(slide, s.action_title, theme)
    _footer(slide, theme, number)
    _source_line(slide, s.source, theme)

    headers = ["Exigence", "Impact", "Entité", "Criticité"]
    widths = [3.2, 5.2, 1.9, 1.5]
    n_rows = len(s.rows) + 1
    row_h = 0.52
    table = slide.shapes.add_table(
        n_rows, 4, Inches(MARGIN), Inches(CONTENT_Y),
        Inches(sum(widths)), Inches(row_h * n_rows),
    ).table
    table.first_row = True
    for i, w in enumerate(widths):
        table.columns[i].width = Inches(w)

    for j, header in enumerate(headers):
        cell = table.cell(0, j)
        cell.text = header
        cell.fill.solid()
        cell.fill.fore_color.rgb = theme.color("primary")
        _style_cell(cell, theme, SMALL_PT, theme.color("background"), bold=True)

    for i, row in enumerate(s.rows, start=1):
        values = [row.requirement, row.impact, row.entity, row.criticality]
        for j, value in enumerate(values):
            cell = table.cell(i, j)
            cell.text = value
            cell.fill.solid()
            cell.fill.fore_color.rgb = theme.color("background") if i % 2 else theme.color("block_background")
            color = (
                theme.criticality.get(row.criticality, theme.color("text"))
                if j == 3
                else theme.color("text")
            )
            _style_cell(cell, theme, SMALL_PT, color, bold=(j == 3))
            if _estimate_height(value, widths[j] - 0.2, SMALL_PT) > row_h - 0.12:
                warnings.append(
                    f"Cell too long (row {i}, column « {headers[j]} »)."
                )
    return warnings


def _style_cell(cell, theme: Theme, pt: int, color: RGBColor, bold=False):
    cell.margin_left = cell.margin_right = Inches(0.1)
    cell.margin_top = cell.margin_bottom = Inches(0.06)
    cell.vertical_anchor = MSO_ANCHOR.MIDDLE
    for p in cell.text_frame.paragraphs:
        for run in p.runs:
            run.font.name = theme.body_font
            run.font.size = Pt(pt)
            run.font.bold = bold
            run.font.color.rgb = color


def _render_kpi(prs, s: KpiSlide, theme: Theme, deck: Deck, number: int) -> list[str]:
    slide = _blank_slide(prs, theme)
    warnings = _action_title(slide, s.action_title, theme)
    _footer(slide, theme, number)
    _source_line(slide, s.source, theme)

    n = len(s.kpis)
    gap = 0.5
    w = (USABLE_W - gap * (n - 1)) / n
    y = CONTENT_Y + 0.7
    for i, kpi in enumerate(s.kpis):
        x = MARGIN + i * (w + gap)
        _text(
            slide, kpi.value, x, y, w, 1.1,
            pt=48, color=theme.color("accent"), font=theme.heading_font, bold=True,
        )
        _text(
            slide, kpi.label, x, y + 1.2, w, 0.9,
            pt=BODY_PT, color=theme.color("text"), font=theme.body_font,
        )
        if len(kpi.value) > 8:
            warnings.append(f"Key figure too long for the column: « {kpi.value} ».")
    return warnings


def _render_timeline(prs, s: TimelineSlide, theme: Theme, deck: Deck, number: int) -> list[str]:
    slide = _blank_slide(prs, theme)
    warnings = _action_title(slide, s.action_title, theme)
    _footer(slide, theme, number)

    n = len(s.milestones)
    axis_y = CONTENT_Y + 1.35
    w = USABLE_W / n
    _rect(slide, MARGIN, axis_y, USABLE_W, 0.02, theme.color("rule"))
    for i, milestone in enumerate(s.milestones):
        x = MARGIN + i * w
        cx = x + w / 2
        _rect(slide, cx - 0.09, axis_y - 0.08, 0.18, 0.18, theme.color("primary"), MSO_SHAPE.OVAL)
        _text(
            slide, milestone.date, x, axis_y - 0.75, w, 0.4,
            pt=BODY_PT, color=theme.color("primary"), font=theme.heading_font,
            bold=True, align=PP_ALIGN.CENTER,
        )
        _text(
            slide, milestone.label, x + 0.15, axis_y + 0.35, w - 0.3, 1.4,
            pt=SMALL_PT, color=theme.color("text"), font=theme.body_font,
            align=PP_ALIGN.CENTER,
        )
        if _estimate_height(milestone.label, w - 0.3, SMALL_PT) > 1.4:
            warnings.append(f"Milestone label too long: « {milestone.label[:40]}… ».")
    return warnings


def _render_summary(prs, s: SummarySlide, theme: Theme, deck: Deck, number: int) -> list[str]:
    slide = _blank_slide(prs, theme)
    warnings = _action_title(slide, s.action_title, theme)
    _footer(slide, theme, number)

    y = CONTENT_Y
    for i, reco in enumerate(s.recommendations, start=1):
        detail_h = _estimate_height(reco.detail, USABLE_W - 0.85, BODY_PT)
        _text(
            slide, str(i), MARGIN, y - 0.05, 0.6, 0.6,
            pt=24, color=theme.color("accent"), font=theme.heading_font, bold=True,
        )
        _text(
            slide, reco.title, MARGIN + 0.65, y, USABLE_W - 0.85, 0.35,
            pt=BODY_PT + 1, color=theme.color("primary"),
            font=theme.heading_font, bold=True,
        )
        _text(
            slide, reco.detail, MARGIN + 0.65, y + 0.38, USABLE_W - 0.85, detail_h + 0.1,
            pt=BODY_PT, color=theme.color("text"), font=theme.body_font,
        )
        y += 0.38 + detail_h + 0.45
    if y > CONTENT_Y + CONTENT_H:
        warnings.append("Summary slide too dense: reduce to 3 recommendations.")
    return warnings


_HELPERS = {
    "title": _render_title,
    "section": _render_section,
    "bullets": _render_bullets,
    "reference": _render_reference,
    "impacts": _render_impacts,
    "kpi": _render_kpi,
    "timeline": _render_timeline,
    "summary": _render_summary,
}


def render_deck(deck: Deck, theme: Theme, output_path: str | Path) -> RenderResult:
    """Renders the deck to .pptx. Deterministic: same inputs, same file."""
    prs = Presentation()
    prs.slide_width = Inches(SLIDE_W)
    prs.slide_height = Inches(SLIDE_H)

    warnings: list[str] = []
    for number, s in enumerate(deck.slides, start=1):
        warnings += _HELPERS[s.type](prs, s, theme, deck, number)

    path = Path(output_path)
    path.parent.mkdir(parents=True, exist_ok=True)
    prs.save(str(path))
    return RenderResult(path=path, warnings=warnings)
