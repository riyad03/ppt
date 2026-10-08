"""Deterministic rendering engine. No LLM call here.

One helper per slide type, a single grid, all colors and fonts sourced from
`Theme`. Each helper returns the list of layout warnings (estimated
overflow), later surfaced in the graph's state.
"""

from __future__ import annotations

import io
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
from .slide_reuse import allocate_designs, build_from_design, catalog
from .theme import Theme

# The grid — every position, size and type size — comes from the `Theme`, so a
# style measured off a client's template drives the layout without touching
# this file. `Theme`'s defaults are the house grid (16:9, 0.75" margins).


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

    # Some designs carry their look in artwork rather than a fill colour — a
    # chalkboard, a texture, a photographic cover. Where the template had a
    # full-slide image, it goes down first and everything else draws on top.
    if theme.background_image and background == "background":
        slide.shapes.add_picture(
            io.BytesIO(theme.background_image), 0, 0,
            width=Inches(theme.slide_w), height=Inches(theme.slide_h),
        )
        # The reused artwork is whatever covered the template's slide — often
        # the patterned surround rather than the dark panel their text sat on.
        # Our text colours were measured against that panel, so writing
        # straight onto the artwork can leave white on pale grey. A panel in
        # the background colour restores the contrast those colours assume.
        _rect(
            slide,
            theme.margin / 2,
            theme.title_y - 0.25,
            theme.slide_w - theme.margin,
            theme.footer_y - theme.title_y + 0.55,
            theme.color("background"),
        )

    # A logo measured off a client template was found repeating across their
    # slides, so it repeats across ours, in the same spot.
    if theme.logo_box:
        _place_logo(slide, theme)
    return slide


def _action_title(slide, text: str, theme: Theme) -> list[str]:
    _text(
        slide,
        text,
        theme.margin,
        theme.title_y,
        theme.usable_w,
        theme.title_h,
        pt=theme.title_pt,
        color=theme.color("primary"),
        font=theme.heading_font,
        bold=True,
        anchor=MSO_ANCHOR.TOP,
    )
    if _estimate_height(text, theme.usable_w, theme.title_pt) > theme.title_h:
        return [f"Title too long, overflows into the content area: « {text[:60]}… »"]
    return []


def _footer(slide, theme: Theme, number: int):
    if theme.footer:
        _text(
            slide,
            theme.footer,
            theme.margin,
            theme.footer_y,
            theme.usable_w - 1,
            0.28,
            pt=theme.source_pt,
            color=theme.color("muted"),
            font=theme.body_font,
        )
    _text(
        slide,
        str(number),
        theme.slide_w - theme.margin - 0.6,
        theme.footer_y,
        0.6,
        0.28,
        pt=theme.source_pt,
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
        theme.margin,
        theme.footer_y - 0.42,
        theme.usable_w,
        0.3,
        pt=theme.source_pt,
        color=theme.color("muted"),
        font=theme.body_font,
    )


# --------------------------------------------------------------------------
# Per-slide-type helpers
# --------------------------------------------------------------------------


def _render_title(prs, s: TitleSlide, theme: Theme, deck: Deck, number: int) -> list[str]:
    slide = _blank_slide(prs, theme)
    y = theme.slide_h * 0.35
    _text(
        slide,
        s.title,
        theme.margin,
        y,
        theme.usable_w - 2,
        1.6,
        pt=theme.cover_title_pt,
        color=theme.color("primary"),
        font=theme.heading_font,
        bold=True,
        line_spacing=1.1,
    )
    y += _estimate_height(s.title, theme.usable_w - 2, theme.cover_title_pt) + 0.25
    if s.subtitle:
        _text(
            slide, s.subtitle, theme.margin, y, theme.usable_w - 2, 0.5,
            pt=int(theme.body_pt * 1.15), color=theme.color("text"), font=theme.body_font,
        )
        y += 0.55
    if s.reference:
        _text(
            slide, s.reference, theme.margin, y, theme.usable_w - 2, 0.4,
            pt=theme.small_pt, color=theme.color("muted"), font=theme.body_font,
        )
    _text(
        slide, deck.entity, theme.margin, theme.footer_y - 0.1, theme.usable_w, 0.35,
        pt=theme.small_pt, color=theme.color("muted"), font=theme.body_font,
    )
    if not theme.logo_box:  # otherwise every slide already carries it
        _place_logo(slide, theme)
    return []


def _place_logo(slide, theme: Theme) -> None:
    """The logo comes either from a YAML theme (a path on disk) or from a
    measured template (the raw image bytes, plus where it sat on their
    slides — reused so it lands where their logo belongs)."""
    if not theme.logo:
        return
    if isinstance(theme.logo, bytes):
        source = io.BytesIO(theme.logo)
    elif theme.logo.exists():
        source = str(theme.logo)
    else:
        return

    if theme.logo_box:
        left, top, width, _ = theme.logo_box
        slide.shapes.add_picture(source, Inches(left), Inches(top), width=Inches(width))
    else:
        slide.shapes.add_picture(
            source, Inches(theme.slide_w - theme.margin - 1.6), Inches(theme.margin), height=Inches(0.55)
        )


def _render_section(prs, s: SectionSlide, theme: Theme, deck: Deck, number: int) -> list[str]:
    slide = _blank_slide(prs, theme, background="section_background")
    _text(
        slide, f"{s.number:02d}", theme.margin, theme.slide_h * 0.36, 2.0, 1.0,
        pt=int(theme.cover_title_pt * 1.15), color=theme.color("background"),
        font=theme.heading_font, bold=True,
    )
    _text(
        slide, s.title, theme.margin, theme.slide_h * 0.49, theme.usable_w - 1.5, 1.2,
        pt=int(theme.title_pt * 1.3), color=theme.color("background"),
        font=theme.heading_font, bold=True,
    )
    return []


def _render_bullets(prs, s: BulletsSlide, theme: Theme, deck: Deck, number: int) -> list[str]:
    slide = _blank_slide(prs, theme)
    warnings = _action_title(slide, s.action_title, theme)
    _footer(slide, theme, number)
    _source_line(slide, s.source, theme)

    x_bullet, x_text = theme.margin, theme.margin + 0.32
    w = theme.usable_w - 0.32
    y = theme.content_y
    for point in s.points:
        h = _estimate_height(point, w, theme.body_pt)
        _rect(slide, x_bullet, y + 0.09, 0.1, 0.1, theme.color("primary"))
        _text(
            slide, point, x_text, y, w, h + 0.1,
            pt=theme.body_pt, color=theme.color("text"), font=theme.body_font,
        )
        y += h + 0.30
    if y > theme.content_y + theme.content_h:
        warnings.append(
            f"Content too dense on « {s.action_title[:50]}… » : "
            f"{len(s.points)} points, split the slide."
        )
    return warnings


def _render_reference(prs, s: ReferenceSlide, theme: Theme, deck: Deck, number: int) -> list[str]:
    slide = _blank_slide(prs, theme)
    warnings = _action_title(slide, s.action_title, theme)
    _footer(slide, theme, number)

    col_w = (theme.usable_w - 0.6) / 2
    x_right = theme.margin + col_w + 0.6

    # Left column: the regulatory text, in an indented block.
    block_h = min(theme.content_h, _estimate_height(s.excerpt, col_w - 0.6, theme.small_pt) + 1.25)
    _rect(slide, theme.margin, theme.content_y, col_w, block_h, theme.color("block_background"))
    _text(
        slide, s.reference, theme.margin + 0.3, theme.content_y + 0.28, col_w - 0.6, 0.5,
        pt=theme.small_pt, color=theme.color("primary"), font=theme.body_font, bold=True,
    )
    _text(
        slide, f"« {s.excerpt} »", theme.margin + 0.3, theme.content_y + 0.85,
        col_w - 0.6, block_h - 1.1,
        pt=theme.small_pt, color=theme.color("text"), font=theme.body_font,
    )
    if block_h >= theme.content_h:
        warnings.append("Regulatory excerpt too long for the left column.")

    # Right column: the interpretation.
    _text(
        slide, "Lecture", x_right, theme.content_y, col_w, 0.35,
        pt=theme.small_pt, color=theme.color("muted"), font=theme.body_font, bold=True,
    )
    y = theme.content_y + 0.5
    for point in s.interpretation:
        h = _estimate_height(point, col_w - 0.32, theme.body_pt)
        _rect(slide, x_right, y + 0.09, 0.1, 0.1, theme.color("primary"))
        _text(
            slide, point, x_right + 0.32, y, col_w - 0.32, h + 0.1,
            pt=theme.body_pt, color=theme.color("text"), font=theme.body_font,
        )
        y += h + 0.30
    return warnings


def _render_impacts(prs, s: ImpactsSlide, theme: Theme, deck: Deck, number: int) -> list[str]:
    slide = _blank_slide(prs, theme)
    warnings = _action_title(slide, s.action_title, theme)
    _footer(slide, theme, number)
    _source_line(slide, s.source, theme)

    headers = ["Exigence", "Impact", "Entité", "Criticité"]
    # Proportions rather than fixed inches: the table has to fill whatever
    # width the template leaves, and the row has to fit the template's own
    # type size — otherwise every cell overflows on a wider slide or in
    # larger type.
    widths = [ratio * theme.usable_w for ratio in (0.27, 0.44, 0.16, 0.13)]
    n_rows = len(s.rows) + 1
    row_h = max(0.52, 2 * theme.small_pt * 1.28 / 72 + 0.16)
    table = slide.shapes.add_table(
        n_rows, 4, Inches(theme.margin), Inches(theme.content_y),
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
        _style_cell(cell, theme, theme.small_pt, theme.color("background"), bold=True)

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
            _style_cell(cell, theme, theme.small_pt, color, bold=(j == 3))
            if _estimate_height(value, widths[j] - 0.2, theme.small_pt) > row_h - 0.12:
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
    w = (theme.usable_w - gap * (n - 1)) / n
    y = theme.content_y + 0.7
    for i, kpi in enumerate(s.kpis):
        x = theme.margin + i * (w + gap)
        _text(
            slide, kpi.value, x, y, w, 1.1,
            pt=int(theme.cover_title_pt * 1.4), color=theme.color("accent"),
            font=theme.heading_font, bold=True,
        )
        _text(
            slide, kpi.label, x, y + 1.2, w, 0.9,
            pt=theme.body_pt, color=theme.color("text"), font=theme.body_font,
        )
        if len(kpi.value) > 8:
            warnings.append(f"Key figure too long for the column: « {kpi.value} ».")
    return warnings


def _render_timeline(prs, s: TimelineSlide, theme: Theme, deck: Deck, number: int) -> list[str]:
    slide = _blank_slide(prs, theme)
    warnings = _action_title(slide, s.action_title, theme)
    _footer(slide, theme, number)

    n = len(s.milestones)
    axis_y = theme.content_y + 1.35
    w = theme.usable_w / n
    _rect(slide, theme.margin, axis_y, theme.usable_w, 0.02, theme.color("rule"))
    for i, milestone in enumerate(s.milestones):
        x = theme.margin + i * w
        cx = x + w / 2
        _rect(slide, cx - 0.09, axis_y - 0.08, 0.18, 0.18, theme.color("primary"), MSO_SHAPE.OVAL)
        _text(
            slide, milestone.date, x, axis_y - 0.75, w, 0.4,
            pt=theme.body_pt, color=theme.color("primary"), font=theme.heading_font,
            bold=True, align=PP_ALIGN.CENTER,
        )
        _text(
            slide, milestone.label, x + 0.15, axis_y + 0.35, w - 0.3, 1.4,
            pt=theme.small_pt, color=theme.color("text"), font=theme.body_font,
            align=PP_ALIGN.CENTER,
        )
        if _estimate_height(milestone.label, w - 0.3, theme.small_pt) > 1.4:
            warnings.append(f"Milestone label too long: « {milestone.label[:40]}… ».")
    return warnings


def _render_summary(prs, s: SummarySlide, theme: Theme, deck: Deck, number: int) -> list[str]:
    slide = _blank_slide(prs, theme)
    warnings = _action_title(slide, s.action_title, theme)
    _footer(slide, theme, number)

    y = theme.content_y
    for i, reco in enumerate(s.recommendations, start=1):
        detail_h = _estimate_height(reco.detail, theme.usable_w - 0.85, theme.body_pt)
        _text(
            slide, str(i), theme.margin, y - 0.05, 0.6, 0.6,
            pt=int(theme.title_pt * 1.15), color=theme.color("accent"),
            font=theme.heading_font, bold=True,
        )
        _text(
            slide, reco.title, theme.margin + 0.65, y, theme.usable_w - 0.85, 0.35,
            pt=theme.body_pt + 1, color=theme.color("primary"),
            font=theme.heading_font, bold=True,
        )
        _text(
            slide, reco.detail, theme.margin + 0.65, y + 0.38, theme.usable_w - 0.85, detail_h + 0.1,
            pt=theme.body_pt, color=theme.color("text"), font=theme.body_font,
        )
        y += 0.38 + detail_h + 0.45
    if y > theme.content_y + theme.content_h:
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


def render_deck(
    deck: Deck,
    theme: Theme,
    output_path: str | Path,
    template_path: str | Path | None = None,
    design_indices: list[int | None] | None = None,
) -> RenderResult:
    """Renders the deck to .pptx. Deterministic: same inputs, same file.

    With `template_path`, each slide is built on the template's own design
    where a suitable one exists — their artwork kept, the repeating part
    stretched to our item count. Anything the template has no design for (an
    impact table) is drawn from scratch as usual.
    """
    prs = Presentation()
    prs.slide_width = Inches(theme.slide_w)
    prs.slide_height = Inches(theme.slide_h)

    source, designs = None, []
    if template_path:
        source = Presentation(str(template_path))
        designs = catalog(source, theme.slide_w, theme.slide_h)

    warnings: list[str] = []
    if designs and design_indices:
        # Reuse the planner's choices: the text was written to fit these exact
        # boxes, so re-deciding here would put it in differently-sized ones.
        by_index = {d.index: d for d in designs}
        allocation = [by_index.get(i) if i is not None else None for i in design_indices]
    else:
        allocation = allocate_designs(designs, deck.slides) if designs else [None] * len(deck.slides)
    for number, (s, design) in enumerate(zip(deck.slides, allocation), start=1):
        if design is not None:
            build_from_design(prs, source, design, s, theme.slide_w, theme.slide_h)
        else:
            warnings += _HELPERS[s.type](prs, s, theme, deck, number)
    if designs:
        chosen = [d for d in allocation if d is not None]
        warnings.append(
            f"{len(chosen)}/{len(deck.slides)} slides built on the template's designs "
            f"({len({d.index for d in chosen})} different ones); the rest drawn from scratch."
        )

    path = Path(output_path)
    path.parent.mkdir(parents=True, exist_ok=True)
    path = _save(prs, path, warnings)
    return RenderResult(path=path, warnings=warnings)


def _save(prs: Presentation, path: Path, warnings: list[str]) -> Path:
    """Save, working around the file being open in PowerPoint.

    Windows refuses to overwrite a deck someone is viewing. Discarding a
    generation that cost minutes of model time over that would be absurd, so
    a numbered variant is written instead and the caller is told which.
    """
    for attempt in range(20):
        candidate = path if attempt == 0 else path.with_name(f"{path.stem}-{attempt + 1}{path.suffix}")
        try:
            prs.save(str(candidate))
            if attempt:
                warnings.append(
                    f"{path.name} is open in another program; saved as {candidate.name} instead."
                )
            return candidate
        except PermissionError:
            continue
    raise PermissionError(
        f"Could not write {path.name} or any numbered variant — the folder or files are locked."
    )
