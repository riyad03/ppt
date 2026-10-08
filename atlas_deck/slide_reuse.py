"""Reuse a template's own slides, stretched to fit our content.

Measuring colours and fonts (`template_probe.py`) only carries over a
template's palette. In a designed template the look *is* the artwork — a
chalkboard frame, a potted plant, hand-drawn circles round the numbers — and
none of that survives being redrawn from scratch.

So here we take their slide as it stands, artwork and all, and change only
what has to change: the text, and how many times the repeating part repeats.
A row of three circles becomes five when we have five things to say.
"""

from __future__ import annotations

import copy
from collections import Counter
from dataclasses import dataclass, field

from pptx.enum.text import MSO_ANCHOR
from pptx.opc.constants import RELATIONSHIP_TYPE as RT
from pptx.util import Emu, Inches, Pt

_R_EMBED = "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}embed"

# An ordinary 16:9 deck is 13.33 inches wide. Templates are authored at any
# size — this one is 20 inches — and a point size that reads well on 13 inches
# is half the apparent size on 20. Anything we assume rather than measure is
# scaled by the canvas it will be seen at.
_REFERENCE_WIDTH = 13.333


def _inches(value) -> float:
    return Emu(value).inches if value is not None else 0.0


def _default_body_pt(slide_w: float) -> float:
    return 14.0 * max(1.0, slide_w / _REFERENCE_WIDTH)


@dataclass
class Column:
    """One instance of a repeating element: every shape in its vertical band."""

    shapes: list = field(default_factory=list)

    @property
    def left(self) -> float:
        return min(_inches(s.left) for s in self.shapes)

    @property
    def right(self) -> float:
        return max(_inches(s.left) + _inches(s.width) for s in self.shapes)

    @property
    def center(self) -> float:
        return (self.left + self.right) / 2

    def texts(self) -> list:
        """Text shapes, top to bottom — the order a reader follows."""
        return sorted(
            (s for s in self.shapes if s.has_text_frame and s.text_frame.text.strip()),
            key=lambda s: _inches(s.top),
        )


def clone_slide(dest_prs, src_slide):
    """Copy a slide wholesale into another deck, artwork intact.

    python-pptx has no API for this, so the shape tree is deep-copied and the
    image relationships re-pointed at the destination — without that, every
    picture silently disappears.
    """
    layout = dest_prs.slide_layouts[6]  # blank
    new_slide = dest_prs.slides.add_slide(layout)
    for shape in src_slide.shapes:
        new_slide.shapes._spTree.append(copy.deepcopy(shape._element))
    _relink_images(new_slide, src_slide)
    return new_slide


def _relink_images(new_slide, src_slide) -> int:
    relinked = 0
    for blip in new_slide.shapes._spTree.iter():
        if not blip.tag.endswith("}blip"):
            continue
        rid = blip.get(_R_EMBED)
        if rid and rid in src_slide.part.rels:
            image_part = src_slide.part.rels[rid].target_part
            blip.set(_R_EMBED, new_slide.part.relate_to(image_part, RT.IMAGE))
            relinked += 1
    return relinked


def detect_columns(slide, slide_w: float, slide_h: float) -> list[Column]:
    """Find the repeating element in a slide, as vertical bands of shapes.

    Backdrop artwork and full-width headings are excluded first: they belong
    to the slide, not to one instance of the repeat.
    """
    candidates = []
    for shape in slide.shapes:
        if shape.left is None or shape.width is None:
            continue
        width = _inches(shape.width)
        area = width * _inches(shape.height)
        # Skip the backdrop and anything spanning most of the slide — a title,
        # a banner, the chalkboard frame itself.
        if area > 0.35 * slide_w * slide_h or width > 0.55 * slide_w:
            continue
        candidates.append(shape)

    if len(candidates) < 4:
        return []

    # Group by horizontal position: a new band starts wherever there is a big
    # horizontal gap between one shape's centre and the next.
    candidates.sort(key=lambda s: _inches(s.left) + _inches(s.width) / 2)
    bands: list[Column] = [Column([candidates[0]])]
    threshold = 0.08 * slide_w
    previous = _inches(candidates[0].left) + _inches(candidates[0].width) / 2
    for shape in candidates[1:]:
        center = _inches(shape.left) + _inches(shape.width) / 2
        # Measured against the previous shape, not the band's running centre:
        # that centre drifts right as shapes join, so the gap to the next
        # column keeps shrinking and two columns merge into one.
        if center - previous > threshold:
            bands.append(Column([shape]))
        else:
            bands[-1].shapes.append(shape)
        previous = center

    # A real repeat has several bands carrying the same number of texts. The
    # signature is the text count, not the total shape count: a column that
    # also happens to sit under a decorative squiggle has one shape more than
    # its neighbours and would otherwise be discarded as not matching — which
    # silently halves the columns we find.
    bands = [b for b in bands if b.texts()]
    if len(bands) < 2:
        return []
    sizes = [len(b.texts()) for b in bands]
    modal = max(set(sizes), key=sizes.count)
    uniform = [b for b in bands if len(b.texts()) == modal]
    if len(uniform) < 2 or not _is_repeat(uniform, slide_h):
        return []
    # A row whose every box holds nothing but a number is a counter, not a
    # content element: the chapter strip « 01 02 03 04 05 » that a template
    # prints across the top of each slide to show where you are. It is a
    # flawless repeat by every structural test, so it was being picked as the
    # element to fill, and a paragraph was then squeezed into a box built for
    # two digits. A column that also carries real text — a numbered heading
    # above a description — is unaffected, since not all of its slots are
    # numeric.
    if all(
        shape.text_frame.text.strip().isdigit()
        for band in uniform
        for shape in band.texts()
    ):
        return []
    return uniform


def _is_repeat(bands: list[Column], slide_h: float) -> bool:
    """Whether these bands are really one element repeated, or just loose text.

    Two text boxes at opposite corners of a slide satisfy every count-based
    test — same number of texts, far enough apart to be separate bands — so a
    closing slide's heading and its small footnote were being read as a
    two-column element. Content was then written for a design with nowhere to
    put it, and the slide came out all artwork and two stray lines.

    What distinguishes a real repeat is alignment: its instances sit at the
    same height and are about as wide as each other. Heights are deliberately
    not compared — designers routinely leave the boxes of a genuine row at
    whatever height their own dummy text needed.
    """
    for index in range(len(bands[0].texts())):
        row = [band.texts()[index] for band in bands]
        tops = [_inches(shape.top) for shape in row]
        widths = [_inches(shape.width) for shape in row]
        if max(tops) - min(tops) > 0.10 * slide_h:
            return False
        if max(widths) > 2.0 * max(0.01, min(widths)):
            return False
    return True


def fit_columns(new_slide, columns: list[Column], wanted: int, slide_w: float) -> list[Column]:
    """Make the repeating element repeat `wanted` times.

    Fewer needed than the template offers: drop the extras. More needed: clone
    one and re-space the row so it still fits the slide. The template's design
    stops being a fixed mould and becomes elastic.
    """
    if not columns or wanted < 1:
        return columns

    columns = sorted(columns, key=lambda c: c.center)
    span_left = columns[0].left
    span_right = columns[-1].right
    original_pitch = (
        (columns[-1].center - columns[0].center) / (len(columns) - 1)
        if len(columns) > 1
        else span_right - span_left
    )

    pitch = (span_right - span_left) / max(wanted, 1)
    # Columns are spaced by their centres, not their left edges: the shapes in
    # a column have different widths, so aligning edges leaves the row looking
    # unevenly spaced even when the maths is right.
    first_center = span_left + pitch / 2

    # Remove the surplus.
    while len(columns) > wanted:
        for shape in columns.pop().shapes:
            shape._element.getparent().remove(shape._element)

    # Clone the last one until there are enough.
    template_col = columns[-1]
    while len(columns) < wanted:
        clone = Column([])
        for shape in template_col.shapes:
            element = copy.deepcopy(shape._element)
            new_slide.shapes._spTree.append(element)
            clone.shapes.append(new_slide.shapes[-1])
        columns.append(clone)

    # Squeezing more columns into the same row means each one has less space.
    # Moving them without narrowing them is what makes the text of one column
    # run into the next — and using fewer than the template has leaves gaps
    # between them instead, so the row widens as readily as it narrows. Growth
    # is bounded: a single column stretched across a row built for four stops
    # looking like the design it came from.
    scale = min(1.6, pitch / original_pitch) if original_pitch else 1.0

    for index, column in enumerate(columns):
        target = first_center + index * pitch
        source_center = column.center
        for shape in column.shapes:
            width = _inches(shape.width)
            offset = (_inches(shape.left) + width / 2) - source_center
            new_width = width * scale
            # Each shape keeps its place within the column, scaled down with it.
            shape.left = Inches(target + offset * scale - new_width / 2)
            shape.width = Inches(new_width)
    return columns


def set_text(shape, text: str) -> None:
    """Replace a shape's text, keeping its formatting.

    The first run carries the font, size and colour the designer chose, so it
    is reused and the rest discarded — writing to `.text` directly would throw
    that styling away.
    """
    frame = shape.text_frame
    paragraph = frame.paragraphs[0]
    runs = paragraph.runs
    if not runs:
        paragraph.add_run()
        runs = paragraph.runs
    runs[0].text = text
    for extra in runs[1:]:
        extra._r.getparent().remove(extra._r)
    for extra_paragraph in frame.paragraphs[1:]:
        extra_paragraph._p.getparent().remove(extra_paragraph._p)


@dataclass
class SlideDesign:
    """One of the template's slides, described by how it is built."""

    index: int
    kind: str               # cover | divider | columns | text
    columns: int = 0
    slots: int = 0          # text shapes inside one column
    text_shapes: int = 0
    images: int = 0
    label: str = ""         # the design's own title, for reporting
    title_chars: int = 0    # how much text its title box holds
    slot_chars: tuple[int, ...] = ()  # and each slot of one repeated column
    # Every box on the slide we could write into, title aside, with how much
    # each one holds. This is what makes a design usable: whether its boxes
    # are a repeated row or two stacked paragraphs is a question about how to
    # fill them, not about whether they can be filled at all.
    box_chars: tuple[int, ...] = ()
    # What the slide's pictures depict, as classified from the images
    # themselves. Empty when no classifier is available, which puts every
    # decision back on geometry alone.
    artwork: tuple[str, ...] = ()

    def describe(self) -> str:
        shape = f"{self.columns}x{self.slots}" if self.kind == "columns" else self.kind
        return f"slide {self.index + 1} ({shape}) « {self.label[:28]} »"


# How sure the classifier has to be before its answer is allowed to change a
# layout decision. Measured against the templates to hand: the readings that
# matter land at 95-100%, while the mistakes sit in the 40s.
_SURE_ENOUGH = 0.60


def _box_key(shape) -> tuple:
    return (
        round(_inches(shape.left), 1),
        round(_inches(shape.top), 1),
        round(_inches(shape.width), 1),
    )


def furniture_of(prs) -> set[tuple]:
    """`furniture`, computed once per template and kept on the presentation."""
    cached = getattr(prs, "_atlas_furniture", None)
    if cached is None:
        cached = furniture(prs)
        prs._atlas_furniture = cached
    return cached


def furniture(prs) -> set[tuple]:
    """Boxes that sit in the same place on many slides: the template's chrome.

    Page numbers, a footer, and the chapter strip « 01 02 03 04 05 » that a
    report prints across the top of every slide to show where you are. They
    look exactly like content to anything examining one slide on its own —
    which is how a KPI ended up written into a navigation marker — and only
    become recognisable across the deck, by never moving.

    The threshold scales with the template so a short one is not judged by a
    long one's standards, and stays at three so that a handful of section
    dividers sharing a number's position are not mistaken for furniture.
    """
    seen: Counter = Counter()
    for slide in prs.slides:
        for shape in slide.shapes:
            if shape.has_text_frame and shape.text_frame.text.strip():
                seen[_box_key(shape)] += 1
    threshold = max(3, len(prs.slides) // 3)
    return {key for key, count in seen.items() if count >= threshold}


def catalog(prs, slide_w: float, slide_h: float) -> list[SlideDesign]:
    """Describe every slide in a template by its structure.

    Nothing here relies on names or layouts — only on what the slide is made
    of — because a template may carry none of those.
    """
    designs: list[SlideDesign] = []
    body_pt = _default_body_pt(slide_w)
    chrome = furniture_of(prs)
    try:
        from .template_vision import label_pictures

        labels = label_pictures(prs)
    except Exception:
        # Classifying artwork is an enhancement, never a dependency: without it
        # the catalogue is exactly what it was before, judged on geometry.
        labels = {}
    for index, slide in enumerate(prs.slides):
        texts = [s for s in slide.shapes if s.has_text_frame and s.text_frame.text.strip()]
        images = [s for s in slide.shapes if s.shape_type == 13]
        columns = detect_columns(slide, slide_w, slide_h)
        slots = len(columns[0].texts()) if columns else 0

        numeric = any(t.text_frame.text.strip().isdigit() for t in texts)
        if len(columns) >= 2 and slots:
            kind = "columns"
        elif index == 0:
            kind = "cover"
        elif numeric and len(texts) <= 4:
            kind = "divider"
        else:
            kind = "text"

        label = max((t.text_frame.text.strip() for t in texts), key=len, default="")
        title_shape = find_title(slide, slide_h)
        title_el = title_shape._element if title_shape is not None else None
        boxes = [
            _budget(s, body_pt)
            for s in sorted(texts, key=lambda s: (_inches(s.top), _inches(s.left)))
            if s._element is not title_el and _box_key(s) not in chrome
        ]
        designs.append(
            SlideDesign(
                index, kind, len(columns), slots, len(texts), len(images),
                label.replace("\n", " ")[:40],
                _budget(title_shape, body_pt * 1.7) if title_shape is not None else 0,
                tuple(_budget(s, body_pt) for s in columns[0].texts()) if columns else (),
                tuple(boxes),
                # Only confident readings are recorded. A label carries real
                # weight in the scoring, and a hesitant guess between a chart
                # and a chalk drawing was enough to disqualify half a
                # template's usable designs. Below this the picture is simply
                # left unclassified and geometry decides, as it did before.
                tuple(
                    labels[s.shape_id].kind
                    for s in images
                    if s.shape_id in labels
                    and labels[s.shape_id].kind
                    and labels[s.shape_id].confidence >= _SURE_ENOUGH
                ),
            )
        )
    return designs


def coverage(design: SlideDesign) -> float:
    """What fraction of a design's text boxes our content can reach.

    Kept for reporting. It no longer decides whether a design may be used:
    that question is answered by how well the content fills the boxes, which
    `_fit_cost` measures directly.
    """
    if not design.text_shapes:
        return 0.0
    return min(1.0, (1 + len(design.box_chars)) / design.text_shapes)


def candidates_for(designs: list[SlideDesign], slide_type: str) -> list[SlideDesign]:
    """Every design that could carry this kind of slide.

    A cover and a section divider are answered by their role: one carries a
    deck title, the other a number, and neither is interchangeable with a
    content slide. Everything else is open. A design used to qualify only if
    it had a repeating row of columns deep enough for the slide type, which
    quietly disqualified every template built from single-column layouts — a
    chart-and-caption report offered fourteen perfectly fillable designs and
    was told it had none.

    What remains is a set of boxes with measured capacities, so eligibility is
    just whether there is a box to write in. Which of them is worth using is
    then a matter of degree, scored by `_fit_cost` and bounded by `_MAX_COST`.
    """
    if slide_type == "title":
        return [d for d in designs if d.kind == "cover"]
    if slide_type == "section":
        return [d for d in designs if d.kind == "divider"]
    if slide_type not in _TYPICAL_CHARS:
        return []
    return [d for d in designs if d.kind not in ("cover", "divider") and d.box_chars]


# A heading in a regulatory note needs room; a design whose title box holds a
# dozen characters cannot carry one, however well its boxes match.
_WANTED_TITLE_CHARS = 45


# Roughly how much text each of our slide types has to offer per item, in
# characters. A KPI's caption is two words however much room it is given; a
# bullet's explanation is a sentence. This says nothing about any particular
# template — it describes our own content — and it is what lets a design be
# rejected for being too roomy as well as for being too tight.
_TYPICAL_CHARS = {
    "summary": (2, 25, 90),
    "kpi": (6, 22),
    "timeline": (8, 40),
    "bullets": (25, 120),
}

# What a content slide is expected to carry. Fixed on purpose: it is the
# yardstick every design is measured against, so a design cannot make itself
# look like a good fit simply by having little to fill.
_TARGET_ITEMS = 3

# Past this the design would be left too empty, or its boxes too overrun, to be
# worth reusing; the slide is drawn from scratch instead, which is never
# half-filled. A design that matches its content loosely scores 2 to 4, one
# that cannot hold it at all runs into the teens.
_MAX_COST = 7.0


def items_wanted(design: SlideDesign, slide_type: str) -> int:
    """How many items this design has room for, before any text exists.

    A repeated row wants one item per column — that is what the repeat is for.
    Any other design wants as many items as its boxes can be divided into,
    since each item occupies one box per piece of text it carries.
    """
    per_item = max(1, len(_TYPICAL_CHARS.get(slide_type, (1,))))
    if design.columns >= 2 and design.slots:
        return design.columns
    # Never fewer than two. This number is handed to the model as the number of
    # items to write, and every content schema requires at least two — so a
    # design with a single pair of boxes asked for one point, got one, and the
    # deck died on a validation error it could never retry its way out of.
    # A surplus item is a scoring cost, not a crash: it goes to a spare box.
    return max(2, len(design.box_chars) // per_item)


def _fit_cost(design: SlideDesign, slide_type: str, item_count: int | None = None) -> float:
    """How poorly a design suits the content — lower is better.

    Both sides are reduced to the same thing: a list of character counts. The
    design offers boxes of a given capacity; the content offers pieces of text
    of a given length. Sorting each and pairing them off says precisely how the
    slide will come out, without needing to know whether the boxes form a row
    of five circles or two stacked paragraphs.

    Three ways it can go wrong, and each is charged for: a box with nothing to
    put in it leaves a hole, a piece of text with no box to go in is lost, and a
    box far bigger or far smaller than its text is filled badly.
    """
    # A cover and a divider are picked for their role — one carries the deck's
    # title, the other its section number — and `candidates_for` has already
    # limited the field to designs that play that role. There is no content to
    # measure against, so only the heading's room counts. Scoring them like
    # content slides rejected every one of them outright.
    if slide_type in ("title", "section"):
        return max(0, _WANTED_TITLE_CHARS - design.title_chars) / 8

    typical = _TYPICAL_CHARS.get(slide_type)
    if not typical or not design.box_chars:
        return float("inf")
    # Scored against what a content slide ought to carry, never against what
    # the design happens to have room for. Sizing the content to the design
    # made small designs cheap — a single quote box scored better than a row
    # of six matched boxes, because having less to fill meant less that could
    # go wrong — and the deck collapsed onto its most threadbare layouts.
    count = max(1, item_count if item_count is not None else _TARGET_ITEMS)

    supply = sorted((size for _ in range(count) for size in typical), reverse=True)
    if design.columns >= 2 and design.slots:
        # A repeated row is elastic: it is cloned or trimmed to the number of
        # items before anything is written, so what it will offer is one
        # column's slots repeated, not the row as the designer left it.
        demand = sorted((cap for _ in range(count) for cap in design.slot_chars), reverse=True)
    else:
        demand = sorted(design.box_chars, reverse=True)

    cost = max(0, _WANTED_TITLE_CHARS - design.title_chars) / 8
    # A slide built around a chart is making a claim with someone else's data.
    # Reusing it means cloning that chart and writing our own words around it,
    # so a note about a banking circular would carry a graph of a template's
    # invented quarterly revenue and appear to source its argument from it.
    # The boxes may fit perfectly; the slide is still wrong, and no measurement
    # of capacity can see that. Priced above the ceiling so it takes something
    # exceptional elsewhere to bring such a design back into range.
    cost += 5.0 * sum(1 for kind in design.artwork if kind == "chart")
    for position in range(max(len(supply), len(demand))):
        text = supply[position] if position < len(supply) else 0
        box = demand[position] if position < len(demand) else 0
        if not box:
            # Content with nowhere to go. Mild: the surplus is dropped or spills
            # into a spare box, which is untidy rather than visibly broken.
            cost += 1.0
        elif not text:
            # A box we cannot fill. This is the one that shows, so it costs
            # most — an empty box in the middle of a design reads as a mistake.
            cost += 2.0
        else:
            ratio = box / max(1, text)
            if ratio > 2.5:
                cost += min(3.0, ratio / 2.5)
            elif ratio < 0.4:
                cost += min(3.0, 0.4 / ratio)
    return cost


def pick_design(designs: list[SlideDesign], slide_type: str, item_count: int) -> SlideDesign | None:
    """The single best design for one slide, ignoring what else is in the deck."""
    options = candidates_for(designs, slide_type)
    if not options:
        return None
    best = min(options, key=lambda d: _fit_cost(d, slide_type, item_count))
    return best if _fit_cost(best, slide_type, item_count) <= _MAX_COST else None


def allocate_by_type(
    designs: list[SlideDesign], types: list[str], counts: list[int | None]
) -> list[SlideDesign | None]:
    """Allocate designs knowing only each slide's type, before its content exists.

    Content has to be written to fit the element it will land in, which means
    choosing the element first — so this runs before any text is generated,
    when the item count is still unknown.
    """
    used: dict[int, int] = {}
    allocation: list[SlideDesign | None] = []
    for slide_type, count in zip(types, counts):
        options = candidates_for(designs, slide_type)
        if not options:
            allocation.append(None)
            continue
        # Reusing a design is a blemish; filling one badly is a broken slide.
        # So the ceiling is applied first, and the taste for variety only
        # chooses among designs that already fit — otherwise an unused but
        # unsuitable design outscores a good one and is then thrown out, losing
        # the slide to a blank redraw.
        affordable = [d for d in options if _fit_cost(d, slide_type, count) <= _MAX_COST]
        if not affordable:
            allocation.append(None)
            continue
        best = min(
            affordable,
            key=lambda d: _fit_cost(d, slide_type, count) + 2.5 * used.get(d.index, 0),
        )
        used[best.index] = used.get(best.index, 0) + 1
        allocation.append(best)
    return allocation


def budget_note(caps: dict) -> str:
    """The room a design gives, phrased as a target rather than a ceiling.

    Given only a maximum the model writes far under it, and the slide comes
    out three short lines on a wide empty board. A range tells it to use the
    space: fill the box, without overflowing it.
    """
    parts = []
    title = caps.get("title") or 0
    if title:
        parts.append(f"le titre : entre {max(12, int(title * 0.6))} et {title} caractères")
    slots = caps.get("slots") or []
    if slots:
        listed = ", puis ".join(
            f"entre {max(2, int(n * 0.6))} et {n}" for n in slots
        )
        parts.append(f"pour chaque élément, les textes font {listed} caractères")
    if not parts:
        return ""
    return (
        "Place disponible dans la maquette du client. Vise le haut de ces "
        "fourchettes : un texte trop court laisse la slide vide, un texte trop "
        "long ne tient pas. " + " ; ".join(parts) + "."
    )


def allocate_designs(designs: list[SlideDesign], deck_slides: list) -> list[SlideDesign | None]:
    """Spread the deck across the template's designs instead of reusing one.

    Picking the best-fitting design per slide independently lands every
    column-based slide on the same template slide, so a deck built from a
    23-slide template ends up showing three of its designs. Each reuse is
    penalised, so a slightly worse but unused design wins — the deck then
    looks as varied as the template it came from.
    """
    used: dict[int, int] = {}
    allocation: list[SlideDesign | None] = []
    for deck_slide in deck_slides:
        options = candidates_for(designs, deck_slide.type)
        if not options:
            allocation.append(None)
            continue
        items = len(items_for(deck_slide))
        affordable = [
            d for d in options if _fit_cost(d, deck_slide.type, items) <= _MAX_COST
        ]
        if not affordable:
            allocation.append(None)
            continue
        best = min(
            affordable,
            key=lambda d: _fit_cost(d, deck_slide.type, items) + 2.5 * used.get(d.index, 0),
        )
        used[best.index] = used.get(best.index, 0) + 1
        allocation.append(best)
    return allocation


def _capacity(width: float, height: float, pt: float) -> int:
    per_line = max(1, int(width * 72 / (pt * 0.55)))
    lines = max(1, int(height / (pt * 1.28 / 72)))
    return per_line * lines


def _budget(shape, default_pt: float) -> int:
    """How many characters a box holds.

    The designer sized it for their own words, so whatever they put in it
    demonstrably fits — a better floor than our deliberately pessimistic
    estimate, which reads a box holding "You can describe the agenda here" as
    good for 15 characters.
    """
    sizes = [
        run.font.size.pt
        for para in shape.text_frame.paragraphs
        for run in para.runs
        if run.font.size
    ]
    computed = _capacity(
        _inches(shape.width), _inches(shape.height), max(sizes) if sizes else default_pt
    )
    return max(computed, len(shape.text_frame.text.strip()))


def fit_text(shape, text: str, default_pt: float = 14.0) -> str:
    """Resize the type so the text fills the template's box — both ways.

    Their boxes are sized for their own dummy text and don't reflow, and our
    sentences are never the same length as theirs. Shrinking alone was only
    half the job: a caption dropped into a box built for a paragraph stayed at
    the designer's size and left most of the box bare, which is what makes a
    slide read as empty. So the type grows to take up the space as well as
    shrinking to fit it, and what is left over is centred rather than parked
    against the top edge.

    Growth is capped just above the designer's own size so a body box never
    outgrows the heading above it, and the shrink floor is high enough to stay
    readable — text small enough to disappear is not a fit, it is a defect.
    """
    width, height = _inches(shape.width), _inches(shape.height)
    if not width or not height:
        return text

    sizes = [
        run.font.size.pt
        for para in shape.text_frame.paragraphs
        for run in para.runs
        if run.font.size
    ]
    original = max(sizes) if sizes else default_pt

    ceiling = original * 1.35
    # Losing words is worse than losing points of size, so the floor stays low
    # enough that a heading is scaled down rather than cut off mid-phrase. The
    # absolute guard is what keeps it readable: a fraction of the designer's
    # size means nothing on its own, since a template authored on a 20-inch
    # canvas sets body text where a 13-inch one sets headings.
    floor = max(default_pt * 0.55, original * 0.35)

    pt = original
    while len(text) <= _capacity(width, height, pt + 1) and pt < ceiling:
        pt += 1
    while len(text) > _capacity(width, height, pt) and pt > floor:
        pt -= 1
    if round(pt) != round(original):
        for para in shape.text_frame.paragraphs:
            for run in para.runs:
                run.font.size = Pt(round(pt))

    if len(text) > _capacity(width, height, pt) * 0.9:
        shape.text_frame.word_wrap = True

    limit = _capacity(width, height, pt)
    if len(text) <= limit:
        return text
    # Still too long: end on a whole word at a sentence-like break, with no
    # trailing "…". A clipped word reads as a bug; a shorter sentence does not.
    clipped = text[:limit]
    for stop in (". ", " ; ", ", "):
        cut = clipped.rfind(stop)
        if cut > limit * 0.5:
            return clipped[:cut].rstrip(" ,;")
    cut = clipped.rfind(" ")
    return (clipped[:cut] if cut > limit * 0.4 else clipped).rstrip(" ,;")


def capacities(src_prs, design: SlideDesign, slide_w: float, slide_h: float) -> dict:
    """How much text each part of a design can actually hold, in characters.

    Measured from the real boxes and the designer's own type sizes, so the
    content can be written to fit instead of being written blind and then
    truncated into « Incident… ».
    """
    # A design without a repeated row still has boxes to write to, and the
    # content still has to be sized for them. Reporting only the row's slots
    # left every single-column design with no guidance at all, so the model
    # wrote to its own instincts and the text fitted the boxes by accident.
    slots = list(design.slot_chars)
    if not slots and design.box_chars:
        per_item = max(1, items_wanted(design, "bullets"))
        slots = sorted(design.box_chars, reverse=True)[
            : max(1, len(design.box_chars) // per_item)
        ]
    return {
        "title": design.title_chars,
        "slots": slots,
        "columns": design.columns or 1,
    }


def _assign_slots(slots: list, values: tuple, default_pt: float) -> list:
    """Decide which slot each piece of text goes into.

    Reading order is the designer's intent, so it is where this starts. But a
    column commonly pairs a one-line label box with a paragraph box five inches
    tall, and zipping in order drops a whole sentence into the label — where it
    is shrunk until it cannot be read — while the paragraph box is blanked.
    That single mismatch produced both complaints about the output: microscopic
    text along the top, and an empty board underneath.

    So a value that plainly does not fit its slot is swapped into one that has
    room, including a slot that would otherwise have been left empty. Only a
    decisive improvement counts: shuffling two values that overflow by similar
    amounts would scramble the designer's order to no purpose.
    """
    caps = [_budget(shape, default_pt) for shape in slots]
    assigned: list = [None] * len(slots)
    for index, value in enumerate(values[: len(slots)]):
        assigned[index] = value

    def overflow(state: list) -> int:
        return sum(
            max(0, len(v) - caps[i]) for i, v in enumerate(state) if v is not None
        )

    for _ in range(len(slots)):
        cost = overflow(assigned)
        if cost == 0:
            break
        best = None
        for i in range(len(slots)):
            for j in range(i + 1, len(slots)):
                trial = list(assigned)
                trial[i], trial[j] = trial[j], trial[i]
                trial_cost = overflow(trial)
                if trial_cost < cost * 0.6:
                    best, cost = trial, trial_cost
        if best is None:
            break
        assigned = best
    return assigned


def _level_slot_sizes(per_column: list[list]) -> None:
    """Give the same slot the same type size in every column.

    Each box is sized on its own to fit its own text, so a row of three columns
    ends up at three different sizes — one line large, its neighbours tiny. A
    row has to read as a row: every column takes the smallest size any of them
    needed.

    Slots are passed in as they stand in the template, before anything is
    written. Reading them back off the finished slide lined up the wrong boxes:
    a blanked slot drops out of the column's text shapes, so every slot below
    it shifted up a place and was levelled against its neighbour's neighbour.
    """
    if len(per_column) < 2:
        return
    for index in range(max(len(column) for column in per_column)):
        shapes = [column[index] for column in per_column if index < len(column)]
        sizes = [
            run.font.size.pt
            for shape in shapes
            if shape.text_frame.text.strip()
            for para in shape.text_frame.paragraphs
            for run in para.runs
            if run.font.size
        ]
        if not sizes:
            continue
        smallest = Pt(round(min(sizes)))
        for shape in shapes:
            for para in shape.text_frame.paragraphs:
                for run in para.runs:
                    run.font.size = smallest


def _center_sparse_text(per_column: list[list]) -> None:
    """Centre text that leaves most of its box empty — a whole row at a time.

    Judged once the row has been levelled to a common size, never while each
    box is still being fitted. Levelling lowers the size afterwards, so text
    that looked as though it filled its box at fitting time ends up occupying a
    fraction of it — a stripe along the top edge with dead space beneath, which
    is what makes an otherwise correct slide look abandoned.

    The decision is taken per slot across every column, not per box. Deciding
    box by box centred one column and left its neighbours at the top, so the
    paragraphs of a single row started at three different heights — worse than
    the gap it was meant to close. A row is only centred when every column in
    it is sparse, which keeps it aligned either way.
    """
    for index in range(max((len(column) for column in per_column), default=0)):
        row = [column[index] for column in per_column if index < len(column)]
        fills = []
        for shape in row:
            text = shape.text_frame.text.strip()
            width, height = _inches(shape.width), _inches(shape.height)
            sizes = [
                run.font.size.pt
                for para in shape.text_frame.paragraphs
                for run in para.runs
                if run.font.size
            ]
            if not text or not width or not height or not sizes:
                fills = []
                break
            pt = max(sizes)
            if height / (pt * 1.28 / 72) <= 2.5:
                fills = []
                break
            fills.append(len(text) / max(1, _capacity(width, height, pt)))
        if fills and max(fills) < 0.5:
            for shape in row:
                shape.text_frame.vertical_anchor = MSO_ANCHOR.MIDDLE


def _same_text(left: str, right: str) -> bool:
    return left.strip().lower().rstrip(" .:") == right.strip().lower().rstrip(" .:")


def items_for(deck_slide) -> list[tuple[str, ...]]:
    """Our content, as one tuple of strings per repeated column."""
    kind = deck_slide.type
    if kind == "summary":
        return [
            (f"{i:02d}", r.title, r.detail)
            for i, r in enumerate(deck_slide.recommendations, start=1)
        ]
    if kind == "kpi":
        return [(k.value, k.label) for k in deck_slide.kpis]
    if kind == "timeline":
        return [(m.date, m.label) for m in deck_slide.milestones]
    if kind == "bullets":
        # Where a design gives a column a lead line above its description, a
        # point written as "intitulé : explication" fills both.
        #
        # Only a colon counts. Splitting on a comma as well used to produce
        # halves that do not stand on their own — "Le seuil est abaissé" above
        # "élargit mécaniquement l'assiette", with the connective dropped and
        # the second box opening mid-sentence. A whole sentence in one box
        # reads better than a broken one across two, so anything that does not
        # divide cleanly stays intact and is placed by capacity.
        split = []
        for point in deck_slide.points:
            head, sep, tail = point.partition(" : ")
            # A lead of four characters is still a lead: « Impact », « Coût »,
            # « Délai » are exactly the captions this is looking for. Requiring
            # twelve rejected them, so the colon stayed inline in the paragraph
            # and the caption box above it was left blank.
            if sep and 4 <= len(head) <= 70 and len(tail) > 12:
                split.append((head.rstrip(" :,"), tail.strip()))
            else:
                split.append((point,))
        return split
    return []


def heading_for(deck_slide) -> str:
    return getattr(deck_slide, "action_title", None) or getattr(deck_slide, "title", "") or ""


def build_from_design(dest_prs, src_prs, design, deck_slide, slide_w: float, slide_h: float):
    """Reproduce one of our slides on top of the template's own design."""
    new_slide = clone_slide(dest_prs, src_prs.slides[design.index])
    default_pt = _default_body_pt(slide_w)
    # The template's own chrome — page number, footer, chapter strip — is left
    # exactly as the designer placed it. It is not ours to fill or to clear.
    chrome = furniture_of(src_prs)

    title_shape = find_title(new_slide, slide_h)
    if title_shape is not None:
        set_text(title_shape, fit_text(title_shape, heading_for(deck_slide), default_pt))

    # Compared by XML element, not by identity: python-pptx hands back a new
    # wrapper object on every iteration, so `is not` would never match and the
    # title we just wrote would be overwritten below.
    title_el = title_shape._element if title_shape is not None else None
    others = [
        s for s in new_slide.shapes
        if s.has_text_frame
        and s.text_frame.text.strip()
        and s._element is not title_el
        and _box_key(s) not in chrome
    ]

    if deck_slide.type == "section":
        # Their divider carries a big number; ours knows its own.
        for shape in others:
            if shape.text_frame.text.strip().isdigit():
                set_text(shape, f"{deck_slide.number:02d}")
        return new_slide

    if deck_slide.type == "title":
        extras = [t for t in (deck_slide.subtitle, deck_slide.reference) if t]
        ordered = sorted(others, key=lambda s: _inches(s.top))
        for shape, value in zip(ordered, extras):
            set_text(shape, fit_text(shape, value, default_pt))
        for shape in ordered[len(extras):]:
            set_text(shape, "")
        return new_slide

    written = {title_el} if title_el is not None else set()

    items = items_for(deck_slide)
    # A lead that restates the slide's own heading prints the same sentence
    # twice — once large across the top, once small directly beneath it.
    heading = heading_for(deck_slide)
    items = [
        tuple(v for v in values if not _same_text(v, heading)) or values
        for values in items
    ]
    used = 0
    columns = (
        fit_columns(new_slide, detect_columns(new_slide, slide_w, slide_h), len(items), slide_w)
        if items
        else []
    )

    if items and not columns:
        # A design with no repeating row is filled as one set of boxes. The
        # placement rule is the same one used inside a column — match each
        # piece of text to a box that can hold it — applied to the whole slide
        # instead. Without this branch such a design was chosen and then left
        # blank but for its heading, which is why templates built from
        # single-column layouts could not be used at all.
        boxes = sorted(others, key=lambda s: (_inches(s.top), _inches(s.left)))
        values = tuple(
            piece for values in items for piece in values if piece and piece != "—"
        )
        assigned = _assign_slots(boxes, values, default_pt)
        for position, shape in enumerate(boxes):
            if assigned[position] is None and shape.text_frame.text.strip().isdigit():
                assigned[position] = f"{position + 1:02d}"
        for shape, value in zip(boxes, assigned):
            set_text(shape, fit_text(shape, value, default_pt) if value else "")
            written.add(shape._element)
        _center_sparse_text([[s] for s in boxes])
        used = len(items)

    per_column: list[list] = []
    if columns:
        for column, values in zip(columns, items):
            # Read while the template's own text is still in place: a slot we
            # blank stops counting as a text shape, so reading the slots back
            # afterwards returns a shorter, misaligned list.
            slots = column.texts()
            per_column.append(slots)
            assigned = _assign_slots(slots, values, default_pt)
            # A slot the designer filled with a bare number is a counter for
            # the element it sits in, and the element's position is something
            # we always know, whatever the content turned out to be. Blanking
            # it leaves the badge of every column empty — a hole the design was
            # never meant to have — so it is numbered instead.
            for position, shape in enumerate(slots):
                if assigned[position] is None and shape.text_frame.text.strip().isdigit():
                    assigned[position] = f"{used + 1:02d}"
            for shape, value in zip(slots, assigned):
                # Their dummy text in a slot we have nothing for would
                # otherwise stay on the slide.
                set_text(shape, fit_text(shape, value, default_pt) if value else "")
                written.add(shape._element)
            used += 1
        _level_slot_sizes(per_column)
        _center_sparse_text(per_column)

    # Boxes we haven't written to still hold the template's placeholder copy
    # ("Elaborate on what you want to discuss.", "Category 1"). Blanking them
    # leaves a hole in the design, so anything we still have to say goes in
    # first — spare item texts, then the source, then the entity — and only
    # what's left over is cleared.
    spare = [" — ".join(p for p in pair if p and p != "—") for pair in items[used:]]
    spare += [t for t in (getattr(deck_slide, "source", None),) if t]
    leftovers = [
        s
        for s in sorted(new_slide.shapes, key=lambda s: (_inches(s.top), _inches(s.left)))
        if s.has_text_frame
        and s.text_frame.text.strip()
        and s._element not in written
        and _box_key(s) not in chrome
    ]
    for shape in leftovers:
        set_text(shape, fit_text(shape, spare.pop(0), default_pt) if spare else "")
    return new_slide


def find_title(slide, slide_h: float) -> object | None:
    """The title is the widest, highest, biggest-type text on the slide."""
    texts = [
        s for s in slide.shapes
        if s.has_text_frame and s.text_frame.text.strip() and s.top is not None
    ]
    # A bare number is the design's section marker, not its title — and it is
    # often set in the largest type on the slide, so it wins on size alone.
    texts = [s for s in texts if not s.text_frame.text.strip().isdigit()] or texts
    if not texts:
        return None

    def weight(shape):
        sizes = [
            run.font.size.pt
            for para in shape.text_frame.paragraphs
            for run in para.runs
            if run.font.size
        ]
        return (max(sizes) if sizes else 0, _inches(shape.width))

    # Sitting high on the slide is a hint, not a requirement. Insisting on it
    # threw away the title of any cover that sets its name across the lower
    # half — a report whose title sat at 7 inches down an 11-inch slide had
    # both of its large title boxes excluded, and a 20pt corner label was
    # promoted to heading while the real ones were blanked. So the upper half
    # only decides between texts of the same size; the biggest type on the
    # slide wins wherever it happens to sit.
    upper = [s for s in texts if _inches(s.top) < 0.55 * slide_h]
    if upper and weight(max(upper, key=weight)) >= weight(max(texts, key=weight)):
        return max(upper, key=weight)
    return max(texts, key=weight)
