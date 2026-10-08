"""Resolution of the visual style: colors, fonts and page geometry.

The visual identity belongs to the platform's client, not to Atlas Guardian:
the renderer only knows logical names (`primary`, `accent`, `muted`) and asks
this object for every position and size. Adding a client = adding a YAML file
or handing over a .pptx, never touching the code.
"""

from __future__ import annotations

from dataclasses import dataclass, replace
from pathlib import Path

import yaml
from pptx.dml.color import RGBColor

from .template_probe import MeasuredStyle

THEMES_DIR = Path(__file__).parent / "themes"


def _rgb(hex_value: str) -> RGBColor:
    return RGBColor.from_string(hex_value.replace("#", "").upper())


def _luminance(color: RGBColor) -> float:
    r, g, b = (c / 255 for c in (color[0], color[1], color[2]))
    channels = [c / 12.92 if c <= 0.03928 else ((c + 0.055) / 1.055) ** 2.4 for c in (r, g, b)]
    return 0.2126 * channels[0] + 0.7152 * channels[1] + 0.0722 * channels[2]


def _contrast(a: RGBColor, b: RGBColor) -> float:
    la, lb = sorted((_luminance(a), _luminance(b)))
    return (lb + 0.05) / (la + 0.05)


@dataclass(frozen=True)
class Theme:
    name: str
    heading_font: str
    body_font: str
    colors: dict[str, RGBColor]
    criticality: dict[str, RGBColor]
    logo: Path | bytes | None
    footer: str | None
    background_image: bytes | None = None

    # Page geometry, in inches. Defaults describe the house grid; a theme
    # measured from a client template overrides them with that template's own.
    slide_w: float = 13.333
    slide_h: float = 7.5
    margin: float = 0.75
    title_y: float = 0.62
    title_h: float = 0.95
    content_y: float = 1.85
    footer_y: float = 6.95
    logo_box: tuple[float, float, float, float] | None = None

    # Type sizes, in points.
    cover_title_pt: int = 34
    title_pt: int = 21
    body_pt: int = 14
    small_pt: int = 11
    source_pt: int = 9

    @property
    def usable_w(self) -> float:
        return self.slide_w - 2 * self.margin

    @property
    def content_h(self) -> float:
        return self.footer_y - 0.25 - self.content_y

    def color(self, name: str) -> RGBColor:
        return self.colors[name]

    @classmethod
    def load(cls, tenant: str = "default", base_dir: Path | None = None) -> "Theme":
        base = base_dir or THEMES_DIR
        path = base / f"{tenant}.yaml"
        if not path.exists():
            path = base / "default.yaml"
        data = yaml.safe_load(path.read_text(encoding="utf-8"))

        logo = data.get("logo")
        return cls(
            name=data["name"],
            heading_font=data["fonts"]["headings"],
            body_font=data["fonts"]["body"],
            colors={k: _rgb(v) for k, v in data["colors"].items()},
            criticality={k: _rgb(v) for k, v in data.get("criticality", {}).items()},
            logo=Path(logo) if logo else None,
            footer=data.get("footer"),
        )

    @classmethod
    def from_template(cls, style: MeasuredStyle, base_dir: Path | None = None) -> "Theme":
        """Build a theme from a client template that was measured.

        Starts from the default theme and overrides only what the measurement
        actually found, so anything a template can't express — the criticality
        colors, which are business semantics — keeps its house value.
        """
        base = cls.load("default", base_dir=base_dir)
        colors = dict(base.colors)
        colors.update({k: _rgb(v) for k, v in style.colors.items() if v})

        # A measured palette that renders text invisible is worse than no
        # measurement at all. This happens for real: a blackboard template
        # writes in white over a dark background *image*, so the text color
        # measures as white while the background reads as the theme's white.
        # Where a background image is reused, the text colors belong to it and
        # are kept; otherwise any role without enough contrast reverts to the
        # house value.
        if not style.background_image:
            for role in ("text", "primary", "accent", "muted"):
                if _contrast(colors[role], colors["background"]) < 2.5:
                    colors[role] = base.colors[role]
            if _contrast(colors["text"], colors["background"]) < 2.5:
                colors["background"] = base.colors["background"]
        elif _luminance(colors["text"]) > 0.5:
            # Light text over a dark background image. The fill underneath is
            # invisible in normal use, but if the image ever fails to draw the
            # text would vanish — so darken the fill to match the artwork.
            colors["background"] = base.colors["primary"]
        # Section dividers invert the palette, so they need the same guarantee.
        if _contrast(colors["background"], colors["section_background"]) < 2.5:
            colors["section_background"] = base.colors["section_background"]

        slide_w = style.slide_width_in
        margin = style.margin_in if style.margin_in is not None else base.margin
        usable_w = slide_w - 2 * margin

        # A template's headings are sized for its own short titles. Ours run to
        # 120 characters, so the measured size is capped to what still fits in
        # about two lines across this template's width — otherwise every title
        # overflows on a narrow slide. The cap is generous enough to leave the
        # house default (21pt on a 16:9 slide) untouched.
        title_pt = int(min(style.heading_pt or base.title_pt, usable_w * 3.27))
        title_pt = max(title_pt, 12)
        line_h = title_pt * 1.28 / 72
        title_h = max(base.title_h, 3 * line_h + 0.16)

        # Content starts below the title block; a template only tells us where
        # its topmost element sits, so the title's own height is added back.
        title_y = style.title_top_in if style.title_top_in is not None else base.title_y
        content_y = title_y + title_h + 0.28

        return replace(
            base,
            name=f"measured from template ({style.slide_width_in:g}x{style.slide_height_in:g}in)",
            heading_font=style.heading_font or base.heading_font,
            body_font=style.body_font or base.body_font,
            colors=colors,
            logo=style.logo or base.logo,
            logo_box=style.logo_box_in,
            background_image=style.background_image,
            footer=style.footer if style.footer is not None else base.footer,
            slide_w=slide_w,
            slide_h=style.slide_height_in,
            margin=margin,
            title_y=title_y,
            title_h=title_h,
            content_y=content_y,
            footer_y=(
                style.content_bottom_in
                if style.content_bottom_in is not None
                else style.slide_height_in - 0.55
            ),
            title_pt=title_pt,
            cover_title_pt=int(title_pt * 1.6),
            body_pt=int(style.body_pt) if style.body_pt else base.body_pt,
            small_pt=int(style.small_pt) if style.small_pt else base.small_pt,
            source_pt=int(style.small_pt) if style.small_pt else base.source_pt,
        )
