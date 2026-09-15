"""Resolution of the visual theme.

The visual identity belongs to the platform's client, not to Atlas Guardian:
the renderer only knows logical names (`primary`, `accent`, `muted`). Adding a
client = adding a YAML file, never touching the code.
"""

from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path

import yaml
from pptx.dml.color import RGBColor

THEMES_DIR = Path(__file__).parent / "themes"


def _rgb(hex_value: str) -> RGBColor:
    return RGBColor.from_string(hex_value.replace("#", "").upper())


@dataclass(frozen=True)
class Theme:
    name: str
    heading_font: str
    body_font: str
    colors: dict[str, RGBColor]
    criticality: dict[str, RGBColor]
    logo: Path | None
    footer: str | None

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
