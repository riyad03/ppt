"""Label a template's artwork by what it depicts, not just where it sits.

Geometry can say that a picture occupies a rectangle. It cannot say whether
that rectangle holds a person, a chart or a decorative squiggle — and a caption
set beneath a photograph of someone reads as describing that person, while the
same caption beneath an abstract shape reads as describing the slide. Placing
text well needs the difference.

CLIP answers it from the image alone: the picture is scored against a fixed
list of candidate descriptions, so what comes back is one of our own labels
with a confidence attached, rather than free prose that would have to be parsed
and retried when it came back malformed. That matters here — most of the
complexity in `graph.py` exists because a local generative model will not
reliably produce structured output, and this avoids repeating it.

The pictures are read straight out of the `.pptx`, so none of this requires the
slide to be rendered — which is just as well, since turning a `.pptx` into an
image needs PowerPoint or LibreOffice and neither runs headless here.

Everything degrades quietly: with no model installed, or none reachable, the
labels come back empty and callers fall back to pure geometry.
"""

from __future__ import annotations

import hashlib
import io
import json
import os
from dataclasses import dataclass, field
from pathlib import Path

# What a picture can be. The wording matters — CLIP compares the image against
# these sentences, so they read as captions someone might write, not as terse
# category names.
_KINDS = {
    "person": "a photo or illustration of one person",
    "people": "a photo or illustration of a group of people",
    "chart": "a chart, graph or data visualisation",
    "icon": "a small simple icon or symbol",
    "logo": "a company logo or wordmark",
    "abstract": "an abstract decorative shape, pattern or background",
    "scene": "a photograph of a place, a building or an object",
}

# Asked only of pictures already labelled as people. The label set is ours to
# choose, which is how a fixed-vocabulary classifier still answers a question
# about composition: which side of the figure should be left clear.
_FACING = {
    "left": "a person facing towards the left",
    "right": "a person facing towards the right",
}

# Below this the classifier is guessing between several readings, and a wrong
# label is worse than none: it would move text for a reason that isn't there.
_MIN_CONFIDENCE = 0.28

_MODEL_NAME = os.getenv("CLIP_MODEL", "openai/clip-vit-base-patch32")
_CACHE_PATH = Path(os.getenv("ARTWORK_CACHE", ".artwork-labels.json"))


@dataclass
class PictureLabel:
    """What one picture is, and where inside its own frame it is painted."""

    kind: str = ""
    confidence: float = 0.0
    facing: str = ""
    # A coarse grid of the image's alpha channel: True where paint is. A
    # cut-out figure fills almost the whole of its bounding box, so the box
    # alone says nothing useful — the silhouette is what shows the empty
    # quarter of the frame beside her that text could occupy.
    painted: tuple[tuple[bool, ...], ...] = field(default_factory=tuple)

    @property
    def is_figure(self) -> bool:
        return self.kind in ("person", "people")

    @property
    def is_content(self) -> bool:
        """Artwork carrying meaning of its own, which our text cannot replace."""
        return self.kind in ("chart", "person", "people", "scene")

    def free_columns(self, threshold: float = 0.12) -> list[int]:
        """Columns of the grid that are essentially unpainted, left to right."""
        if not self.painted:
            return []
        rows = len(self.painted)
        return [
            column
            for column in range(len(self.painted[0]))
            if sum(row[column] for row in self.painted) / rows <= threshold
        ]


def silhouette(blob: bytes, cells: int = 16) -> tuple[tuple[bool, ...], ...]:
    """Where a picture actually paints, as a coarse grid.

    The alpha channel's bounding box is not enough: a cut-out person spans
    nearly the full width and height of her own file, so the box reports the
    whole frame as occupied while most of it is transparent around her.
    """
    try:
        from PIL import Image
    except ImportError:
        return ()
    try:
        image = Image.open(io.BytesIO(blob))
    except Exception:
        return ()
    if image.mode not in ("RGBA", "LA"):
        # No transparency: every pixel is painted.
        return tuple(tuple(True for _ in range(cells)) for _ in range(cells))
    alpha = image.getchannel("A").resize((cells, cells), Image.BOX)
    values = list(alpha.getdata())
    # Shrinking averages the alpha, so a cell only a fifth covered still comes
    # back around 50. Testing against a low value therefore called almost every
    # cell painted and the silhouette reported no free space anywhere. A cell
    # counts as occupied when most of it is.
    return tuple(
        tuple(values[row * cells + column] > 128 for column in range(cells))
        for row in range(cells)
    )


class _Classifier:
    """Loads CLIP once, on first use, and never again."""

    def __init__(self) -> None:
        self._model = None
        self._processor = None
        self._failed = False

    def available(self) -> bool:
        if self._failed:
            return False
        if self._model is not None:
            return True
        try:
            from transformers import CLIPModel, CLIPProcessor

            self._model = CLIPModel.from_pretrained(_MODEL_NAME)
            self._model.eval()
            self._processor = CLIPProcessor.from_pretrained(_MODEL_NAME)
            return True
        except Exception:
            # No transformers, no torch, or no downloaded weights. Callers fall
            # back to geometry, which is how this behaved before CLIP existed.
            self._failed = True
            return False

    def score(self, blob: bytes, choices: dict[str, str]) -> tuple[str, float]:
        if not self.available():
            return "", 0.0
        try:
            import torch
            from PIL import Image

            image = Image.open(io.BytesIO(blob)).convert("RGB")
            keys = list(choices)
            inputs = self._processor(
                text=[choices[k] for k in keys],
                images=image,
                return_tensors="pt",
                padding=True,
            )
            with torch.no_grad():
                logits = self._model(**inputs).logits_per_image[0]
                probabilities = logits.softmax(dim=-1)
            best = int(probabilities.argmax())
            return keys[best], float(probabilities[best])
        except Exception:
            return "", 0.0


_CLASSIFIER = _Classifier()


def _load_cache() -> dict:
    try:
        return json.loads(_CACHE_PATH.read_text(encoding="utf-8"))
    except (OSError, ValueError):
        return {}


def _save_cache(cache: dict) -> None:
    try:
        _CACHE_PATH.write_text(json.dumps(cache, indent=1), encoding="utf-8")
    except OSError:
        pass


def label_picture(blob: bytes, cache: dict | None = None) -> PictureLabel:
    """Classify one image, reusing an earlier answer where there is one.

    Keyed by the image's own bytes, so the same artwork recurring across a
    template's slides — and across runs on that template — is classified once.
    """
    digest = hashlib.sha1(blob).hexdigest()
    store = _load_cache() if cache is None else cache
    remembered = store.get(digest)
    if remembered is None:
        kind, confidence = _CLASSIFIER.score(blob, _KINDS)
        if confidence < _MIN_CONFIDENCE:
            kind, confidence = "", 0.0
        facing = ""
        if kind in ("person", "people"):
            facing, facing_confidence = _CLASSIFIER.score(blob, _FACING)
            if facing_confidence < 0.5:
                facing = ""
        remembered = {"kind": kind, "confidence": confidence, "facing": facing}
        store[digest] = remembered
        if cache is None:
            _save_cache(store)
    return PictureLabel(
        kind=remembered.get("kind", ""),
        confidence=remembered.get("confidence", 0.0),
        facing=remembered.get("facing", ""),
        painted=silhouette(blob),
    )


def label_pictures(prs) -> dict[int, PictureLabel]:
    """Label every picture in a template, keyed by shape id.

    The silhouette is recomputed each time because it is cheap; only the
    classification, which is not, is remembered on disk.
    """
    cache = _load_cache()
    labels: dict[int, PictureLabel] = {}
    for slide in prs.slides:
        for shape in slide.shapes:
            if shape.shape_type != 13:  # picture
                continue
            try:
                blob = shape.image.blob
            except (AttributeError, ValueError):
                continue
            labels[shape.shape_id] = label_picture(blob, cache)
    _save_cache(cache)
    return labels
