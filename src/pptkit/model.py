"""Addressable view of a .pptx file.

A deck is a list of slides; a slide is a list of *blocks*. A block is one
addressable piece of text: a text frame, one table cell, a SmartArt node or
the slide notes. Every block carries

* ``id``   - stable address, derived from OOXML shape ids, e.g. ``b6``,
             ``b12/b5`` (inside a group), ``b9.r2c3`` (table cell),
             ``notes``. Survives text edits: it names a *place*, not a string.
* ``key``  - legacy key from the original MCP server: the text itself, or
             ``{text}_{index}_{shape_id}`` when the same text appears more
             than once on the slide. Kept so existing prompts keep working.
"""

from __future__ import annotations

from dataclasses import dataclass, field
from typing import Any, Optional

EMU_PER_INCH = 914400


def emu_to_in(value: Optional[int]) -> Optional[float]:
    if value is None:
        return None
    try:
        return round(int(value) / EMU_PER_INCH, 3)
    except (TypeError, ValueError):
        return None


def inches(value: Optional[float]) -> Optional[int]:
    if value is None:
        return None
    return int(round(float(value) * EMU_PER_INCH))


@dataclass
class Block:
    id: str
    kind: str  # text | cell | notes | diagram | chart | picture
    text: str
    shape_id: Optional[int] = None
    index: Optional[int] = None  # walk ordinal, legacy keys only
    name: str = ""
    shape_type: str = ""
    is_placeholder: bool = False
    placeholder_type: str = ""
    left: Optional[int] = None
    top: Optional[int] = None
    width: Optional[int] = None
    height: Optional[int] = None
    row: Optional[int] = None
    col: Optional[int] = None
    editable: bool = True
    key: str = ""
    # live python-pptx object used to write back; never serialized
    target: Any = field(default=None, repr=False, compare=False)

    def geometry(self) -> dict:
        return {
            "left": emu_to_in(self.left),
            "top": emu_to_in(self.top),
            "width": emu_to_in(self.width),
            "height": emu_to_in(self.height),
        }

    def to_dict(self) -> dict:
        out = {
            "id": self.id,
            "key": self.key,
            "kind": self.kind,
            "text": self.text,
            "shape_id": self.shape_id,
            "name": self.name,
            "shape_type": self.shape_type,
            "is_placeholder": self.is_placeholder,
            "editable": self.editable,
            **self.geometry(),
        }
        if self.placeholder_type:
            out["placeholder_type"] = self.placeholder_type
        if self.row is not None:
            out["row"] = self.row
            out["col"] = self.col
        return out


@dataclass
class Slide:
    index: int  # 1-based
    layout: str
    blocks: list[Block] = field(default_factory=list)
    notes: str = ""
    shapes: list[dict] = field(default_factory=list)

    def text_blocks(self) -> list[Block]:
        """Blocks that hold text an agent can read or rewrite (pictures excluded)."""
        return [b for b in self.blocks if b.kind != "picture" and b.text.strip()]

    def to_dict(self, *, include_empty: bool = False, include_shapes: bool = False) -> dict:
        if include_empty:
            blocks = [b for b in self.blocks if b.kind != "picture"]
        else:
            blocks = self.text_blocks()
        out = {
            "index": self.index,
            "layout": self.layout,
            "notes": self.notes,
            "blocks": [b.to_dict() for b in blocks],
        }
        if include_shapes:
            out["shapes"] = self.shapes
        return out


@dataclass
class Deck:
    path: str
    prs: Any
    slides: list[Slide] = field(default_factory=list)
    warnings: list[str] = field(default_factory=list)

    @property
    def size(self) -> dict:
        return {
            "width": emu_to_in(self.prs.slide_width),
            "height": emu_to_in(self.prs.slide_height),
            "unit": "in",
        }

    def summary(self) -> dict:
        return {
            "file": self.path,
            "slide_count": len(self.prs.slides),
            "slide_size": self.size,
        }

    def slide(self, index: int) -> Slide:
        from .errors import NotFoundError

        for s in self.slides:
            if s.index == index:
                return s
        raise NotFoundError(
            f"slide {index} does not exist (the deck has {len(self.slides)})",
            detail={"slide_count": len(self.slides)},
        )

    def all_blocks(self):
        for s in self.slides:
            for b in s.blocks:
                yield s, b
