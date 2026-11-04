"""Write text back into a deck.

The one rule that matters here: **never drop run formatting**. Naively doing
``shape.text_frame.text = new`` collapses every run into one and the slide
loses bold words, colored words and inline emphasis. Instead we write the new
text into the *first* run and delete the rest, and when the new text has more
lines than the shape has paragraphs we clone the last paragraph (keeping its
bullet/numbering/indent) rather than inventing a bare one.

Edits are planned first, applied second: a batch either lands whole or not at
all, so an agent never has to reason about half-applied state.
"""

from __future__ import annotations

import copy
import re
from dataclasses import dataclass
from typing import Optional

from pptx.text.text import _Paragraph

from .errors import NotFoundError
from .model import Block, Deck, Slide

A_P = "{http://schemas.openxmlformats.org/drawingml/2006/main}p"
LEGACY_KEY = re.compile(
    r"^(?P<text>.*)_(?P<index>\d+)_(?P<shape_id>\d+)"
    r"(?:_r(?P<row>\d+)c(?P<col>\d+))?(?:_(?P<ordinal>\d+))?$",
    re.S,
)


@dataclass
class Change:
    slide: int
    ref: str
    block_id: str
    kind: str
    old: str
    new: str

    def to_dict(self) -> dict:
        return {
            "slide": self.slide,
            "id": self.block_id,
            "ref": self.ref,
            "kind": self.kind,
            "old": self.old,
            "new": self.new,
        }


# --------------------------------------------------------------------------
# text frame primitives
# --------------------------------------------------------------------------


def _set_paragraph_text(paragraph: _Paragraph, text: str) -> None:
    runs = list(paragraph.runs)
    if not runs:
        paragraph.text = text
        return
    runs[0].text = text
    for run in runs[1:]:
        run._r.getparent().remove(run._r)


def set_text_frame_text(text_frame, new_text: str) -> None:
    """Replace a text frame's text, keeping paragraph and run formatting."""
    if text_frame is None:
        raise ValueError("no text frame")
    tx_body = text_frame._txBody
    elements = tx_body.findall(A_P)
    lines = new_text.replace("\r\n", "\n").replace("\r", "\n").split("\n")
    if not elements:
        text_frame.text = new_text
        return

    while len(elements) < len(lines):
        clone = copy.deepcopy(elements[-1])
        tx_body.append(clone)
        elements.append(clone)

    for element, line in zip(elements, lines):
        _set_paragraph_text(_Paragraph(element, text_frame), line)

    for element in elements[len(lines):]:
        tx_body.remove(element)


def block_text_frame(block: Block):
    if block.kind == "notes":
        return block.target.notes_slide.notes_text_frame
    return block.target.text_frame


def set_block_text(block: Block, new_text: str) -> None:
    set_text_frame_text(block_text_frame(block), new_text)


# --------------------------------------------------------------------------
# reference resolution
# --------------------------------------------------------------------------


def resolve(slide: Slide, ref: str) -> list[Block]:
    """Address/id -> legacy key -> legacy ``text_index_shapeid`` -> exact text.

    An exact text hit returns *every* occurrence on that slide, which is what
    the original server did; the reply lists each block id so nothing is
    silently multi-edited.
    """
    ref = ref.strip()
    for block in slide.blocks:
        if block.id == ref:
            return [block]
    for block in slide.blocks:
        if block.key and block.key == ref:
            return [block]

    match = LEGACY_KEY.match(ref)
    if match:
        text = match.group("text")
        index = int(match.group("index"))
        shape_id = int(match.group("shape_id"))
        row = int(match.group("row")) if match.group("row") else None
        col = int(match.group("col")) if match.group("col") else None

        def at_place(b: Block) -> bool:
            if b.index != index or b.shape_id != shape_id:
                return False
            if row is not None and (b.row, b.col) != (row, col):
                return False
            return True

        exact = [b for b in slide.blocks if b.text == text and at_place(b)]
        if exact:
            return exact
        # position beats wording: the key is an address, so a stale key from
        # before a text edit still points at the block it was made for
        same_place = [b for b in slide.blocks if at_place(b)]
        if same_place:
            return same_place
        loose = [b for b in slide.blocks if b.text == text]
        if loose:
            return loose

    by_text = [b for b in slide.blocks if b.text == ref]
    if by_text:
        return by_text

    raise NotFoundError(
        f"slide {slide.index} has no block matching {_clip(ref)}",
        detail={
            "slide": slide.index,
            "hint": "run `ppt read --json` to list ids, or `ppt find` to locate by keyword",
            "candidates": [
                {"id": b.id, "text": _clip(b.text, 60)}
                for b in slide.blocks
                if b.text and _loose_match(b.text, ref)
            ][:10],
        },
    )


def _clip(value: str, limit: int = 40) -> str:
    value = str(value).replace("\n", "\\n")
    return value if len(value) <= limit else value[: limit - 1] + "…"


def _loose_match(text: str, ref: str) -> bool:
    head = ref.strip()[:20]
    return bool(head) and head in text


# --------------------------------------------------------------------------
# batch: plan, then apply
# --------------------------------------------------------------------------


@dataclass
class PlanItem:
    slide: int
    ref: str
    block: Block
    new: str


def plan_edits(
    deck: Deck,
    edits: dict,
    *,
    slides: Optional[list[int]] = None,
) -> tuple[list[PlanItem], list[dict]]:
    only = set(slides) if slides else None
    plan: list[PlanItem] = []
    errors: list[dict] = []

    for slide_key, pairs in edits.items():
        try:
            slide_no = int(str(slide_key).lstrip("sS"))
        except ValueError:
            errors.append({"slide": slide_key, "error": "slide number must be an integer"})
            continue
        if only is not None and slide_no not in only:
            continue
        if not isinstance(pairs, dict):
            errors.append({"slide": slide_no, "error": "expected {address: new text}"})
            continue
        try:
            slide = deck.slide(slide_no)
        except NotFoundError as e:
            errors.append({"slide": slide_no, "error": e.message})
            continue

        for ref, new_text in pairs.items():
            if new_text is None:
                new_text = ""
            if new_text == "spaceholder":  # legacy no-op marker
                continue
            try:
                blocks = resolve(slide, str(ref))
            except NotFoundError as e:
                errors.append(
                    {"slide": slide_no, "ref": str(ref), "error": e.message, "detail": e.detail}
                )
                continue
            for block in blocks:
                if not block.editable:
                    errors.append(
                        {
                            "slide": slide_no,
                            "ref": str(ref),
                            "id": block.id,
                            "error": f"{block.kind} blocks cannot be edited directly (edit SmartArt/charts in PowerPoint)",
                        }
                    )
                    continue
                plan.append(PlanItem(slide=slide_no, ref=str(ref), block=block, new=str(new_text)))

    if errors:
        # a batch lands whole or not at all — never hand back a partial plan
        return [], errors
    return plan, errors


def apply_plan(plan: list[PlanItem]) -> list[Change]:
    changes = []
    for item in plan:
        block = item.block
        old = block.text
        set_block_text(block, item.new)
        block.text = item.new
        changes.append(
            Change(
                slide=item.slide,
                ref=item.ref,
                block_id=block.id,
                kind=block.kind,
                old=old,
                new=item.new,
            )
        )
    return changes
