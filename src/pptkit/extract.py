"""Read a deck into addressable blocks.

Ports the shape-walking rules of the original MCP ``ppt_class``:

* text frames, including inside groups (recursively)
* every table cell as its own block
* SmartArt nodes, read from the diagram *data* part (where the text lives)
* slide notes
* duplicate texts get the legacy ``{text}_{index}_{shape_id}`` key

Addresses are built from OOXML shape ids so they stay valid after a text edit.
"""

from __future__ import annotations

import re
from typing import Any, Iterator, Optional

from lxml import etree
from pptx import Presentation
from pptx.enum.shapes import MSO_SHAPE_TYPE

from .errors import FileError
from .model import Block, Deck, Slide, emu_to_in

A_NS = "http://schemas.openxmlformats.org/drawingml/2006/main"
DGM_NS = "http://schemas.openxmlformats.org/drawingml/2006/diagram"
R_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"


def _shape_type_name(shape) -> str:
    try:
        st = shape.shape_type
    except (NotImplementedError, ValueError, AttributeError):
        return ""
    if st is None:
        return ""
    return getattr(st, "name", str(st))


def _placeholder_type(shape) -> str:
    if not getattr(shape, "is_placeholder", False):
        return ""
    try:
        pt = shape.placeholder_format.type
    except (ValueError, AttributeError):
        return ""
    return getattr(pt, "name", str(pt or ""))


def _is_diagram(shape) -> bool:
    if _shape_type_name(shape) == "DIAGRAM":
        return True
    el = getattr(shape, "_element", None)
    if el is None or not str(el.tag).endswith("}graphicFrame"):
        return False
    for graphic_data in el.iter(f"{{{A_NS}}}graphicData"):
        if (graphic_data.get("uri") or "").endswith("/diagram"):
            return True
    return False


def _diagram_part(shape):
    """Return the DataModel part holding the SmartArt text, if resolvable."""
    el = shape._element
    for node in el.iter(f"{{{DGM_NS}}}relIds"):
        rid = node.get(f"{{{R_NS}}}dm")
        if not rid:
            continue
        try:
            return shape.part.related_part(rid)
        except Exception:
            return None
    return None


def _diagram_nodes(shape) -> list[tuple[str, Any]]:
    """(text, lxml element of the <dgm:t> body) for each SmartArt node."""
    part = _diagram_part(shape)
    if part is None:
        return []
    try:
        root = etree.fromstring(part.blob)
    except Exception:
        return []
    out = []
    for pt in root.iter(f"{{{DGM_NS}}}pt"):
        if pt.get("type") not in (None, "node"):
            continue
        tx = pt.find(f"{{{DGM_NS}}}t")
        if tx is None:
            continue
        text = "".join(t.text or "" for t in tx.iter(f"{{{A_NS}}}t")).strip()
        if text:
            out.append((text, tx))
    return out


def _chart_summary(shape) -> str:
    chart = shape.chart
    parts = []
    try:
        if chart.has_title:
            title = chart.chart_title.text_frame.text.strip()
            if title:
                parts.append(title)
    except Exception:
        pass
    try:
        parts.append(f"series={len(chart.series)}")
    except Exception:
        pass
    return " · ".join(parts)


def _picture_name(shape) -> str:
    try:
        image = shape.image
        return image.filename or image.ext or "picture"
    except Exception:
        return "picture"


def _walk(
    shapes,
    *,
    parent_id: str = "",
    parent_index: Optional[int] = None,
    slide_no: int,
    counter: list[int],
    blocks: list[Block],
    shapes_out: list[dict],
) -> None:
    for shape in shapes:
        counter[0] += 1
        ordinal = counter[0]
        index = ordinal if parent_index is None else parent_index * 1000 + ordinal
        sid = getattr(shape, "shape_id", None)
        path = f"{parent_id}b{sid}" if parent_id else f"b{sid}" if sid is not None else f"x{ordinal}"
        stype = _shape_type_name(shape)
        ph_type = _placeholder_type(shape)
        name = getattr(shape, "name", "") or ""
        left = getattr(shape, "left", None)
        top = getattr(shape, "top", None)
        width = getattr(shape, "width", None)
        height = getattr(shape, "height", None)

        shapes_out.append(
            {
                "id": path,
                "shape_id": sid,
                "name": name,
                "shape_type": stype,
                "is_placeholder": bool(getattr(shape, "is_placeholder", False)),
                "placeholder_type": ph_type,
                "left": emu_to_in(left),
                "top": emu_to_in(top),
                "width": emu_to_in(width),
                "height": emu_to_in(height),
                "has_text": bool(getattr(shape, "has_text_frame", False)),
                "row": None,
                "col": None,
            }
        )

        # groups first: a group's own text is rarely meaningful
        if stype == "GROUP" or getattr(shape, "shape_type", None) == MSO_SHAPE_TYPE.GROUP:
            try:
                _walk(
                    shape.shapes,
                    parent_id=f"{path}/",
                    parent_index=index,
                    slide_no=slide_no,
                    counter=counter,
                    blocks=blocks,
                    shapes_out=shapes_out,
                )
            except Exception:
                pass
            continue

        if _is_diagram(shape):
            for n, (text, _el) in enumerate(_diagram_nodes(shape), 1):
                blocks.append(
                    Block(
                        id=f"{path}/dgm{n}",
                        kind="diagram",
                        text=text,
                        shape_id=sid,
                        index=index,
                        name=name,
                        shape_type=stype,
                        left=left,
                        top=top,
                        width=width,
                        height=height,
                        editable=False,
                    )
                )
            continue

        if getattr(shape, "has_chart", False):
            blocks.append(
                Block(
                    id=path,
                    kind="chart",
                    text=_chart_summary(shape),
                    shape_id=sid,
                    index=index,
                    name=name,
                    shape_type=stype,
                    left=left,
                    top=top,
                    width=width,
                    height=height,
                    editable=False,
                )
            )
            continue

        if getattr(shape, "has_table", False):
            table = shape.table
            for r, row in enumerate(table.rows, 1):
                for c, cell in enumerate(row.cells, 1):
                    text = cell.text.strip()
                    cell_shape = getattr(cell, "shape", None)
                    blocks.append(
                        Block(
                            id=f"{path}.r{r}c{c}",
                            kind="cell",
                            text=text,
                            shape_id=sid,
                            index=index,
                            name=f"{name}_Cell_{r}_{c}",
                            shape_type=stype,
                            left=cell_shape.left if cell_shape is not None else left,
                            top=cell_shape.top if cell_shape is not None else top,
                            width=cell_shape.width if cell_shape is not None else width,
                            height=cell_shape.height if cell_shape is not None else height,
                            row=r,
                            col=c,
                            target=cell,
                        )
                    )
            continue

        if getattr(shape, "has_text_frame", False):
            text = shape.text_frame.text.strip()
            blocks.append(
                Block(
                    id=path,
                    kind="text",
                    text=text,
                    shape_id=sid,
                    index=index,
                    name=name,
                    shape_type=stype,
                    is_placeholder=bool(getattr(shape, "is_placeholder", False)),
                    placeholder_type=ph_type,
                    left=left,
                    top=top,
                    width=width,
                    height=height,
                    target=shape,
                )
            )
            continue

        if stype == "PICTURE":
            blocks.append(
                Block(
                    id=path,
                    kind="picture",
                    text=_picture_name(shape),
                    shape_id=sid,
                    index=index,
                    name=name,
                    shape_type=stype,
                    left=left,
                    top=top,
                    width=width,
                    height=height,
                    editable=False,
                )
            )


def _apply_keys(blocks: list[Block]) -> None:
    """Legacy key: the text, or ``{text}_{index}_{shape_id}`` when duplicated.

    ``index``/``shape_id`` alone are not enough: every cell of one table shares
    both, so identical cell text would hand out the same key twice — and the
    first of the pair would keep the bare, ambiguous key. So any base key that
    is claimed more than once is *expanded for all of its claimants*: cells get
    their ``_r{row}c{col}``, anything else gets an ordinal. Every key on a slide
    ends up naming exactly one block.
    """
    counts: dict[str, int] = {}
    for b in blocks:
        if b.text:
            counts[b.text] = counts.get(b.text, 0) + 1

    ordered = sorted(blocks, key=lambda x: (x.index if x.index is not None else 0))
    base: dict[int, str] = {}
    for b in ordered:
        if not b.text:
            continue
        if counts[b.text] > 1:
            base[id(b)] = f"{b.text}_{b.index}_{b.shape_id}"
        else:
            base[id(b)] = b.text

    claims: dict[str, int] = {}
    for b in ordered:
        if id(b) in base:
            claims[base[id(b)]] = claims.get(base[id(b)], 0) + 1

    used: set[str] = set()
    for b in ordered:
        if id(b) not in base:
            continue
        key = base[id(b)]
        if claims[key] > 1:
            if b.row is not None:
                key = f"{key}_r{b.row}c{b.col}"
            else:
                key = f"{key}_2"
        n = 2
        while key in used:  # last-resort guard, should not trigger
            key = f"{base[id(b)]}_{n}"
            n += 1
        used.add(key)
        b.key = key


def _notes_text(slide) -> str:
    try:
        if not slide.has_notes_slide:
            return ""
        tf = slide.notes_slide.notes_text_frame
        return tf.text.strip() if tf is not None else ""
    except Exception:
        return ""


def locate_shape(slide, block_id: str):
    """The python-pptx shape a block id names, for whole-shape operations.

    ``b12/b5`` walks into a group; ``b9.r2c3`` stops at the table shape (a cell
    is not a shape, so it cannot be deleted on its own).
    """
    node = None
    shapes = slide.shapes
    for part in block_id.split("/"):
        if not part.startswith("b"):
            break
        try:
            shape_id = int(part[1:].split(".")[0])
        except ValueError:
            return None
        node = next((sh for sh in shapes if getattr(sh, "shape_id", None) == shape_id), None)
        if node is None:
            return None
        shapes = getattr(node, "shapes", [])
    return node


def open_deck(path: str) -> Deck:
    import os

    if not os.path.exists(path):
        raise FileError(f"file not found: {path}")
    if not path.lower().endswith((".pptx", ".pptm", ".potx")):
        raise FileError(
            f"only .pptx/.pptm/.potx are supported (save legacy .ppt as .pptx first): {path}"
        )
    try:
        prs = Presentation(path)
    except Exception as e:  # noqa: BLE001 - surface the real reason
        raise FileError(f"cannot open this file: {e}") from e

    deck = Deck(path=path, prs=prs)
    for idx, slide in enumerate(prs.slides, 1):
        blocks: list[Block] = []
        shapes_out: list[dict] = []
        _walk(
            slide.shapes,
            slide_no=idx,
            counter=[0],
            blocks=blocks,
            shapes_out=shapes_out,
        )
        notes = _notes_text(slide)
        if notes:
            blocks.append(
                Block(
                    id="notes",
                    kind="notes",
                    text=notes,
                    index=10**6,
                    shape_id=0,
                    editable=True,
                    target=slide,
                )
            )
        _apply_keys(blocks)
        try:
            layout_name = slide.slide_layout.name
        except Exception:
            layout_name = ""
        deck.slides.append(
            Slide(index=idx, layout=layout_name, blocks=blocks, notes=notes, shapes=shapes_out)
        )
    return deck


def iter_text(deck: Deck) -> Iterator[tuple[Slide, Block]]:
    for slide in deck.slides:
        for block in slide.blocks:
            if block.text:
                yield slide, block


def find_blocks(
    deck: Deck,
    query: str,
    *,
    regex: bool = False,
    ignore_case: bool = False,
    slides: Optional[list[int]] = None,
    kinds: Optional[set[str]] = None,
) -> list[tuple[Slide, Block]]:
    flags = re.IGNORECASE if ignore_case else 0
    pattern = re.compile(query if regex else re.escape(query), flags)
    hits = []
    for slide, block in iter_text(deck):
        if slides and slide.index not in slides:
            continue
        if kinds and block.kind not in kinds:
            continue
        if pattern.search(block.text):
            hits.append((slide, block))
    return hits
