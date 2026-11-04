"""Slide-level CRUD: new deck, list layouts, add / delete / duplicate / move."""

from __future__ import annotations

import copy
import os

from lxml import etree
from pptx import Presentation
from pptx.util import Inches

from .errors import NotFoundError, UsageError, WriteError

A_NS = "http://schemas.openxmlformats.org/drawingml/2006/main"
P_NS = "http://schemas.openxmlformats.org/presentationml/2006/main"
R_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"

DEFAULT_WIDTH = Inches(13.333)
DEFAULT_HEIGHT = Inches(7.5)


def _sld_id_lst(prs):
    return prs.slides._sldIdLst


def list_layouts(prs) -> list[dict]:
    out = []
    n = 0
    for mi, master in enumerate(prs.slide_masters):
        for layout in master.slide_layouts:
            placeholders = []
            for ph in layout.placeholders:
                try:
                    pf = ph.placeholder_format
                    placeholders.append(
                        {
                            "idx": pf.idx,
                            "type": getattr(pf.type, "name", str(pf.type)),
                            "name": ph.name,
                        }
                    )
                except Exception:
                    continue
            out.append(
                {
                    "index": n,
                    "master": mi,
                    "name": layout.name,
                    "placeholders": placeholders,
                }
            )
            n += 1
    return out


def all_layouts(prs) -> list:
    """Every layout across every master, in the order `ppt layouts` prints.

    ``prs.slide_layouts`` only reaches the first master, so a deck with two
    masters (common for real templates) would report 29 layouts but refuse
    index 20.
    """
    layouts = []
    for master in prs.slide_masters:
        layouts.extend(master.slide_layouts)
    return layouts


def find_layout(prs, selector: str | None):
    layouts = all_layouts(prs)
    if selector is None:
        return layouts[0] if layouts else None
    text = str(selector).strip()
    if text.isdigit():
        idx = int(text)
        if 0 <= idx < len(layouts):
            return layouts[idx]
        raise NotFoundError(f"no layout at index {idx} (there are {len(layouts)})")
    for layout in layouts:
        if layout.name == text:
            return layout
    for layout in layouts:
        if layout.name.lower() == text.lower():
            return layout
    if ":" in text:
        master_part, _, name = text.partition(":")
        for midx, master in enumerate(prs.slide_masters):
            if str(midx) != master_part.strip():
                continue
            for layout in master.slide_layouts:
                if layout.name.lower() == name.strip().lower():
                    return layout
    names = [l.name for l in layouts]
    raise NotFoundError(f"no layout named '{text}'", detail={"layouts": names})


def new_presentation(template: str | None = None, keep: int = 0):
    if template:
        if not os.path.exists(template):
            raise NotFoundError(f"template file not found: {template}")
        try:
            prs = Presentation(template)
        except Exception as e:  # noqa: BLE001
            raise WriteError(f"cannot open the template: {e}") from e
        delete_slides(
            prs, list(range(keep + 1, len(prs.slides) + 1)), allow_all=True
        )
    else:
        prs = Presentation()
        prs.slide_width = DEFAULT_WIDTH
        prs.slide_height = DEFAULT_HEIGHT
    return prs


def add_slide(prs, layout_selector: str | None = None, at: int | None = None) -> int:
    layout = find_layout(prs, layout_selector)
    if layout is None:
        raise UsageError("this file has no usable layout")
    slide = prs.slides.add_slide(layout)
    position = len(prs.slides) if at is None else at
    _place_slide(prs, len(prs.slides) - 1, position - 1)
    return _index_of(prs, slide)


def _index_of(prs, slide) -> int:
    for i, s in enumerate(prs.slides, 1):
        if s.slide_id == slide.slide_id:
            return i
    return -1


def _place_slide(prs, current_index0: int, target_index0: int) -> None:
    lst = _sld_id_lst(prs)
    elements = list(lst)
    target_index0 = max(0, min(target_index0, len(elements) - 1))
    if current_index0 == target_index0:
        return
    el = elements[current_index0]
    lst.remove(el)
    lst.insert(target_index0, el)


def delete_slides(prs, indexes: list[int], *, allow_all: bool = False) -> list[int]:
    total = len(prs.slides)
    for idx in indexes:
        if idx < 1 or idx > total:
            raise NotFoundError(f"slide {idx} does not exist (the deck has {total})")
    if not allow_all and len(set(indexes)) >= total:
        raise UsageError("refusing to delete every slide; use `ppt new` for an empty deck")
    removed = []
    for idx in sorted(set(indexes), reverse=True):
        lst = _sld_id_lst(prs)
        el = list(lst)[idx - 1]
        r_id = el.get(f"{{{R_NS}}}id")
        lst.remove(el)
        try:
            prs.part.drop_rel(r_id)
        except KeyError:
            pass
        removed.append(idx)
    return sorted(removed)


def move_slide(prs, index: int, to: int) -> None:
    total = len(prs.slides)
    if not 1 <= index <= total:
        raise NotFoundError(f"slide {index} does not exist (the deck has {total})")
    if not 1 <= to <= total:
        raise UsageError(f"target position must be within 1..{total}")
    _place_slide(prs, index - 1, to - 1)


def duplicate_slide(prs, index: int, at: int | None = None) -> int:
    """Copy a slide; by default the copy lands right after its source."""
    total = len(prs.slides)
    if not 1 <= index <= total:
        raise NotFoundError(f"slide {index} does not exist (the deck has {total})")
    if at is None:
        at = index + 1
    src = prs.slides[index - 1]
    dest = prs.slides.add_slide(src.slide_layout)

    # add_slide() seeds the layout's placeholders; the copy brings its own
    for shape in list(dest.shapes):
        shape._element.getparent().remove(shape._element)

    dest_tree = dest.shapes._spTree
    for element in src.shapes._spTree:
        tag = etree.QName(element).localname
        if tag in ("nvGrpSpPr", "grpSpPr"):
            continue
        clone = _clone_with_rels(element, src.part, dest.part)
        dest_tree.append(clone)

    _copy_background(src, dest)
    _copy_notes(src, dest)

    if at is not None:
        _place_slide(prs, len(prs.slides) - 1, at - 1)
    return _index_of(prs, dest)


def _clone_with_rels(element, src_part, dest_part):
    """Deep-copy XML, re-pointing every r:id/r:embed at the destination part."""
    clone = copy.deepcopy(element)
    for node in clone.iter():
        for attr, value in list(node.attrib.items()):
            if not attr.startswith(f"{{{R_NS}}}"):
                continue
            try:
                rel = src_part.rels[value]
            except KeyError:
                continue
            try:
                if rel.is_external:
                    new_id = dest_part.relate_to(
                        rel.target_ref, rel.reltype, is_external=True
                    )
                else:
                    new_id = dest_part.relate_to(rel.target_part, rel.reltype)
            except Exception:
                continue
            node.set(attr, new_id)
    return clone


def _copy_background(src, dest) -> None:
    try:
        src_cSld = src._element.find(f"{{{P_NS}}}cSld")
        dest_cSld = dest._element.find(f"{{{P_NS}}}cSld")
        bg = src_cSld.find(f"{{{P_NS}}}bg")
        if bg is not None and dest_cSld.find(f"{{{P_NS}}}bg") is None:
            dest_cSld.insert(0, copy.deepcopy(bg))
    except Exception:
        pass


def _copy_notes(src, dest) -> None:
    try:
        if not src.has_notes_slide:
            return
        text = src.notes_slide.notes_text_frame.text
        if text:
            dest.notes_slide.notes_text_frame.text = text
    except Exception:
        pass
