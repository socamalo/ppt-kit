"""Round-trip fidelity: a deck that is read and written back loses nothing."""

from __future__ import annotations

import posixpath
import zipfile

from lxml import etree
from pptx.enum.shapes import MSO_SHAPE_TYPE

from pptkit.extract import open_deck
from pptkit.store import save


def assert_package_is_sound(path: str) -> None:
    """A cheap stand-in for "PowerPoint will open this": every part parses and
    every relationship points at a part that exists."""
    with zipfile.ZipFile(path) as zf:
        names = set(zf.namelist())
        for name in names:
            if name.endswith((".xml", ".rels")):
                etree.fromstring(zf.read(name))  # well-formed
        for name in names:
            if not name.endswith(".rels"):
                continue
            base = posixpath.dirname(posixpath.dirname(name))
            root = etree.fromstring(zf.read(name))
            for rel in root:
                if rel.get("TargetMode") == "External":
                    continue
                target = rel.get("Target") or ""
                if target.startswith("/"):
                    resolved = target.lstrip("/")
                else:
                    resolved = posixpath.normpath(posixpath.join(base, target))
                assert resolved in names, f"{name} -> {target} is missing"


def _shape_inventory(prs):
    out = []
    for i, slide in enumerate(prs.slides, 1):
        for shape in slide.shapes:
            out.append((i, shape.shape_id, str(shape.shape_type), shape.name))
    return sorted(out)


def test_save_without_edits_is_lossless(chaos):
    before = open_deck(str(chaos))
    inventory_before = _shape_inventory(before.prs)
    texts_before = {(s.index, b.id): b.text for s in before.slides for b in s.blocks}

    save(before)

    after = open_deck(str(chaos))
    assert _shape_inventory(after.prs) == inventory_before
    assert {(s.index, b.id): b.text for s in after.slides for b in s.blocks} == texts_before


def test_text_edit_does_not_disturb_other_shapes(chaos):
    from pptkit.edit import apply_plan, plan_edits

    before = open_deck(str(chaos))
    inventory = _shape_inventory(before.prs)
    pictures_before = sum(
        1
        for s in before.prs.slides
        for sh in s.shapes
        if sh.shape_type == MSO_SHAPE_TYPE.PICTURE
    )

    plan, errors = plan_edits(before, {"1": {"Unique cell": "changed this one"}})
    assert not errors
    apply_plan(plan)
    save(before)

    after = open_deck(str(chaos))
    assert _shape_inventory(after.prs) == inventory
    assert (
        sum(
            1
            for s in after.prs.slides
            for sh in s.shapes
            if sh.shape_type == MSO_SHAPE_TYPE.PICTURE
        )
        == pictures_before
    )


def test_deck_reopens_after_many_operations(chaos):
    from pptkit import slides as ops
    from pptkit.edit import apply_plan, plan_edits

    deck = open_deck(str(chaos))
    plan, errors = plan_edits(
        deck, {"1": {"Unique cell": "one"}, "2": {"Section title": "two"}}
    )
    assert not errors
    apply_plan(plan)
    ops.add_slide(deck.prs, "6", at=1)
    ops.duplicate_slide(deck.prs, 3)
    ops.move_slide(deck.prs, 2, 4)
    ops.delete_slides(deck.prs, [6])
    save(deck)

    fresh = open_deck(str(chaos))
    assert len(fresh.slides) == 6
    assert any(b.text == "one" for s in fresh.slides for b in s.blocks)
    assert any(b.text == "two" for s in fresh.slides for b in s.blocks)


def test_saved_package_is_structurally_sound(chaos):
    from pptkit.edit import apply_plan, plan_edits

    deck = open_deck(str(chaos))
    plan, errors = plan_edits(deck, {"1": {"Unique cell": "structure check"}})
    assert not errors
    apply_plan(plan)
    save(deck)

    assert_package_is_sound(str(chaos))


def test_locked_target_fails_loudly_and_keeps_the_original(chaos, monkeypatch):
    """The 'the file is open in PowerPoint' path: clear error, file untouched."""
    import os

    import pytest

    from pptkit.errors import WriteError

    original = chaos.read_bytes()

    def boom(src, dst):
        raise PermissionError("file is locked")

    monkeypatch.setattr(os, "replace", boom)
    deck = open_deck(str(chaos))
    with pytest.raises(WriteError) as excinfo:
        save(deck)

    assert "replace" in str(excinfo.value) or "PowerPoint" in str(excinfo.value)
    assert chaos.read_bytes() == original
    leftovers = [p for p in os.listdir(chaos.parent) if "pptk-tmp" in p]
    assert leftovers == []


def test_no_temp_files_left_behind(chaos):
    import os

    save(open_deck(str(chaos)))
    leftovers = [p for p in os.listdir(os.path.dirname(str(chaos))) if "pptk-tmp" in p]
    assert leftovers == []
