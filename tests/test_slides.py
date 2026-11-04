"""Slide-level CRUD."""

from __future__ import annotations

from pptx.enum.shapes import MSO_SHAPE_TYPE

from pptkit.extract import open_deck
from pptkit.store import save


def _slides(deck):
    return deck.prs.slides


def test_add_slide_appends_and_inserts(chaos):
    from pptkit import slides as ops

    deck = open_deck(str(chaos))
    layout_index = 6

    new_index = ops.add_slide(deck.prs, str(layout_index))
    assert new_index == 6

    inserted = ops.add_slide(deck.prs, str(layout_index), at=1)
    assert inserted == 1
    assert len(_slides(deck)) == 7

    save(deck)
    assert len(open_deck(str(chaos)).slides) == 7


def test_add_slide_by_layout_name(chaos):
    from pptkit import slides as ops

    deck = open_deck(str(chaos))
    name = ops.list_layouts(deck.prs)[-1]["name"]

    index = ops.add_slide(deck.prs, name)
    assert index == 6
    assert deck.prs.slides[5].slide_layout.name == name


def test_delete_slide_removes_the_right_one(chaos):
    from pptkit import slides as ops

    deck = open_deck(str(chaos))
    removed = ops.delete_slides(deck.prs, [2])
    assert removed == [2]
    save(deck)

    fresh = open_deck(str(chaos))
    assert len(fresh.slides) == 4
    assert all("Second page only" not in [b.text for b in s.blocks] for s in fresh.slides)
    assert any("Third page only" in [b.text for b in s.blocks] for s in fresh.slides)


def test_delete_multiple_in_one_go(chaos):
    from pptkit import slides as ops

    deck = open_deck(str(chaos))
    removed = ops.delete_slides(deck.prs, [1, 4])
    assert removed == [1, 4]
    assert len(_slides(deck)) == 3


def test_duplicate_slide_copies_text_and_picture(chaos):
    from pptkit import slides as ops

    deck = open_deck(str(chaos))
    src = deck.prs.slides[3]  # slide 4: text + picture
    pictures_before = sum(1 for sh in src.shapes if sh.shape_type == MSO_SHAPE_TYPE.PICTURE)
    assert pictures_before == 1

    index = ops.duplicate_slide(deck.prs, 4)
    assert index == 5  # lands right after its source
    save(deck)

    fresh = open_deck(str(chaos))
    copy = fresh.prs.slides[4]
    texts = [sh.text_frame.text for sh in copy.shapes if sh.has_text_frame]
    assert "Page with picture" in texts
    copied_pictures = sum(1 for sh in copy.shapes if sh.shape_type == MSO_SHAPE_TYPE.PICTURE)
    assert copied_pictures == 1
    # the original is untouched
    assert any(
        sh.text_frame.text == "Page with picture"
        for sh in fresh.prs.slides[3].shapes
        if sh.has_text_frame
    )


def test_duplicate_preserves_notes(chaos):
    from pptkit import slides as ops

    deck = open_deck(str(chaos))
    index = ops.duplicate_slide(deck.prs, 2)
    assert index == 3
    save(deck)
    assert open_deck(str(chaos)).slide(3).notes == "notes for slide two"


def test_move_slide(chaos):
    from pptkit import slides as ops

    deck = open_deck(str(chaos))
    first_text = [b.text for b in deck.slides[0].blocks if b.text][0]
    ops.move_slide(deck.prs, 1, 5)
    save(deck)

    fresh = open_deck(str(chaos))
    assert first_text in [b.text for b in fresh.slides[4].blocks]
    assert "Section title" in [b.text for b in fresh.slides[0].blocks]


def test_new_deck_from_template_keeps_layouts(chaos, tmp_path):
    from pptkit import slides as ops
    from pptkit.model import Deck

    out = tmp_path / "from_template.pptx"
    prs = ops.new_presentation(str(chaos), keep=1)
    assert len(prs.slides) == 1
    assert len(ops.list_layouts(prs)) >= 1

    save(Deck(path=str(out), prs=prs), str(out))
    assert len(open_deck(str(out)).slides) == 1


def test_new_blank_deck_is_16_9(tmp_path):
    from pptkit import slides as ops

    prs = ops.new_presentation()
    assert prs.slide_width == int(13.333 * 914400)
    assert prs.slide_height == 7.5 * 914400


def test_cannot_delete_every_slide(chaos):
    import pytest

    from pptkit import slides as ops
    from pptkit.errors import UsageError

    deck = open_deck(str(chaos))
    with pytest.raises(UsageError):
        ops.delete_slides(deck.prs, [1, 2, 3, 4, 5])


def test_move_out_of_range_is_refused(chaos):
    import pytest

    from pptkit import slides as ops
    from pptkit.errors import UsageError

    deck = open_deck(str(chaos))
    with pytest.raises(UsageError):
        ops.move_slide(deck.prs, 1, 99)
