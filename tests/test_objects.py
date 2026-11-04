"""Inserting text boxes, pictures and tables."""

from __future__ import annotations

from pptx.enum.shapes import MSO_SHAPE_TYPE
from pptx.util import Emu

from build_decks import PNG_1PX
from pptkit import objects
from pptkit.extract import locate_shape, open_deck
from pptkit.store import save


def test_add_textbox_with_style(simple):
    deck = open_deck(str(simple))
    slide = deck.prs.slides[0]

    objects.add_textbox(
        slide,
        "A new block\nsecond line",
        left=1.5,
        top=4.0,
        width=6.0,
        height=1.2,
        font_size=20,
        bold=True,
        color="C00000",
        align="center",
    )
    save(deck)

    fresh = open_deck(str(simple))
    block = next(b for b in fresh.slide(1).blocks if "A new block" in b.text)
    runs = block.target.text_frame.paragraphs[0].runs
    assert runs[0].font.size.pt == 20
    assert runs[0].font.bold is True
    assert str(runs[0].font.color.rgb) == "C00000"
    assert len(block.target.text_frame.paragraphs) == 2
    assert block.left == Emu(int(1.5 * 914400))


def test_add_picture(simple, tmp_path):
    deck = open_deck(str(simple))
    image = tmp_path / "dot.png"
    image.write_bytes(PNG_1PX)

    objects.add_picture(deck.prs.slides[0], str(image), left=2, top=3, width=1, height=1)
    save(deck)

    fresh = open_deck(str(simple))
    pictures = [
        sh for sh in fresh.prs.slides[0].shapes if sh.shape_type == MSO_SHAPE_TYPE.PICTURE
    ]
    assert len(pictures) == 1


def test_add_table_and_read_it_back(simple):
    deck = open_deck(str(simple))
    rows = [["Metric", "Value"], ["Revenue", "100"], ["Profit", "20"]]

    objects.add_table(deck.prs.slides[0], rows, left=1, top=3, width=6)
    save(deck)

    fresh = open_deck(str(simple))
    cells = {(b.row, b.col): b.text for b in fresh.slide(1).blocks if b.kind == "cell"}
    assert cells[(1, 1)] == "Metric"
    assert cells[(3, 2)] == "20"
    assert len(cells) == 6


def test_delete_one_of_several_identical_boxes(chaos):
    deck = open_deck(str(chaos))
    dups = [b for b in deck.slide(1).blocks if b.text == "Duplicate" and b.kind == "text"]
    target = dups[1]
    before = len([b for b in deck.slide(1).blocks if b.text == "Duplicate"])
    shape = locate_shape(deck.prs.slides[0], target.id)
    assert shape is not None and shape.shape_id == target.shape_id

    shape._element.getparent().remove(shape._element)
    save(deck)

    fresh = open_deck(str(chaos))
    remaining = [b for b in fresh.slide(1).blocks if b.text == "Duplicate"]
    assert len(remaining) == before - 1
    assert target.id not in {b.id for b in remaining}


def test_locate_shape_walks_into_groups(chaos):
    deck = open_deck(str(chaos))
    child = next(b for b in deck.slide(1).blocks if b.text == "Inside group")
    shape = locate_shape(deck.prs.slides[0], child.id)
    assert shape is not None
    assert shape.text_frame.text == "Inside group"


def test_added_table_cells_are_editable_by_id(simple):
    from pptkit.edit import apply_plan, plan_edits

    deck = open_deck(str(simple))
    objects.add_table(deck.prs.slides[0], [["A", "B"]], left=1, top=3, width=4)
    save(deck)

    fresh = open_deck(str(simple))
    cell = next(
        b for b in fresh.slide(1).blocks if b.kind == "cell" and (b.row, b.col) == (1, 2)
    )
    plan, errors = plan_edits(fresh, {"1": {cell.id: "edited"}})
    assert not errors
    apply_plan(plan)
    save(fresh)
    assert "edited" in [b.text for b in open_deck(str(simple)).slide(1).blocks]
