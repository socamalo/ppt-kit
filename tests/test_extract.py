"""Reading: every container yields addressable blocks, duplicates get keys."""

from __future__ import annotations

from pptkit.extract import find_blocks, open_deck


def test_reads_all_containers(chaos):
    deck = open_deck(str(chaos))
    assert len(deck.slides) == 5

    slide1 = deck.slide(1)
    kinds = {b.kind for b in slide1.blocks}
    assert {"text", "cell"} <= kinds

    texts = [b.text for b in slide1.blocks]
    assert texts.count("Duplicate") >= 5  # 3 boxes + 1 table cell + 1 in group


def test_duplicate_ids_are_unique_and_stable(chaos):
    deck = open_deck(str(chaos))
    slide1 = deck.slide(1)
    dups = [b for b in slide1.blocks if b.text == "Duplicate"]
    assert len(dups) >= 5
    ids = [b.id for b in dups]
    assert len(ids) == len(set(ids))
    assert all(i.startswith("b") for i in ids)


def test_duplicate_keys_carry_index_and_shape_id(chaos):
    deck = open_deck(str(chaos))
    slide1 = deck.slide(1)
    dups = [b for b in slide1.blocks if b.text == "Duplicate"]
    for block in dups:
        assert block.key != "Duplicate"
        assert block.key.startswith("Duplicate_")
        assert f"_{block.index}_{block.shape_id}" in block.key
    keys = [b.key for b in dups]
    assert len(keys) == len(set(keys))


def test_identical_cells_in_one_table_get_distinct_keys(chaos):
    """Two cells of the same table share index *and* shape_id; keys must not."""
    from pptkit.edit import resolve

    deck = open_deck(str(chaos))
    cells = [b for b in deck.slide(1).blocks if b.kind == "cell" and b.text == "Duplicate"]
    assert len(cells) == 2
    assert len({c.key for c in cells}) == 2
    for cell in cells:
        assert f"_r{cell.row}c{cell.col}" in cell.key
        assert resolve(deck.slide(1), cell.key) == [cell]


def test_every_address_points_at_exactly_itself(chaos):
    """The whole duplicate problem in one invariant, on a deck built to be nasty."""
    from pptkit.edit import resolve

    deck = open_deck(str(chaos))
    for slide in deck.slides:
        for block in slide.blocks:
            if not block.text:
                continue
            assert resolve(slide, block.id) == [block], block.id
            assert resolve(slide, block.key) == [block], block.key


def test_unique_text_key_is_the_text_itself(chaos):
    deck = open_deck(str(chaos))
    unique = [b for b in deck.slide(1).blocks if b.text == "Unique cell"]
    assert len(unique) == 1
    assert unique[0].key == "Unique cell"


def test_group_children_are_walked(chaos):
    deck = open_deck(str(chaos))
    grouped = [b for b in deck.slide(1).blocks if "/b" in b.id]
    assert {b.text for b in grouped} == {"Inside group", "Duplicate"}
    assert all(b.shape_id for b in grouped)


def test_table_cells_have_row_col(chaos):
    deck = open_deck(str(chaos))
    cells = [b for b in deck.slide(1).blocks if b.kind == "cell"]
    assert {(b.row, b.col) for b in cells} >= {(1, 1), (1, 2), (2, 1)}
    assert all(".r" in b.id and "c" in b.id for b in cells)


def test_notes_are_a_block(chaos):
    deck = open_deck(str(chaos))
    notes = [b for b in deck.slide(2).blocks if b.kind == "notes"]
    assert len(notes) == 1
    assert notes[0].id == "notes"
    assert notes[0].text == "notes for slide two"
    assert deck.slide(2).notes == "notes for slide two"


def test_same_text_on_different_slides_is_not_a_duplicate(chaos):
    deck = open_deck(str(chaos))
    for idx in (2, 3):
        block = next(b for b in deck.slide(idx).blocks if b.text == "Section title")
        assert block.key == "Section title"


def test_find_matches_substring_and_regex(chaos):
    deck = open_deck(str(chaos))
    hits = find_blocks(deck, "Duplicat")
    assert len(hits) >= 5
    assert {s.index for s, _ in hits} == {1}

    regex_hits = find_blocks(deck, r"^(Second|Third) page only$", regex=True)
    assert sorted(b.text for _, b in regex_hits) == ["Second page only", "Third page only"]


def test_empty_text_boxes_are_hidden_by_default(chaos):
    deck = open_deck(str(chaos))
    cells = [b for b in deck.slide(1).blocks if b.kind == "cell"]
    assert any(b.text == "" for b in cells)  # present in the model
    payload = deck.slide(1).to_dict()
    assert all(b["text"].strip() for b in payload["blocks"])
