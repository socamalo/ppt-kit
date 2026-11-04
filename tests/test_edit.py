"""The hard part: picking the *right* one out of many identical strings."""

from __future__ import annotations

from pptkit.edit import apply_plan, plan_edits, resolve, set_text_frame_text
from pptkit.extract import open_deck
from pptkit.store import save


def _dups(deck, slide_no=1):
    return [b for b in deck.slide(slide_no).blocks if b.text == "Duplicate"]


def _edit(deck, edits, slides=None):
    plan, errors = plan_edits(deck, edits, slides=slides)
    assert not errors, errors
    return apply_plan(plan)


def test_edit_one_duplicate_by_id_leaves_the_others(chaos):
    deck = open_deck(str(chaos))
    dups = _dups(deck)
    target = dups[1]

    changes = _edit(deck, {"1": {target.id: "Only this one"}})

    assert len(changes) == 1
    assert changes[0].block_id == target.id
    assert len([b for b in deck.slide(1).blocks if b.text == "Duplicate"]) == len(dups) - 1
    assert "Only this one" in [b.text for b in deck.slide(1).blocks]

    save(deck)
    reopened = open_deck(str(chaos))
    assert "Only this one" in [b.text for b in reopened.slide(1).blocks]
    assert len([b for b in reopened.slide(1).blocks if b.text == "Duplicate"]) == len(dups) - 1


def test_edit_one_duplicate_by_legacy_key(chaos):
    deck = open_deck(str(chaos))
    dups = _dups(deck)
    target = dups[2]

    changes = _edit(deck, {"1": {target.key: "Changed by key"}})

    assert len(changes) == 1
    assert changes[0].block_id == target.id
    assert len([b for b in deck.slide(1).blocks if b.text == "Duplicate"]) == len(dups) - 1


def test_edit_by_plain_text_hits_every_occurrence(chaos):
    deck = open_deck(str(chaos))
    dups = _dups(deck)

    changes = _edit(deck, {"1": {"Duplicate": "All replaced"}})

    assert len(changes) == len(dups)
    assert {c.block_id for c in changes} == {b.id for b in dups}
    assert not [b for b in deck.slide(1).blocks if b.text == "Duplicate"]


def test_legacy_key_survives_text_drift(chaos):
    """Key was captured before an edit; the text moved on; key still resolves."""
    deck = open_deck(str(chaos))
    dup = _dups(deck)[0]
    stale_key = dup.key
    set_text_frame_text(dup.target.text_frame, "changed elsewhere")
    dup.text = "changed elsewhere"

    changes = _edit(deck, {"1": {stale_key: "rescued"}})
    assert len(changes) == 1
    assert changes[0].block_id == dup.id


def test_edit_group_child_by_id(chaos):
    deck = open_deck(str(chaos))
    child = next(b for b in deck.slide(1).blocks if b.text == "Inside group")

    changes = _edit(deck, {"1": {child.id: "Group updated"}})

    assert len(changes) == 1
    assert "Group updated" in [b.text for b in deck.slide(1).blocks]
    save(deck)
    assert "Group updated" in [b.text for b in open_deck(str(chaos)).slide(1).blocks]


def test_edit_table_cell_by_id(chaos):
    deck = open_deck(str(chaos))
    cell = next(b for b in deck.slide(1).blocks if b.kind == "cell" and b.text == "Unique cell")

    changes = _edit(deck, {"1": {cell.id: "New cell"}})

    assert len(changes) == 1
    save(deck)
    cells = [b for b in open_deck(str(chaos)).slide(1).blocks if b.kind == "cell"]
    by_pos = {(b.row, b.col): b.text for b in cells}
    assert by_pos[(1, 2)] == "New cell"
    assert by_pos[(1, 1)] == "Duplicate"


def test_same_text_other_slide_untouched(chaos):
    deck = open_deck(str(chaos))
    _edit(deck, {"2": {"Section title": "New section"}})

    assert "Section title" not in [b.text for b in deck.slide(2).blocks]
    assert "Section title" in [b.text for b in deck.slide(3).blocks]


def test_multiline_edit_creates_paragraphs_and_keeps_bullets(chaos):
    deck = open_deck(str(chaos))
    target = next(b for b in deck.slide(1).blocks if b.text == "First paragraph")

    _edit(deck, {"1": {target.id: "one\ntwo\nthree"}})

    frame = target.target.text_frame
    assert [p.text for p in frame.paragraphs] == ["one", "two", "three"]
    for paragraph in frame.paragraphs:
        p_pr = paragraph._p.find(
            "{http://schemas.openxmlformats.org/drawingml/2006/main}pPr"
        )
        assert p_pr is not None
        assert p_pr.get("marL") == "228600"
        assert p_pr.find(
            "{http://schemas.openxmlformats.org/drawingml/2006/main}buChar"
        ) is not None


def test_multiline_collapse_into_one_paragraph(chaos):
    deck = open_deck(str(chaos))
    target = next(b for b in deck.slide(1).blocks if b.text == "First paragraph")
    _edit(deck, {"1": {target.id: "alpha\nbeta"}})

    _edit(deck, {"1": {target.id: "back to one line"}})
    assert [p.text for p in target.target.text_frame.paragraphs] == ["back to one line"]


def test_bold_survives_the_edit(chaos):
    deck = open_deck(str(chaos))
    target = next(b for b in deck.slide(1).blocks if b.text == "Bold title")
    assert target.target.text_frame.paragraphs[0].runs[0].font.bold is True

    _edit(deck, {"1": {target.id: "still bold after the edit"}})
    save(deck)

    fresh = open_deck(str(chaos))
    block = next(b for b in fresh.slide(1).blocks if b.text == "still bold after the edit")
    assert block.target.text_frame.paragraphs[0].runs[0].font.bold is True


def test_mixed_formatting_collapses_to_first_run_style(chaos):
    """Documented trade-off: the replacement inherits run 0, later runs go."""
    deck = open_deck(str(chaos))
    mixed = next(b for b in deck.slide(1).blocks if "Plain text" in b.text)

    _edit(deck, {"1": {mixed.id: "A fresh line"}})

    runs = mixed.target.text_frame.paragraphs[0].runs
    assert len(runs) == 1
    assert runs[0].text == "A fresh line"
    assert not runs[0].font.bold


def test_spaceholder_is_a_noop(chaos):
    deck = open_deck(str(chaos))
    target = _dups(deck)[0]
    plan, errors = plan_edits(deck, {"1": {target.id: "spaceholder"}})
    assert not errors and plan == []


def test_notes_can_be_edited(chaos):
    deck = open_deck(str(chaos))
    _edit(deck, {"2": {"notes": "a new note"}})
    save(deck)
    assert open_deck(str(chaos)).slide(2).notes == "a new note"


def test_batch_is_all_or_nothing(chaos):
    deck = open_deck(str(chaos))
    before = [b.text for b in deck.slide(1).blocks]

    plan, errors = plan_edits(
        deck, {"1": {"Unique cell": "changed", "no such text": "also changed"}}
    )
    assert errors and len(errors) == 1
    assert not plan  # caller must refuse the whole batch

    after = [b.text for b in deck.slide(1).blocks]
    assert before == after


def test_uneditable_diagram_is_reported_not_crashed(chaos):
    deck = open_deck(str(chaos))
    block = deck.slide(1).blocks[0]
    block.editable = False
    plan, errors = plan_edits(deck, {"1": {block.id: "x"}})
    assert errors and "cannot be edited" in errors[0]["error"]
    assert not plan


def test_text_with_underscores_still_parses_as_a_key(chaos):
    """`A_B_12` duplicated twice must not be mistaken for a text_index_sid key."""
    deck = open_deck(str(chaos))
    for _ in range(2):
        box = deck.prs.slides[1].shapes.add_textbox(0, 0, 914400, 914400)
        box.text_frame.text = "A_B_12"
    save(deck)

    fresh = open_deck(str(chaos))
    dups = [b for b in fresh.slide(2).blocks if b.text == "A_B_12"]
    assert len(dups) == 2
    for block in dups:
        assert resolve(fresh.slide(2), block.key) == [block]
        assert resolve(fresh.slide(2), block.id) == [block]


def test_ids_are_stable_across_a_save(chaos):
    deck = open_deck(str(chaos))
    before = {(s.index, b.id) for s in deck.slides for b in s.blocks}
    _edit(deck, {"1": {_dups(deck)[0].id: "tweak"}})
    save(deck)

    fresh = open_deck(str(chaos))
    after = {(s.index, b.id) for s in fresh.slides for b in s.blocks}
    assert before == after


def test_untouched_duplicates_keep_their_old_keys(chaos):
    """A key captured before an edit must still address its block afterwards."""
    deck = open_deck(str(chaos))
    dups = _dups(deck)
    keys = {b.id: b.key for b in dups}

    _edit(deck, {"1": {dups[0].id: "first one changed"}})
    save(deck)

    fresh = open_deck(str(chaos))
    for block_id, old_key in keys.items():
        if block_id == dups[0].id:
            continue
        block = next(b for b in fresh.slide(1).blocks if b.id == block_id)
        assert block.key == old_key
        assert resolve(fresh.slide(1), old_key) == [block]


def test_emoji_and_mixed_script_survive(chaos):
    deck = open_deck(str(chaos))
    target = next(b for b in deck.slide(5).blocks if "Mixed script" in b.text)
    _edit(deck, {"5": {target.id: "Reworked 5G-A · Ünïcödé 🚁"}})
    save(deck)
    assert "Reworked 5G-A · Ünïcödé 🚁" in [b.text for b in open_deck(str(chaos)).slide(5).blocks]
