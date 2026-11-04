"""End-to-end: the CLI contract an agent codes against."""

from __future__ import annotations

import json
import subprocess
import sys

from pptkit.extract import open_deck


def run(*args, expect=0):
    proc = subprocess.run(
        [sys.executable, "-m", "pptkit", *[str(a) for a in args]],
        capture_output=True,
    )
    stdout = proc.stdout.decode("utf-8")
    stderr = proc.stderr.decode("utf-8")
    assert proc.returncode == expect, (
        f"exit {proc.returncode} != {expect}\nstdout={stdout}\nstderr={stderr}"
    )
    payload = json.loads(stdout) if stdout.strip().startswith("{") else None
    return payload, stdout, stderr


def test_read_envelope(chaos):
    payload, _, _ = run("read", chaos, "--json")
    assert set(payload) == {"ok", "command", "file", "data", "warnings", "error"}
    assert payload["ok"] is True
    assert payload["command"] == "read"
    data = payload["data"]
    assert data["slide_count"] == 5
    assert data["slide_size"]["width"] == 13.333
    ids = [(s["index"], b["id"]) for s in data["slides"] for b in s["blocks"]]
    assert ids and len(ids) == len(set(ids))  # unique per slide


def test_read_output_is_valid_utf8(chaos):
    _, stdout, _ = run("read", chaos, "--json")
    assert "Duplicate" in stdout  # decoded as utf-8 above; would raise otherwise


def test_read_human_output(chaos):
    payload, stdout, _ = run("read", chaos)
    assert payload is None
    assert "## Slide 1" in stdout
    assert "`b" in stdout


def test_set_with_edits_file(chaos, tmp_path):
    payload, _, _ = run("read", chaos, "--json")
    target = next(
        b for b in payload["data"]["slides"][0]["blocks"] if b["text"] == "Duplicate"
    )
    edits = tmp_path / "edits.json"
    edits.write_text(
        json.dumps({"1": {target["id"]: "Only this one"}}, ensure_ascii=False),
        encoding="utf-8",
    )

    out, _, _ = run("set", chaos, "--edits", edits, "--json")
    assert out["ok"] is True
    assert out["data"]["updated"] == 1

    fresh = open_deck(str(chaos))
    assert "Only this one" in [b.text for b in fresh.slide(1).blocks]
    assert len([b for b in fresh.slide(1).blocks if b.text == "Duplicate"]) >= 4


def test_set_dry_run_does_not_write(chaos):
    before = open_deck(str(chaos)).slide(1).to_dict()
    payload, _, _ = run("set", chaos, "--set", "Unique cell=dry run", "--dry-run", "--json")
    assert payload["data"]["dry_run"] is True
    assert payload["data"]["saved"] is None
    assert open_deck(str(chaos)).slide(1).to_dict() == before


def test_set_all_or_nothing_exit_3(chaos, tmp_path):
    edits = tmp_path / "bad.json"
    edits.write_text(
        json.dumps(
            {"1": {"Unique cell": "should be rolled back", "no such text": "must not apply"}},
            ensure_ascii=False,
        ),
        encoding="utf-8",
    )
    before = open_deck(str(chaos)).slide(1).to_dict()

    payload, _, _ = run("set", chaos, "--edits", edits, "--json", expect=3)
    assert payload["ok"] is False
    assert payload["error"]["kind"] == "not_found"
    assert payload["error"]["detail"]["errors"]
    assert open_deck(str(chaos)).slide(1).to_dict() == before


def test_missing_file_exits_2(tmp_path):
    payload, _, _ = run("read", tmp_path / "nope.pptx", "--json", expect=2)
    assert payload["ok"] is False
    assert payload["error"]["kind"] == "file"


def test_find_command(chaos):
    payload, _, _ = run("find", chaos, "Duplicat", "--json")
    assert payload["data"]["count"] >= 5
    assert {m["slide"] for m in payload["data"]["matches"]} == {1}


def test_layouts_command(chaos):
    payload, _, _ = run("layouts", chaos, "--json")
    assert payload["data"]["layouts"]
    assert "index" in payload["data"]["layouts"][0]


def test_verify_command(chaos):
    payload, _, _ = run("verify", chaos, "--json")
    assert payload["data"]["loads"] is True
    assert payload["data"]["slide_count"] == 5


def test_notes_roundtrip(chaos, tmp_path):
    text = tmp_path / "note.txt"
    text.write_text("notes read from a file, keeping argv ASCII", encoding="utf-8")
    run("notes", chaos, "--slide", "2", "--set-from", text, "--json")

    payload, _, _ = run("notes", chaos, "--slide", "2", "--json")
    assert payload["data"]["notes"] == "notes read from a file, keeping argv ASCII"


def test_add_and_delete_slide_via_cli(chaos):
    payload, _, _ = run("add-slide", chaos, "--layout", "6", "--at", "2", "--json")
    assert payload["data"]["index"] == 2
    assert payload["data"]["slide_count"] == 6

    payload, _, _ = run("delete-slide", chaos, "--slides", "2", "--json")
    assert payload["data"]["removed"] == [2]
    assert payload["data"]["slide_count"] == 5


def test_out_flag_writes_a_copy(chaos, tmp_path):
    copy = tmp_path / "copy.pptx"
    payload, _, _ = run(
        "set", chaos, "--replace", "Unique cell", "--with", "copied", "-o", copy, "--json"
    )
    assert payload["data"]["saved"].endswith("copy.pptx")
    assert "copied" in [b.text for b in open_deck(str(copy)).slide(1).blocks]
    assert "Unique cell" in [b.text for b in open_deck(str(chaos)).slide(1).blocks]


def test_regex_replace(chaos):
    payload, _, _ = run(
        "set", chaos, "--replace", r"^(Second|Third) page only$", "--with", r"\1 page", "--regex", "--json"
    )
    assert payload["data"]["updated"] == 2

    fresh = open_deck(str(chaos))
    assert {"Second page", "Third page"} <= {b.text for s in fresh.slides for b in s.blocks}


def test_add_textbox_and_table_cli(simple, tmp_path):
    text = tmp_path / "t.txt"
    text.write_text("a text box added from the CLI", encoding="utf-8")
    run(
        "add-textbox", simple, "--slide", "1", "--text-file", text,
        "--left", "1", "--top", "4", "--font-size", "22", "--json",
    )

    data = tmp_path / "table.json"
    data.write_text(
        json.dumps([["Metric", "Value"], ["A", "1"]], ensure_ascii=False), encoding="utf-8"
    )
    run("add-table", simple, "--slide", "1", "--data", data, "--json")

    fresh = open_deck(str(simple))
    assert "a text box added from the CLI" in [b.text for b in fresh.slide(1).blocks]
    assert {(b.row, b.col) for b in fresh.slide(1).blocks if b.kind == "cell"} == {
        (1, 1), (1, 2), (2, 1), (2, 2)
    }


def test_inspect_slide(simple):
    payload, _, _ = run("inspect", simple, "--slide", "1", "--json")
    assert payload["data"]["shapes"]
    assert payload["data"]["shapes"][0]["left"] is not None


def test_new_refuses_to_overwrite_without_force(simple, tmp_path):
    target = tmp_path / "exists.pptx"
    target.write_bytes(simple.read_bytes())
    payload, _, _ = run("new", target, "--json", expect=1)
    assert payload["ok"] is False
    run("new", target, "--force", "--json")


def test_unknown_command_exits_1():
    proc = subprocess.run([sys.executable, "-m", "pptkit", "nope"], capture_output=True)
    assert proc.returncode == 1
