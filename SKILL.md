---
name: ppt-kit
description: Read and write PowerPoint files (.pptx/.potx) from the command line — use it for any .pptx work instead of writing your own script. Every text block gets a stable address (b6, b12/b5, b9.r2c3, notes), so identical strings on one slide stay individually editable; add/remove/duplicate/move slides, speaker notes, text boxes, pictures, tables, and new decks from a template or from scratch. Reach for `ppt` when reading a deck, rewriting or translating text, batch replacing, restructuring slides, or building a new one. Do not hand-roll python-pptx or pptxgenjs for this, and do not edit OOXML by hand — the only jobs it will not do are SmartArt/chart text and swapping a slide master. Triggers: pptx, potx, powerpoint, deck, slides, presentation, slide deck.
---

# ppt-kit — CRUD a .pptx from the shell

The command is `ppt`. Install it from this repository:

```bash
uv tool install --from /path/to/ppt-kit ppt-kit
```

It edits the OOXML package directly: no PowerPoint, no open window.

## Don't write a script first

When a .pptx needs changing, the first move is not "write a Python script".
Check whether this chain already does the job:

| You want to | Run |
|---|---|
| read a deck | `ppt read --json` |
| change text (one block or a batch) | `ppt set --edits edits.json` |
| replace a phrase everywhere | `ppt set --replace X --with Y` |
| add / delete / copy / reorder slides | `add-slide` / `delete-slide` / `duplicate-slide` / `move-slide` |
| add text, pictures, tables | `add-textbox` / `add-image` / `add-table` |
| build a deck from a template | `ppt new --template`, then `add-slide` + `set` |

**Don't wrap these in a loop script.** One `--edits` batch changes dozens of
blocks in a single atomic write; a shell loop is slower and can leave the deck
half-edited. If you write code at all, write the *content* (edits.json, text
files, table JSON), not a program that drives PowerPoint.

The only jobs this tool will not do are editing SmartArt/chart internals and
swapping a slide master.

## Rule one: read the ids, then edit by id

Every piece of text has an **address** that names a *place*, not a string:

| Address | Points at |
|---|---|
| `b6` | the text box whose shape id is 6 |
| `b12/b5` | shape 5 inside group 12 |
| `b9.r2c3` | row 2, column 3 of table 9 |
| `b7/dgm3` | SmartArt node 3 (read-only) |
| `notes` | speaker notes |

**Five boxes on one slide all reading "Subtitle" are five different ids.** Editing
by id moves exactly one. Editing by text moves all of them — which is why text is
the wrong key here.

Legacy keys (`Subtitle_2_3`, i.e. `text_index_shapeid`, and the bare text) still
resolve, so older prompts keep working, but **new work goes by id**.

## Standard flow

```bash
ppt read deck.pptx --json                    # every block: id, text, kind
# ... build edits.json from those ids ...
ppt set deck.pptx --edits edits.json --json  # all or nothing
ppt verify deck.pptx --json                  # page count, empty slides
```

`edits.json` is keyed by slide, then by address:

```json
{"1": {"b6": "Q3 results: revenue up 20%"},
 "3": {"b9.r2c1": "Metric", "notes": "put the conclusion here"}}
```

Want to see what would change first? `--dry-run`. Want a copy instead of an
in-place edit? `-o new.pptx`.

## Intent → command

| You want to | Run |
|---|---|
| see what a deck says | `ppt read f.pptx --json` |
| read only some pages | `ppt read f.pptx --slides 2,4-6 --json` |
| replace a phrase everywhere | `ppt find f.pptx Revenue --json`, or `ppt set f.pptx --replace Revenue --with Sales --json` |
| change one block | edits.json → `ppt set f.pptx --edits edits.json --json` |
| translate a deck | `read`, translate yourself, `set` back by id (page by page if it is long) |
| see shapes / placeholders | `ppt inspect f.pptx --slide 3 --json` |
| start from a template | `ppt new out.pptx --template brand.pptx --json` |
| build a page like an existing one | `ppt duplicate-slide f.pptx --slide 7` then `set` the text — best when layouts carry artwork but no placeholders |
| add a page | `ppt add-slide f.pptx --layout 4 --at 2 --json` (see `ppt layouts`) |
| remove pages | `ppt delete-slide f.pptx --slides 3,5-6 --json` |
| copy a page (with pictures and tables) | `ppt duplicate-slide f.pptx --slide 2 --json` (lands right after the source) |
| reorder pages | `ppt move-slide f.pptx --slide 5 --to 1 --json` |
| write notes | `ppt notes f.pptx --slide 2 --set-from note.txt --json` |
| insert a text box / picture / table | `ppt add-textbox` / `add-image` / `add-table` |
| delete a text box or picture | `ppt delete-shape f.pptx --slide 2 --id b9 --json` |

## The contract

Every command with `--json` prints exactly one envelope:

```json
{"ok": true, "command": "read", "file": "deck.pptx",
 "data": {}, "warnings": [], "error": null}
```

On failure `ok` is false and `error` carries `kind` / `message` / `detail`.
Exit codes are fixed: `0` ok · `1` usage · `2` file · `3` not found/ambiguous · `4` write.

**Batches are all-or-nothing.** One unresolvable id and the whole batch is
skipped, the file untouched, every problem listed in `error.detail.errors`. Do not
mix guesses into a batch.

## Two traps

**Non-ASCII arguments.** Some shells (notably Git Bash on Windows) abort on
non-ASCII characters in a command line. Put the text in a file and pass the path:
`--edits`, `--text-file`, `--set-from`, `--data`. Paths with non-ASCII names have
the same problem — copy the file to an ASCII path first.

**The file is open in PowerPoint.** Saving then fails with a clear message (exit
4) and the original is untouched. Close the file first.

## Boundaries

- **SmartArt and chart text are read-only.** `read` shows them, `set` refuses with
  a clear error. Edit those in PowerPoint.
- **No theme swap, no built-in model.** Restyling and translation are your job:
  read, think, `set`.
- **No live PowerPoint window.** Everything acts on a file.
- **Legacy `.ppt` is unsupported.** Save as `.pptx` first.

## Long decks

A 20-page deck reads in well under a second and a few tens of KB. For hundreds of
pages, read with `--slides` in chunks, or run `ppt inspect` first to see which
pages actually carry text.

Full flag reference: `references/api.md`.
