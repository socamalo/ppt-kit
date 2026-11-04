# ppt-kit

**Edit PowerPoint files the way you edit code — by address, from the command line.**

`ppt-kit` reads a `.pptx` into addressable text blocks and writes changes back in
place. No PowerPoint install, no COM, no window to keep open. Point it at a file,
ask for the blocks, change the ones you mean, and it saves an OOXML package your
users can still open.

```console
$ ppt read deck.pptx --json | jq -r '.data.slides[0].blocks[] | "\(.id)  \(.text)"'
b6   [Full title:]{10-15 words}
b9   Q3 results

$ ppt set deck.pptx --set 'b6=Q3 results: revenue up 20%' --json
{
  "ok": true,
  "command": "set",
  "file": "deck.pptx",
  "data": {
    "updated": 1,
    "changes": [
      {
        "slide": 1,
        "id": "b6",
        "old": "[Full title:]{10-15 words}",
        "new": "Q3 results: revenue up 20%"
      }
    ],
    "saved": "/work/deck.pptx"
  },
  "warnings": [],
  "error": null
}
```

## Why this exists

Every "AI edits my slides" demo dies on the same problem: **the same words appear
several times on one slide**. A template has five boxes labelled "Subtitle".
Find-and-replace changes all five; a naive library binds to the first. You wanted
one.

`ppt-kit` gives every block an **address that names a place, not a string**:

| Address | Points at |
|---|---|
| `b6` | the shape whose OOXML id is 6 |
| `b12/b5` | shape 5 inside group 12 |
| `b9.r2c3` | row 2, column 3 of table 9 |
| `b7/dgm3` | SmartArt node 3 (read-only) |
| `notes` | speaker notes |

Five boxes reading "Subtitle" are five addresses. Editing `b12/b5` moves that one
and nothing else. The address survives text edits because it is derived from the
shape id, not the content.

## Install

```bash
git clone https://github.com/socamalo/ppt-kit
uv tool install --from ./ppt-kit ppt-kit
ppt --version
```

Needs Python 3.10+. The only runtime dependency is `python-pptx`.

## A five-minute tour

```bash
ppt read deck.pptx --json                # every block with its id
ppt read deck.pptx --slides 1,3-5        # just some pages
ppt find deck.pptx Revenue --json        # locate by keyword or regex
ppt inspect deck.pptx --slide 3 --json   # shapes, geometry, placeholders
ppt layouts deck.pptx --json             # layouts available for new slides
```

Change things — one block, or a whole batch:

```bash
# one edit, in place
ppt set deck.pptx --set 'b6=New title' --json

# a batch from a file: {"slide": {"address": "new text"}}
ppt set deck.pptx --edits edits.json --json

# search and replace across the deck, regex allowed
ppt set deck.pptx --replace 'Q([1-4])' --with 'Quarter \1' --regex --json

# show what would change, write nothing
ppt set deck.pptx --edits edits.json --dry-run --json
```

Slides:

```bash
ppt new out.pptx --template brand.pptx      # start from a template
ppt add-slide deck.pptx --layout 3 --at 2   # add a slide from a layout
ppt duplicate-slide deck.pptx --slide 4     # copy one, pictures and tables included
ppt move-slide deck.pptx --slide 5 --to 1
ppt delete-slide deck.pptx --slides 3,5-6
ppt delete-shape deck.pptx --slide 2 --id b9
```

Content:

```bash
ppt notes deck.pptx --slide 2 --set "speaker cue"
ppt add-textbox deck.pptx --slide 2 --text "Big idea" --left 1 --top 1 --font-size 32 --bold
ppt add-image deck.pptx --slide 2 --image chart.png --left 6 --top 1 --width 5
ppt add-table deck.pptx --slide 2 --data rows.json
```

Every command takes `--json` (uniform envelope) and `-o out.pptx` (write
elsewhere instead of in place).

## What you can rely on

**One envelope.** `{ok, command, file, data, warnings, error}` — branch on `ok`,
read `data`, never parse prose.

**Fixed exit codes.** `0` ok · `1` usage · `2` file · `3` not found/ambiguous · `4` write.

**All or nothing.** A batch with one bad reference applies nothing and reports
every problem: no half-edited deck to reason about.

**Atomic saves.** Write to a temp file, re-open it to prove it is valid, then
swap. A crash, a lock, a corrupt write never eats the original.

**Formatting survives.** Text is written into the existing run, not re-created,
so bold words, colours, bullets and paragraph indents stay. Newlines grow or
shrink the paragraph list by cloning what is already there.

## What it will not do

- **SmartArt and chart text are read-only.** They live in a separate diagram
  part; edit them in PowerPoint.
- **Theme swap is out of scope.** Moving a slide master between packages is a
  different job — build the new deck on the template instead.
- **No model inside.** Translation, summarising and copywriting are yours:
  `read`, think, `set`.
- **Legacy `.ppt` unsupported.** Save as `.pptx` first.

## Library use

The CLI is a thin shell over a small library:

```python
from pptkit.extract import open_deck
from pptkit.edit import plan_edits, apply_plan
from pptkit.store import save

deck = open_deck("deck.pptx")
plan, errors = plan_edits(deck, {"1": {"b6": "New title"}})
if not errors:
    apply_plan(plan)
    save(deck)
```

## Tests

```bash
uv run --with pytest --with-editable . pytest -q
```

The fixtures are built to be ambiguous on purpose — identical strings in three
text boxes, inside and outside a group, in two cells of one table, on two
slides. The suite asserts the invariant that matters:

> every address resolves to exactly one block

…for both the stable id and the legacy `text_index_shapeid` key.

## Agent skill

`SKILL.md` is written for coding agents: when to reach for the CLI, how to read
addresses, the JSON contract, and the traps. `references/api.md` has the full
flag reference.

## License

MIT.
