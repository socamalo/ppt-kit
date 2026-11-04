# `ppt` — command reference

Every command accepts `--json` (uniform envelope). Exit codes: `0` ok · `1` usage
· `2` file · `3` not found/ambiguous · `4` write.

Commands that write accept `-o/--out PATH` (write elsewhere instead of in place)
and `--backup` (keep `PATH.bak` before overwriting).

---

## read

```
ppt read FILE [--slides SPEC] [--empty] [--shapes] [--json]
```

`data` = `{file, slide_count, slide_size:{width,height,unit}, slides:[…]}`;
each slide `{index, layout, notes, blocks:[…]}`; each block
`{id, key, kind, text, shape_id, name, shape_type, is_placeholder, editable, left, top, width, height}`.

- `kind`: `text` / `cell` / `notes` / `diagram` / `chart` (pictures are not blocks; see `--shapes`)
- `SPEC`: `1`, `3-5`, `1,4-6`
- `--empty` also returns empty text boxes (useful when filling placeholders)
- `--shapes` attaches every shape on the slide, pictures included, with geometry

## inspect

```
ppt inspect FILE [--slide N] [--json]
```

Without `--slide`: per slide `{index, layout, shape_count, text_blocks, notes}`.
With `--slide N`: every shape (`id/shape_id/shape_type/placeholder_type/left/top/width/height/name`) plus all blocks.

## layouts

```
ppt layouts FILE [--json]
```

`data.layouts[]` = `{index, master, name, placeholders:[{idx,type,name}]}`.
`add-slide --layout` takes either the `index` or the `name`.

## find

```
ppt find FILE QUERY [--regex] [-i] [--slides SPEC] [--json]
```

`data.matches[]` = `{slide, id, key, kind, text}`; `count` is the number of hits.

## set

```
ppt set FILE [--edits PATH|-] [--set REF=TEXT ...] [--slide N]
             [--replace TEXT --with TEXT [--regex] [-i]] [--slides SPEC]
             [--dry-run] [-o OUT] [--backup] [--json]
```

- `--edits`: `{"slide": {"address": "new text"}}`; a value of `"spaceholder"` is skipped (legacy marker)
- `--set`: a single edit; the target slide comes from `--slide` (default 1); repeatable
- `--replace/--with`: find by text and replace every hit; with `--regex` the pattern is a regular expression
- `--dry-run`: report only, write nothing

`data` = `{updated, changes:[{slide,id,ref,kind,old,new}], saved, dry_run}`.

Reference resolution order: `id` → legacy key → `text_index_shapeid` → exact text
(all hits on that slide, each one listed in `changes`). If any reference fails to
resolve the command exits 3, applies nothing, and lists every problem in
`error.detail.errors`.

## new

```
ppt new FILE [--template T.pptx] [--keep N] [--force] [--json]
```

Without a template the deck is 16:9 (13.333×7.5 in). `--template` keeps the
template's masters and layouts and drops its sample slides by default; `--keep N`
keeps the first N as examples. Refuses to overwrite an existing file without
`--force`.

## add-slide

```
ppt add-slide FILE [--layout NAME|INDEX] [--at N] [--json]
```

New slides go at the end unless `--at` says otherwise. `data.index` is the new
page number.

## delete-slide

```
ppt delete-slide FILE --slides SPEC [--json]
```

Refuses to delete every slide; use `ppt new` for an empty deck. `data.removed`
lists the page numbers that went.

## duplicate-slide

```
ppt duplicate-slide FILE --slide N [--at M] [--json]
```

The copy lands right after its source (same as PowerPoint). Pictures, tables,
notes and background come along; relationships and media parts are re-pointed at
the new slide.

## delete-shape

```
ppt delete-shape FILE --slide N --id REF [--id REF ...] [--json]
```

Deletes a whole shape (text box, picture, group, table). `REF` is a block id or a
legacy key; a text reference that matches several blocks deletes all of them —
pass ids when you mean one. A table cell is not a shape: clear it with `ppt set`.
`data.removed[]` = `{id, kind, text}`. Any unresolved reference aborts the batch
(exit 3).

## move-slide

```
ppt move-slide FILE --slide N --to M [--json]
```

## notes

```
ppt notes FILE [--slide N] [--set TEXT | --set-from PATH] [--slides SPEC] [--json]
```

Without `--set`/`--set-from` this reads: with `--slide` it returns that page,
otherwise every non-empty note. With one of them it writes to `--slide` (default
1) or to every page in `--slides`.

## add-textbox

```
ppt add-textbox FILE --slide N (--text TEXT | --text-file PATH)
    [--left --top --width --height] [--font-size F] [--font NAME]
    [--color RRGGBB] [--bold] [--italic] [--align left|center|right|justify]
    [--anchor top|middle|bottom] [--json]
```

Units are inches. Without width/height the box spans the slide width and is 1 in
tall.

## add-image

```
ppt add-image FILE --slide N --image PATH [--left --top --width --height] [--json]
```

Give only one of width/height to keep the original aspect ratio.

## add-table

```
ppt add-table FILE --slide N --data ROWS.json
    [--left --top --width --height] [--font-size F] [--no-header] [--json]
```

`ROWS.json` is a 2-D array: `[["Metric","Value"],["Revenue","100"]]`, optionally
wrapped as `{"rows": [...]}`. Rows may differ in length; missing cells are empty.

## verify

```
ppt verify FILE [--json]
```

`data` = `{loads, slide_count, slide_size, text_blocks, empty_slides, broken}`;
anything suspicious also lands in `warnings`.
