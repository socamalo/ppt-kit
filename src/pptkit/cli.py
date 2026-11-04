"""`ppt` — the command surface an agent drives to CRUD a .pptx file.

Contract (same for every subcommand):

* ``--json`` prints ``{ok, command, file, data, warnings, error}``
* exit codes: 0 ok · 1 usage · 2 file · 3 not found/ambiguous · 4 write
* a write either lands whole or not at all; failures never leave a half deck
"""

from __future__ import annotations

import argparse
import json
import os
import re
import sys
from typing import Any

from . import extract, objects, render, slides as slide_ops, store
from .edit import apply_plan, plan_edits, resolve
from .errors import (
    EXIT_USAGE,
    EXIT_WRITE,
    FileError,
    NotFoundError,
    PptError,
    UsageError,
    dump,
    envelope,
)
from .model import Deck

VERSION = "0.1.0"


# --------------------------------------------------------------------------
# helpers
# --------------------------------------------------------------------------


def _utf8_streams() -> None:
    """Windows consoles default to a legacy code page; a pipe expects UTF-8."""
    for stream in (sys.stdout, sys.stderr):
        try:
            stream.reconfigure(encoding="utf-8", errors="replace")
        except Exception:
            pass


def parse_slides(spec: str | None) -> list[int] | None:
    if not spec:
        return None
    out: list[int] = []
    for chunk in str(spec).split(","):
        chunk = chunk.strip()
        if not chunk:
            continue
        if "-" in chunk:
            a, _, b = chunk.partition("-")
            start, end = int(a), int(b)
            step = 1 if end >= start else -1
            out.extend(range(start, end + step, step))
        else:
            out.append(int(chunk))
    return sorted(set(out))


def _read_json(path: str) -> Any:
    if path == "-":
        raw = sys.stdin.read()
    else:
        if not os.path.exists(path):
            raise FileError(f"JSON file not found: {path}")
        with open(path, encoding="utf-8-sig") as fh:
            raw = fh.read()
    try:
        return json.loads(raw)
    except json.JSONDecodeError as e:
        raise UsageError(f"cannot parse JSON ({path}): {e}") from e


def _read_text(path: str) -> str:
    if path == "-":
        return sys.stdin.read()
    if not os.path.exists(path):
        raise FileError(f"file not found: {path}")
    with open(path, encoding="utf-8-sig") as fh:
        return fh.read()


def _open(path: str) -> Deck:
    return extract.open_deck(path)


def _reopen(path: str) -> Deck:
    return extract.open_deck(path)


def _split_pairs(pairs: list[str] | None) -> dict[str, str]:
    out: dict[str, str] = {}
    for item in pairs or []:
        if "=" not in item:
            raise UsageError(f"--set expects ADDRESS=TEXT (got {item!r})")
        key, _, value = item.partition("=")
        out[key.strip()] = value
    return out


# --------------------------------------------------------------------------
# handlers: each returns the `data` part of the envelope
# --------------------------------------------------------------------------


def cmd_read(args) -> dict:
    deck = _open(args.file)
    wanted = parse_slides(args.slides)
    selected = deck.slides if not wanted else [deck.slide(i) for i in wanted]
    data = {**deck.summary(), "slides": []}
    for slide in selected:
        data["slides"].append(
            slide.to_dict(include_empty=args.empty, include_shapes=args.shapes)
        )
    return data


def cmd_inspect(args) -> dict:
    deck = _open(args.file)
    if args.slide:
        slide = deck.slide(args.slide)
        return {
            "slide": slide.index,
            "layout": slide.layout,
            "size": deck.size,
            "shapes": slide.shapes,
            "blocks": [b.to_dict() for b in slide.blocks],
        }
    return {
        "size": deck.size,
        "slides": [
            {
                "index": s.index,
                "layout": s.layout,
                "shape_count": len(s.shapes),
                "text_blocks": len(s.text_blocks()),
                "notes": bool(s.notes),
            }
            for s in deck.slides
        ],
    }


def cmd_layouts(args) -> dict:
    deck = _open(args.file)
    layouts = slide_ops.list_layouts(deck.prs)
    return {"size": deck.size, "layouts": layouts}


def cmd_find(args) -> dict:
    deck = _open(args.file)
    hits = extract.find_blocks(
        deck,
        args.query,
        regex=args.regex,
        ignore_case=args.ignore_case,
        slides=parse_slides(args.slides),
    )
    return {
        "query": args.query,
        "count": len(hits),
        "matches": [
            {
                "slide": slide.index,
                "id": block.id,
                "key": block.key,
                "kind": block.kind,
                "text": block.text,
            }
            for slide, block in hits
        ],
    }


def cmd_set(args) -> dict:
    edits: dict[str, dict[str, str]] = {}
    if args.edits:
        payload = _read_json(args.edits)
        if isinstance(payload, list):
            raise UsageError(
                'the edits file must be an object like {"1": {"b6": "new text"}}; convert arrays first'
            )
        if not isinstance(payload, dict):
            raise UsageError("the edits file must be a JSON object")
        for slide_key, pairs in payload.items():
            if slide_key in ("meta", "_meta"):
                continue
            edits.setdefault(str(slide_key), {}).update(pairs or {})

    direct = _split_pairs(args.set)
    if direct:
        # An id is only unique *within* a slide, so --set needs one explicit
        # target slide; defaulting to "every slide" would silently fan out.
        target_slides = parse_slides(args.slides) or [args.slide or 1]
        for ref, value in direct.items():
            for slide_no in target_slides:
                edits.setdefault(str(slide_no), {})[ref] = value

    if args.replace is not None:
        if args.with_text is None:
            raise UsageError("--replace needs --with")
        deck = _open(args.file)
        hits = extract.find_blocks(
            deck,
            args.replace,
            regex=args.regex,
            ignore_case=args.ignore_case,
            slides=parse_slides(args.slides),
        )
        if not hits:
            raise NotFoundError(f"no match for '{args.replace}'")
        flags = re.IGNORECASE if args.ignore_case else 0
        for slide, block in hits:
            # replace *inside* the block: rewriting the whole block would throw
            # away the rest of the line around the match
            if args.regex:
                new_value = re.sub(args.replace, args.with_text, block.text, flags=flags)
            elif args.ignore_case:
                new_value = re.sub(
                    re.escape(args.replace), args.with_text, block.text, flags=flags
                )
            else:
                new_value = block.text.replace(args.replace, args.with_text)
            edits.setdefault(str(slide.index), {})[block.id] = new_value

    if not edits:
        raise UsageError(
            "nothing to change: pass --edits FILE, --set ADDRESS=TEXT, or --replace/--with"
        )

    deck = _open(args.file)
    plan, errors = plan_edits(deck, edits, slides=parse_slides(args.slides))
    if errors:
        raise NotFoundError(
            f"{len(errors)} reference(s) could not be resolved; nothing was applied",
            detail={"errors": errors},
        )
    if not plan:
        return {
            "updated": 0,
            "changes": [],
            "saved": None,
            "dry_run": bool(args.dry_run),
            "note": "no block matched the given references",
        }

    if args.dry_run:
        return {
            "updated": len(plan),
            "changes": [
                {
                    "slide": item.slide,
                    "id": item.block.id,
                    "ref": item.ref,
                    "kind": item.block.kind,
                    "old": item.block.text,
                    "new": item.new,
                }
                for item in plan
            ],
            "saved": None,
            "dry_run": True,
        }

    changes = apply_plan(plan)
    saved = store.save(deck, args.out, backup=args.backup)
    return {
        "updated": len(changes),
        "changes": [c.to_dict() for c in changes],
        "saved": saved,
        "dry_run": False,
    }


def cmd_new(args) -> dict:
    if os.path.exists(args.file) and not args.force:
        raise UsageError(
            f"{args.file} already exists; add --force to overwrite or pick another path",
            detail={"exists": True},
        )
    prs = slide_ops.new_presentation(args.template, keep=args.keep or 0)
    deck = Deck(path=args.file, prs=prs)
    saved = store.save(deck, args.file)
    fresh = _reopen(saved)
    return {
        "created": saved,
        "template": args.template,
        "slide_count": len(fresh.slides),
        "slide_size": fresh.size,
        "layouts": len(slide_ops.list_layouts(fresh.prs)),
    }


def _mutate(args, action) -> dict:
    deck = _open(args.file)
    detail = action(deck)
    saved = store.save(deck, args.out, backup=args.backup)
    fresh = _reopen(saved)
    return {"saved": saved, "slide_count": len(fresh.slides), **detail}


def cmd_add_slide(args) -> dict:
    def action(deck):
        index = slide_ops.add_slide(deck.prs, args.layout, args.at)
        return {"index": index, "layout": _reopen_layout(deck, index)}

    return _mutate(args, action)


def _reopen_layout(deck, index: int) -> str:
    try:
        return deck.prs.slides[index - 1].slide_layout.name
    except Exception:
        return ""


def cmd_delete_slide(args) -> dict:
    wanted = parse_slides(args.slides)
    if not wanted:
        raise UsageError("--slides is required, e.g. --slides 3 or 3,5-6")

    def action(deck):
        removed = slide_ops.delete_slides(deck.prs, wanted)
        return {"removed": removed}

    return _mutate(args, action)


def cmd_duplicate_slide(args) -> dict:
    def action(deck):
        index = slide_ops.duplicate_slide(deck.prs, args.slide, args.at)
        return {"index": index}

    return _mutate(args, action)


def cmd_move_slide(args) -> dict:
    def action(deck):
        slide_ops.move_slide(deck.prs, args.slide, args.to)
        return {"moved": args.slide, "to": args.to}

    return _mutate(args, action)


def cmd_notes(args) -> dict:
    if args.set is None and not args.set_from:
        deck = _open(args.file)
        if args.slide:
            slide = deck.slide(args.slide)
            return {"slide": slide.index, "notes": slide.notes}
        return {"notes": {str(s.index): s.notes for s in deck.slides if s.notes}}

    text = args.set if args.set is not None else _read_text(args.set_from)
    deck = _open(args.file)
    target = args.slide or 1
    slides = [deck.slide(i) for i in (parse_slides(args.slides) or [target])]
    for slide in slides:
        deck.prs.slides[slide.index - 1].notes_slide.notes_text_frame.text = text
    saved = store.save(deck, args.out, backup=args.backup)
    return {"saved": saved, "slides": [s.index for s in slides], "notes": text}


def cmd_add_textbox(args) -> dict:
    text = args.text if args.text is not None else _read_text(args.text_file)

    def action(d):
        slide = d.prs.slides[args.slide - 1]
        shape = objects.add_textbox(
            slide,
            text,
            left=args.left,
            top=args.top,
            width=args.width,
            height=args.height,
            font_size=args.font_size,
            bold=args.bold,
            italic=args.italic,
            color=args.color,
            font=args.font,
            align=args.align,
            anchor=args.anchor,
        )
        return {"shape_id": shape.shape_id}

    return _mutate(args, action)


def cmd_add_image(args) -> dict:
    def action(deck):
        slide = deck.prs.slides[args.slide - 1]
        shape = objects.add_picture(
            slide,
            args.image,
            left=args.left,
            top=args.top,
            width=args.width,
            height=args.height,
        )
        return {"shape_id": shape.shape_id}

    return _mutate(args, action)


def cmd_add_table(args) -> dict:
    rows = _read_json(args.data)
    if isinstance(rows, dict):
        rows = rows.get("rows") or rows.get("data")
    if not isinstance(rows, list):
        raise UsageError('table data must be a 2-D array, e.g. [["Item","Value"],["A","1"]]')

    def action(deck):
        slide = deck.prs.slides[args.slide - 1]
        shape = objects.add_table(
            slide,
            rows,
            left=args.left,
            top=args.top,
            width=args.width,
            height=args.height,
            header=args.header,
            font_size=args.font_size,
        )
        return {"shape_id": shape.shape_id}

    return _mutate(args, action)


def cmd_delete_shape(args) -> dict:
    deck = _open(args.file)
    slide = deck.slide(args.slide)
    live_slide = deck.prs.slides[slide.index - 1]
    refs = list(args.id or [])
    if not refs:
        raise UsageError("--id is required, e.g. --id b9 (repeatable)")

    resolved = []
    errors = []
    for ref in refs:
        try:
            blocks = resolve(slide, ref)
        except NotFoundError as e:
            errors.append({"ref": ref, "error": e.message})
            continue
        for block in blocks:
            if block.kind == "cell":
                errors.append(
                    {
                        "ref": ref,
                        "id": block.id,
                        "error": "a table cell is not a shape; clear it with `ppt set` instead",
                    }
                )
                continue
            shape = extract.locate_shape(live_slide, block.id)
            if shape is None:
                errors.append({"ref": ref, "id": block.id, "error": "shape could not be located"})
                continue
            resolved.append((block, shape))

    if errors:
        raise NotFoundError(
            f"{len(errors)} reference(s) could not be deleted; nothing was applied",
            detail={"errors": errors},
        )

    removed = []
    seen = set()
    for block, shape in resolved:
        if block.id in seen:
            continue
        seen.add(block.id)
        shape._element.getparent().remove(shape._element)
        removed.append({"id": block.id, "kind": block.kind, "text": block.text})

    saved = store.save(deck, args.out, backup=args.backup)
    return {"saved": saved, "removed": removed}


def cmd_verify(args) -> dict:
    deck = _open(args.file)
    warnings = []
    empty = [s.index for s in deck.slides if not s.text_blocks()]
    if empty:
        warnings.append(f"{len(empty)} slide(s) carry no text at all: {empty}")
    broken = []
    for slide in deck.slides:
        for block in slide.blocks:
            if block.kind in ("text", "cell") and block.target is None:
                broken.append({"slide": slide.index, "id": block.id})
    return {
        "loads": True,
        "slide_count": len(deck.slides),
        "slide_size": deck.size,
        "text_blocks": sum(len(s.text_blocks()) for s in deck.slides),
        "empty_slides": empty,
        "broken": broken,
        "warnings": warnings,
    }


HANDLERS = {
    "read": cmd_read,
    "inspect": cmd_inspect,
    "layouts": cmd_layouts,
    "find": cmd_find,
    "set": cmd_set,
    "new": cmd_new,
    "add-slide": cmd_add_slide,
    "delete-slide": cmd_delete_slide,
    "delete-shape": cmd_delete_shape,
    "duplicate-slide": cmd_duplicate_slide,
    "move-slide": cmd_move_slide,
    "notes": cmd_notes,
    "add-textbox": cmd_add_textbox,
    "add-image": cmd_add_image,
    "add-table": cmd_add_table,
    "verify": cmd_verify,
}


# --------------------------------------------------------------------------
# human output
# --------------------------------------------------------------------------


def _human(args, data, command: str, warnings: list[str] | None = None) -> str:
    if command == "read":
        lines = [
            f"{data['file']}",
            f"{data['slide_count']} slides · {data['slide_size']['width']}×{data['slide_size']['height']} in",
        ]
        for slide in data["slides"]:
            lines.append("")
            lines.append(f"## Slide {slide['index']} · {slide['layout']}")
            for block in slide["blocks"]:
                text = (block["text"] or "(empty)").replace("\n", " ⏎ ")
                lines.append(f"- `{block['id']}` {text}")
            if slide.get("notes"):
                lines.append(f"  - notes: {slide['notes'].replace(chr(10), ' ⏎ ')}")
        return "\n".join(lines)

    if command == "layouts":
        return render.layouts_table(data["layouts"])

    if command == "inspect":
        if "shapes" in data:
            return f"Slide {data['slide']} · {data['layout']}\n\n" + render.shapes_table(
                None, data["shapes"]
            )
        rows = ["# | shapes | text blocks | layout", "---|---|---|---"]
        for s in data["slides"]:
            rows.append(f"{s['index']} | {s['shape_count']} | {s['text_blocks']} | {s['layout']}")
        return "\n".join(rows)

    if command == "find":
        if not data["matches"]:
            return f"no match for {data['query']}"
        return "\n".join(
            f"[{m['slide']}] `{m['id']}` {m['text'].replace(chr(10), ' ⏎ ')}"
            for m in data["matches"]
        )

    if command == "set":
        if data.get("dry_run"):
            head = f"dry run: would change {data['updated']} block(s), nothing written"
        else:
            head = f"changed {data['updated']} block(s) · {data.get('saved')}"
        body = [
            f"  [{c['slide']}] {c['id']}: {c['old'][:40]!r} -> {c['new'][:40]!r}"
            for c in data.get("changes", [])
        ]
        return "\n".join([head] + body)

    if command == "delete-shape":
        lines = [f"removed {len(data['removed'])} shape(s) · {data.get('saved')}"]
        lines.extend(f"  {r['id']} ({r['kind']}) {r['text'][:40]!r}" for r in data["removed"])
        return "\n".join(lines)

    if command == "verify":
        lines = [
            f"opens fine · {data['slide_count']} slides · {data['text_blocks']} text blocks",
        ]
        lines.extend(f"! {w}" for w in (warnings or []))
        return "\n".join(lines)

    # default: readable key/value summary
    return "\n".join(f"{k}: {v}" for k, v in data.items())


# --------------------------------------------------------------------------
# parser
# --------------------------------------------------------------------------


class _Parser(argparse.ArgumentParser):
    """argparse exits 2 on bad usage; the contract says usage is 1."""

    def error(self, message):
        self.print_usage(sys.stderr)
        raise UsageError(message)


def build_parser() -> argparse.ArgumentParser:
    p = _Parser(
        prog="ppt",
        description=(
            "Read and write a .pptx by file: text as addressable blocks, "
            "edit in place, add/remove slides, notes, pictures and tables."
        ),
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog=(
            "examples:\n"
            "  ppt read deck.pptx --json\n"
            "  ppt set deck.pptx --edits edits.json --json\n"
            "  ppt find deck.pptx Revenue --json\n"
            "  ppt add-slide deck.pptx --layout 3 --at 2 --json\n"
            "  ppt delete-slide deck.pptx --slides 4 --json\n"
        ),
    )
    p.add_argument("--version", action="version", version=f"ppt-kit {VERSION}")
    sub = p.add_subparsers(dest="command", metavar="COMMAND")

    def common(sp, *, needs_out=True):
        sp.add_argument("file", help="the .pptx to work on")
        sp.add_argument("--json", action="store_true", help="print the uniform JSON envelope")
        if needs_out:
            sp.add_argument("-o", "--out", help="write to another file (default: in place)")
            sp.add_argument("--backup", action="store_true", help="keep a .bak before overwriting")

    sp = sub.add_parser("read", help="dump every text block, with ids")
    common(sp)
    sp.add_argument("--slides", help="page range, e.g. 1,3-5 (default: all)")
    sp.add_argument("--empty", action="store_true", help="include empty text boxes")
    sp.add_argument("--shapes", action="store_true", help="attach geometry for every shape")
    sp.set_defaults(handler=cmd_read)

    sp = sub.add_parser("inspect", help="shapes, layouts, geometry")
    common(sp, needs_out=False)
    sp.add_argument("--slide", type=int, help="only this slide")
    sp.set_defaults(handler=cmd_inspect)

    sp = sub.add_parser("layouts", help="list layouts and their placeholders")
    common(sp, needs_out=False)
    sp.set_defaults(handler=cmd_layouts)

    sp = sub.add_parser("find", help="locate blocks by keyword or regex")
    common(sp, needs_out=False)
    sp.add_argument("query", help="keyword or regular expression")
    sp.add_argument("--regex", action="store_true")
    sp.add_argument("-i", "--ignore-case", action="store_true")
    sp.add_argument("--slides", help="limit to a page range, e.g. 1,3-5")
    sp.set_defaults(handler=cmd_find)

    sp = sub.add_parser("set", help="batch text edits (all or nothing)")
    common(sp)
    sp.add_argument("--edits", help='JSON file, or - for stdin: {"1": {"b6": "new text"}}')
    sp.add_argument("--set", action="append", metavar="REF=TEXT", help="one edit, repeatable")
    sp.add_argument("--slide", type=int, help="slide for --set (default 1)")
    sp.add_argument("--replace", help="search for this text (with --with)")
    sp.add_argument("--with", dest="with_text", help="replacement text")
    sp.add_argument("--regex", action="store_true", help="treat --replace as a regex")
    sp.add_argument("-i", "--ignore-case", action="store_true")
    sp.add_argument("--slides", help="limit to a page range (for --edits / --replace)")
    sp.add_argument("--dry-run", action="store_true", help="report changes without writing")
    sp.set_defaults(handler=cmd_set)

    sp = sub.add_parser("new", help="create a deck, optionally from a template")
    common(sp, needs_out=False)
    sp.add_argument("--template", help="template pptx; keeps its masters and layouts")
    sp.add_argument("--keep", type=int, default=0, help="keep the first N template slides")
    sp.add_argument("--force", action="store_true", help="overwrite an existing file")
    sp.set_defaults(handler=cmd_new)

    sp = sub.add_parser("add-slide", help="add a slide")
    common(sp)
    sp.add_argument("--layout", help="layout name or index (see `layouts`)")
    sp.add_argument("--at", type=int, help="insert at this position (default: end)")
    sp.set_defaults(handler=cmd_add_slide)

    sp = sub.add_parser("delete-slide", help="remove slides")
    common(sp)
    sp.add_argument("--slides", required=True, help="e.g. 3 or 3,5-6")
    sp.set_defaults(handler=cmd_delete_slide)

    sp = sub.add_parser("delete-shape", help="remove one shape (text box, picture, group)")
    common(sp)
    sp.add_argument("--slide", type=int, required=True)
    sp.add_argument("--id", action="append", help="block id or legacy key, repeatable")
    sp.set_defaults(handler=cmd_delete_shape)

    sp = sub.add_parser("duplicate-slide", help="copy a slide, pictures and tables included")
    common(sp)
    sp.add_argument("--slide", type=int, required=True)
    sp.add_argument("--at", type=int, help="place the copy at this position")
    sp.set_defaults(handler=cmd_duplicate_slide)

    sp = sub.add_parser("move-slide", help="reorder slides")
    common(sp)
    sp.add_argument("--slide", type=int, required=True)
    sp.add_argument("--to", type=int, required=True)
    sp.set_defaults(handler=cmd_move_slide)

    sp = sub.add_parser("notes", help="read or write speaker notes")
    common(sp)
    sp.add_argument("--slide", type=int, help="which slide (default 1; reads return all)")
    sp.add_argument("--set", help="notes text")
    sp.add_argument("--set-from", help="read notes text from a file (keeps non-ASCII out of argv)")
    sp.add_argument("--slides", help="write to several slides")
    sp.set_defaults(handler=cmd_notes)

    sp = sub.add_parser("add-textbox", help="add a text box")
    common(sp)
    sp.add_argument("--slide", type=int, required=True)
    sp.add_argument("--text")
    sp.add_argument("--text-file", help="read text from a file (keeps non-ASCII out of argv)")
    sp.add_argument("--left", type=float)
    sp.add_argument("--top", type=float)
    sp.add_argument("--width", type=float)
    sp.add_argument("--height", type=float)
    sp.add_argument("--font-size", type=float)
    sp.add_argument("--font")
    sp.add_argument("--color", help="6-digit hex, e.g. 1F4E79")
    sp.add_argument("--bold", action="store_true")
    sp.add_argument("--italic", action="store_true")
    sp.add_argument("--align", choices=["left", "center", "right", "justify"])
    sp.add_argument("--anchor", choices=["top", "middle", "bottom"])
    sp.set_defaults(handler=cmd_add_textbox)

    sp = sub.add_parser("add-image", help="insert a picture")
    common(sp)
    sp.add_argument("--slide", type=int, required=True)
    sp.add_argument("--image", required=True)
    sp.add_argument("--left", type=float)
    sp.add_argument("--top", type=float)
    sp.add_argument("--width", type=float)
    sp.add_argument("--height", type=float)
    sp.set_defaults(handler=cmd_add_image)

    sp = sub.add_parser("add-table", help="insert a table from a JSON 2-D array")
    common(sp)
    sp.add_argument("--slide", type=int, required=True)
    sp.add_argument("--data", required=True, help='JSON file: [["Header","Value"],["A","1"]]')
    sp.add_argument("--left", type=float, default=0.5)
    sp.add_argument("--top", type=float, default=1.5)
    sp.add_argument("--width", type=float, default=9.0)
    sp.add_argument("--height", type=float)
    sp.add_argument("--font-size", type=float)
    sp.add_argument("--no-header", dest="header", action="store_false")
    sp.set_defaults(handler=cmd_add_table, header=True)

    sp = sub.add_parser("verify", help="check the file opens and has no empty slides")
    common(sp, needs_out=False)
    sp.set_defaults(handler=cmd_verify)

    return p


def main(argv: list[str] | None = None) -> int:
    _utf8_streams()
    parser = build_parser()
    try:
        args = parser.parse_args(argv)
    except UsageError as e:
        print(f"usage error: {e.message}", file=sys.stderr)
        return e.exit_code
    if not getattr(args, "command", None):
        parser.print_help()
        return EXIT_USAGE

    command = args.command
    handler = HANDLERS[command]
    file_hint = getattr(args, "file", None)
    warnings: list[str] = []

    try:
        data = handler(args)
        warnings = list(data.pop("warnings", []) or [])
        payload = envelope(command, ok=True, file=file_hint, data=data, warnings=warnings)
        code = 0
    except PptError as e:
        payload = envelope(
            command,
            ok=False,
            file=file_hint,
            error={"kind": e.kind, "message": e.message, "detail": e.detail},
        )
        code = e.exit_code
    except Exception as e:  # noqa: BLE001
        payload = envelope(
            command,
            ok=False,
            file=file_hint,
            error={"kind": "internal", "message": str(e), "detail": None},
        )
        code = EXIT_WRITE

    if args.json:
        print(dump(payload))
    elif code == 0:
        print(_human(args, payload["data"], command, warnings))
    else:
        err = payload["error"]
        print(f"failed: {err['message']}", file=sys.stderr)
        if err.get("detail"):
            print(json.dumps(err["detail"], ensure_ascii=False, indent=2), file=sys.stderr)
    return code


if __name__ == "__main__":
    raise SystemExit(main())
