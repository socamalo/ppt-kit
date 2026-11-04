"""Human-readable renderings (the default output when --json is not passed)."""

from __future__ import annotations

from .model import Deck, Slide


def deck_markdown(deck: Deck, *, include_empty: bool = False, notes: bool = True) -> str:
    lines = [
        f"# {deck.path}",
        f"{len(deck.prs.slides)} slides · {deck.size['width']}×{deck.size['height']} in",
        "",
    ]
    for slide in deck.slides:
        subtitle = f" · {slide.layout}" if slide.layout else ""
        lines.append(f"## Slide {slide.index}{subtitle}")
        shown = 0
        for block in slide.blocks:
            if block.kind == "picture":
                continue
            if not block.text.strip() and not include_empty:
                continue
            if not block.text.strip():
                lines.append(f"- `{block.id}` (empty) {block.name}")
                shown += 1
                continue
            text = block.text.replace("\n", " ⏎ ")
            lines.append(f"- `{block.id}` {text}")
            shown += 1
        if shown == 0:
            lines.append("- (this slide has no text)")
        if notes and slide.notes:
            note = slide.notes.replace("\n", " ⏎ ")
            lines.append(f"  - notes: {note}")
        lines.append("")
    return "\n".join(lines).rstrip() + "\n"


def slides_markdown(slides: list[Slide], *, include_empty: bool = False) -> str:
    lines = []
    for slide in slides:
        lines.append(f"## Slide {slide.index} · {slide.layout}")
        for block in slide.blocks:
            if block.kind == "picture":
                continue
            if not block.text.strip() and not include_empty:
                continue
            lines.append(f"- `{block.id}` {block.text.replace(chr(10), ' ⏎ ')}")
        lines.append("")
    return "\n".join(lines).rstrip() + "\n"


def shapes_table(slide=None, shapes: list[dict] | None = None) -> str:
    rows = ["id | shape_id | type | placeholder | position (in) | name", "---|" + "---|" * 5]
    for s in (shapes if shapes is not None else slide.shapes):
        rows.append(
            "{id} | {sid} | {type} | {ph} | {l},{t} {w}×{h} | {name}".format(
                id=s["id"],
                sid=s["shape_id"],
                type=s["shape_type"],
                ph=s["placeholder_type"] or ("yes" if s["is_placeholder"] else ""),
                l=s["left"],
                t=s["top"],
                w=s["width"],
                h=s["height"],
                name=s["name"],
            )
        )
    return "\n".join(rows)


def layouts_table(layouts: list[dict]) -> str:
    rows = ["# | name | placeholders", "---|---|---"]
    for item in layouts:
        ph = ", ".join(f"{p['idx']}:{p['type']}" for p in item["placeholders"])
        rows.append(f"{item['index']} | {item['name']} | {ph}")
    return "\n".join(rows)
