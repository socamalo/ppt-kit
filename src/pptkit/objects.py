"""Add non-slide objects: text boxes, pictures, tables."""

from __future__ import annotations

from pptx.dml.color import RGBColor
from pptx.enum.text import MSO_ANCHOR, PP_ALIGN
from pptx.util import Inches, Pt

from .errors import FileError, UsageError

ALIGNMENTS = {
    "left": PP_ALIGN.LEFT,
    "center": PP_ALIGN.CENTER,
    "centre": PP_ALIGN.CENTER,
    "right": PP_ALIGN.RIGHT,
    "justify": PP_ALIGN.JUSTIFY,
}

ANCHORS = {
    "top": MSO_ANCHOR.TOP,
    "middle": MSO_ANCHOR.MIDDLE,
    "bottom": MSO_ANCHOR.BOTTOM,
}


def _rgb(value: str) -> RGBColor:
    text = str(value).strip().lstrip("#")
    if len(text) != 6:
        raise UsageError(f"colours are 6-digit hex, e.g. 1F4E79 (got {value!r})")
    try:
        return RGBColor.from_string(text.upper())
    except Exception as e:  # noqa: BLE001
        raise UsageError(f"not a valid colour: {value!r}") from e


def add_textbox(
    slide,
    text: str,
    *,
    left: float | None = None,
    top: float | None = None,
    width: float | None = None,
    height: float | None = None,
    font_size: float | None = None,
    bold: bool = False,
    italic: bool = False,
    color: str | None = None,
    font: str | None = None,
    align: str | None = None,
    anchor: str | None = None,
    wrap: bool = True,
) -> object:
    sw = slide.part.package.presentation_part.presentation.slide_width or Inches(13.333)
    box_left = Inches(left) if left is not None else 0
    box_top = Inches(top) if top is not None else 0
    box_width = Inches(width) if width is not None else sw
    box_height = Inches(height) if height is not None else Inches(1.0)

    shape = slide.shapes.add_textbox(box_left, box_top, box_width, box_height)
    tf = shape.text_frame
    tf.word_wrap = wrap
    tf.text = text
    if anchor:
        key = anchor.lower()
        if key not in ANCHORS:
            raise UsageError(f"anchor must be one of {sorted(ANCHORS)}")
        tf.vertical_anchor = ANCHORS[key]

    for paragraph in tf.paragraphs:
        if align:
            key = align.lower()
            if key not in ALIGNMENTS:
                raise UsageError(f"align must be one of {sorted(ALIGNMENTS)}")
            paragraph.alignment = ALIGNMENTS[key]
        # an empty box has no runs, so styling would have nothing to land on
        # and text written later (`ppt set`) would come back unstyled
        if not paragraph.runs and not paragraph.text:
            paragraph.add_run()
        for run in paragraph.runs:
            if font_size:
                run.font.size = Pt(font_size)
            if bold:
                run.font.bold = True
            if italic:
                run.font.italic = True
            if color:
                run.font.color.rgb = _rgb(color)
            if font:
                run.font.name = font
    return shape


def add_picture(
    slide,
    image: str,
    *,
    left: float | None = None,
    top: float | None = None,
    width: float | None = None,
    height: float | None = None,
):
    import os

    if not os.path.exists(image):
        raise FileError(f"image not found: {image}")
    kwargs = {}
    if left is not None:
        kwargs["left"] = Inches(left)
    if top is not None:
        kwargs["top"] = Inches(top)
    if width is not None:
        kwargs["width"] = Inches(width)
    if height is not None:
        kwargs["height"] = Inches(height)
    try:
        return slide.shapes.add_picture(image, **kwargs)
    except Exception as e:  # noqa: BLE001
        raise FileError(f"could not insert the image: {e}") from e


def add_table(
    slide,
    rows: list[list[str]],
    *,
    left: float = 0.5,
    top: float = 1.5,
    width: float = 9.0,
    height: float | None = None,
    header: bool = True,
    font_size: float | None = None,
):
    if not rows or not rows[0]:
        raise UsageError("data needs at least one row and one column")
    n_rows = len(rows)
    n_cols = max(len(r) for r in rows)
    if height is None:
        height = 0.4 * n_rows
    frame = slide.shapes.add_table(
        n_rows, n_cols, Inches(left), Inches(top), Inches(width), Inches(height)
    )
    table = frame.table
    for r in range(n_rows):
        for c in range(n_cols):
            value = rows[r][c] if c < len(rows[r]) else ""
            cell = table.cell(r, c)
            cell.text = "" if value is None else str(value)
            if font_size:
                for paragraph in cell.text_frame.paragraphs:
                    for run in paragraph.runs:
                        run.font.size = Pt(font_size)
                        if header and r == 0:
                            run.font.bold = True
    if header:
        for c in range(n_cols):
            for paragraph in table.cell(0, c).text_frame.paragraphs:
                for run in paragraph.runs:
                    run.font.bold = True
    return frame
