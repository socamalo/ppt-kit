"""Build the fixture decks the tests run against.

The theme of these fixtures is *ambiguity*: the same string living in several
places at once — three identical text boxes on one slide, the same wording in a
group and outside it, table cells repeating a label, the same title on two
slides. Landing the edit on the right one, and only that one, is the whole job.
"""

from __future__ import annotations

import base64
import os

from lxml import etree
from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.util import Inches, Pt

A = "http://schemas.openxmlformats.org/drawingml/2006/main"
P = "http://schemas.openxmlformats.org/presentationml/2006/main"

# 1x1 red PNG
PNG_1PX = base64.b64decode(
    "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8z8BQDwAEhQGAhKmMIQAAAABJRU5ErkJggg=="
)


def _blank(prs):
    return prs.slides.add_slide(prs.slide_layouts[6])


def _textbox(slide, text, *, left=1.0, top=1.0, bold=False, size=18):
    box = slide.shapes.add_textbox(Inches(left), Inches(top), Inches(4), Inches(0.8))
    frame = box.text_frame
    frame.text = text
    for paragraph in frame.paragraphs:
        for run in paragraph.runs:
            run.font.size = Pt(size)
            run.font.bold = bold
    return box


def _group_with_two_texts(slide, texts, group_id=900):
    """A real <p:grpSp> containing two text boxes — python-pptx has no API."""
    tree = slide.shapes._spTree
    group = etree.SubElement(tree, f"{{{P}}}grpSp")
    etree.SubElement(group, f"{{{P}}}nvGrpSpPr")
    c_nv = etree.SubElement(group.find(f"{{{P}}}nvGrpSpPr"), f"{{{P}}}cNvPr")
    c_nv.set("id", str(group_id))
    c_nv.set("name", f"Group {group_id}")
    etree.SubElement(group.find(f"{{{P}}}nvGrpSpPr"), f"{{{P}}}cNvGrpSpPr")
    etree.SubElement(group.find(f"{{{P}}}nvGrpSpPr"), f"{{{P}}}nvPr")
    grp_pr = etree.SubElement(group, f"{{{P}}}grpSpPr")
    xfrm = etree.SubElement(grp_pr, f"{{{A}}}xfrm")
    for tag, attrs in (
        ("off", {"x": "0", "y": "0"}),
        ("ext", {"cx": "0", "cy": "0"}),
        ("chOff", {"x": "0", "y": "0"}),
        ("chExt", {"cx": "0", "cy": "0"}),
    ):
        node = etree.SubElement(xfrm, f"{{{A}}}{tag}")
        for k, v in attrs.items():
            node.set(k, v)

    for n, text in enumerate(texts, 1):
        sp = etree.SubElement(group, f"{{{P}}}sp")
        nv = etree.SubElement(sp, f"{{{P}}}nvSpPr")
        cnv = etree.SubElement(nv, f"{{{P}}}cNvPr")
        cnv.set("id", str(group_id + n))
        cnv.set("name", f"GroupText {n}")
        etree.SubElement(nv, f"{{{P}}}cNvSpPr").set("txBox", "1")
        etree.SubElement(nv, f"{{{P}}}nvPr")
        sppr = etree.SubElement(sp, f"{{{P}}}spPr")
        x = etree.SubElement(sppr, f"{{{A}}}xfrm")
        o = etree.SubElement(x, f"{{{A}}}off")
        o.set("x", str(int(Inches(1 + n))))
        o.set("y", str(int(Inches(4))))
        e = etree.SubElement(x, f"{{{A}}}ext")
        e.set("cx", str(int(Inches(3))))
        e.set("cy", str(int(Inches(0.6))))
        geom = etree.SubElement(sppr, f"{{{A}}}prstGeom")
        geom.set("prst", "rect")
        etree.SubElement(geom, f"{{{A}}}avLst")
        body = etree.SubElement(sp, f"{{{P}}}txBody")
        etree.SubElement(body, f"{{{A}}}bodyPr")
        etree.SubElement(body, f"{{{A}}}lstStyle")
        para = etree.SubElement(body, f"{{{A}}}p")
        run = etree.SubElement(para, f"{{{A}}}r")
        rpr = etree.SubElement(run, f"{{{A}}}rPr")
        rpr.set("lang", "en-US")
        t = etree.SubElement(run, f"{{{A}}}t")
        t.text = text
    return group


def build_chaos_deck(path: str) -> str:
    """5 slides full of identical strings in different containers."""
    prs = Presentation()
    prs.slide_width = Inches(13.333)
    prs.slide_height = Inches(7.5)

    # --- slide 1: three identical text boxes + mixed runs + a table + a group
    s1 = _blank(prs)
    _textbox(s1, "Duplicate", left=1.0, top=1.0)
    _textbox(s1, "Duplicate", left=1.0, top=2.0)
    _textbox(s1, "Duplicate", left=1.0, top=3.0)

    mixed = s1.shapes.add_textbox(Inches(6), Inches(1), Inches(5), Inches(1.2))
    frame = mixed.text_frame
    frame.text = ""
    para = frame.paragraphs[0]
    r1 = para.add_run()
    r1.text = "Plain text"
    r2 = para.add_run()
    r2.text = "Bold text"
    r2.font.bold = True
    r2.font.color.rgb = RGBColor(0xC0, 0x00, 0x00)

    multi = s1.shapes.add_textbox(Inches(6), Inches(3), Inches(5), Inches(2))
    multi.text_frame.text = "First paragraph"
    multi.text_frame.paragraphs[0].add_run().text = ""
    # a real bullet on the paragraph, to prove paragraph properties survive
    p_pr = multi.text_frame.paragraphs[0]._p.get_or_add_pPr()
    p_pr.set("marL", "228600")
    p_pr.set("indent", "-228600")
    etree.SubElement(p_pr, f"{{{A}}}buChar").set("char", "•")

    _textbox(s1, "Bold title", left=1.0, top=6.6, bold=True)

    table = s1.shapes.add_table(2, 2, Inches(1), Inches(5), Inches(6), Inches(1.2)).table
    table.cell(0, 0).text = "Duplicate"
    table.cell(0, 1).text = "Unique cell"
    table.cell(1, 0).text = "Duplicate"
    table.cell(1, 1).text = ""

    _group_with_two_texts(s1, ["Inside group", "Duplicate"])

    # --- slide 2: same title as slide 3, plus notes
    s2 = _blank(prs)
    _textbox(s2, "Section title")
    _textbox(s2, "Second page only", top=2.0)
    s2.notes_slide.notes_text_frame.text = "notes for slide two"

    # --- slide 3: deliberately reuses slide 2's first text
    s3 = _blank(prs)
    _textbox(s3, "Section title")
    _textbox(s3, "Third page only", top=2.0)

    # --- slide 4: picture + text
    s4 = _blank(prs)
    _textbox(s4, "Page with picture")
    image = os.path.splitext(path)[0] + "_dot.png"
    with open(image, "wb") as fh:
        fh.write(PNG_1PX)
    s4.shapes.add_picture(image, Inches(1), Inches(2), Inches(0.5), Inches(0.5))

    # --- slide 5: unicode / emoji / a very long line
    s5 = _blank(prs)
    _textbox(s5, "Mixed script 5G-A · Ünïcödé 🚀")
    _textbox(s5, "A very long line of text " * 20)

    prs.save(path)
    return path


def build_simple_deck(path: str) -> str:
    prs = Presentation()
    prs.slide_width = Inches(13.333)
    prs.slide_height = Inches(7.5)
    slide = _blank(prs)
    _textbox(slide, "Title")
    _textbox(slide, "Body", top=2.0)
    prs.save(path)
    return path
