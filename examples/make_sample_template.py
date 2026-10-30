#!/usr/bin/env python3
"""Generate a demo 30-up (3x10) Word label template for testing.

The result is a single 10x3 table with dashed borders, 0 cell margins and
fixed layout -- a valid stand-in for a real Avery 5160 sheet, so you can try
the pipeline without your own template.
"""
import sys
from docx import Document
from docx.shared import Inches, Twips
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_ALIGN_VERTICAL, WD_TABLE_ALIGNMENT
from docx.oxml.ns import qn
from docx.oxml import OxmlElement


def set_cell_margins(cell, val=0):
    tcPr = cell._tc.get_or_add_tcPr()
    m = tcPr.find(qn("w:tcMar"))
    if m is None:
        m = OxmlElement("w:tcMar")
        tcPr.append(m)
    for edge in ("top", "left", "bottom", "right"):
        x = m.find(qn(f"w:{edge}"))
        if x is None:
            x = OxmlElement(f"w:{edge}")
            m.append(x)
        x.set(qn("w:w"), str(val))
        x.set(qn("w:type"), "dxa")


def set_borders(table):
    tblPr = table._tbl.tblPr
    b = OxmlElement("w:tblBorders")
    for edge in ("top", "left", "bottom", "right", "insideH", "insideV"):
        e = OxmlElement(f"w:{edge}")
        e.set(qn("w:val"), "dashed")
        e.set(qn("w:color"), "000000")
        e.set(qn("w:sz"), "4")
        e.set(qn("w:space"), "0")
        b.append(e)
    tblPr.append(b)
    lay = OxmlElement("w:tblLayout")
    lay.set(qn("w:type"), "fixed")
    tblPr.append(lay)
    cm = OxmlElement("w:tblCellMar")
    for edge in ("top", "left", "bottom", "right"):
        e = OxmlElement(f"w:{edge}")
        e.set(qn("w:w"), "0")
        e.set(qn("w:type"), "dxa")
        cm.append(e)
    tblPr.append(cm)


def main():
    out = sys.argv[1] if len(sys.argv) > 1 else "demo_template.docx"
    doc = Document()
    s = doc.sections[0]
    s.page_width = Inches(8.5)
    s.page_height = Inches(11)
    s.top_margin = Twips(293)
    s.bottom_margin = Twips(0)
    s.left_margin = Twips(265)
    s.right_margin = Twips(265)

    t = doc.add_table(rows=10, cols=3)
    t.alignment = WD_TABLE_ALIGNMENT.LEFT
    set_borders(t)
    for r in range(10):
        for c in range(3):
            cell = t.cell(r, c)
            cell.width = Inches(2.625)
            set_cell_margins(cell, 0)
            cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
            p = cell.paragraphs[0]
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    doc.save(out)
    print("wrote", out)


if __name__ == "__main__":
    main()
