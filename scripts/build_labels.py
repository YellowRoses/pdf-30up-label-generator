#!/usr/bin/env python3
"""Inject each rendered PNG into ALL cells of a Word label template.

For every source PDF we:
  * copy the TEMPLATE .docx,
  * embed the matching PNG once into word/media/label.png,
  * add one relationship,
  * insert an inline <w:drawing> (stretched to fill) into every <w:tc> cell.

Cell geometry (tcW / trH / borders / position) is NEVER modified -- we only
drop a drawing inside each existing cell. Works whether the template is a single
30-cell table or several tables (incl. floating text-box tables).

Usage:
  python scripts/build_labels.py \
      --template amazon_fnsku_label_template_30up.docx \
      --src "STYX (new SKU code) Label" --png-dir pdfpng \
      --out STYX_label_word_docs --imgw-in 2.602 --imgh-in 0.992
"""
import os
import glob
import zipfile
import argparse
import re
from lxml import etree

W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
NS = {
    "w": W,
    "wp": "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing",
    "a": "http://schemas.openxmlformats.org/drawingml/2006/main",
    "pic": "http://schemas.openxmlformats.org/drawingml/2006/picture",
    "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
}


def drawing_xml(docpr_id: int, embed_rid: str, cx: int, cy: int) -> str:
    # No cNvGraphicFramePr/noChangeAspect -> image truly stretches to extent.
    return f'''<w:drawing xmlns:w="{NS['w']}" xmlns:wp="{NS['wp']}" xmlns:a="{NS['a']}" xmlns:pic="{NS['pic']}" xmlns:r="{NS['r']}">
  <wp:inline distT="0" distB="0" distL="0" distR="0">
    <wp:extent cx="{cx}" cy="{cy}"/>
    <wp:effectExtent l="0" t="0" r="0" b="0"/>
    <wp:docPr id="{docpr_id}" name="Label {docpr_id}"/>
    <a:graphic xmlns:a="{NS['a']}">
      <a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">
        <pic:pic xmlns:pic="{NS['pic']}">
          <pic:nvPicPr><pic:cNvPr id="{docpr_id}" name="label{docpr_id}"/><pic:cNvPicPr/></pic:nvPicPr>
          <pic:blipFill><a:blip r:embed="{embed_rid}"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill>
          <pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="{cx}" cy="{cy}"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr>
        </pic:pic>
      </a:graphicData>
    </a:graphic>
  </wp:inline>
</w:drawing>'''


def main():
    ap = argparse.ArgumentParser(description="Inject PNGs into a Word label template's cells.")
    ap.add_argument("--template", required=True, help="path to the .docx label template")
    ap.add_argument("--src", required=True, help="folder of source PDFs (used for naming/count)")
    ap.add_argument("--png-dir", required=True, help="folder of rendered PNGs (base-named)")
    ap.add_argument("--out", required=True, help="output .docx folder")
    ap.add_argument("--imgw-in", type=float, default=2.602, help="image width in inches")
    ap.add_argument("--imgh-in", type=float, default=0.992, help="image height in inches")
    args = ap.parse_args()

    cx = int(round(args.imgw_in * 914400))
    cy = int(round(args.imgh_in * 914400))
    os.makedirs(args.out, exist_ok=True)

    pdfs = sorted(glob.glob(os.path.join(args.src, "*.pdf")))
    print("PDFs:", len(pdfs))
    ok = 0
    for pdf in pdfs:
        base = os.path.splitext(os.path.basename(pdf))[0]
        png = os.path.join(args.png_dir, base + ".png")
        if not os.path.exists(png):
            print("MISSING png:", png)
            continue
        out = os.path.join(args.out, base + ".docx")
        png_bytes = open(png, "rb").read()

        zin = zipfile.ZipFile(args.template)

        # ---- new relationship id for the embedded image ----
        rels_bytes = zin.read("word/_rels/document.xml.rels")
        rids = re.findall(r'Id="(rId\d+)"', rels_bytes.decode("utf-8"))
        maxr = max(int(r[3:]) for r in rids)
        newrid = f"rId{maxr + 1}"
        rels_tree = etree.fromstring(rels_bytes)
        RNS = "http://schemas.openxmlformats.org/package/2006/relationships"
        newrel = etree.SubElement(rels_tree, "{%s}Relationship" % RNS)
        newrel.set("Id", newrid)
        newrel.set("Type", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image")
        newrel.set("Target", "media/label.png")
        new_rels = etree.tostring(rels_tree, xml_declaration=True, encoding="UTF-8", standalone=True)

        # ---- inject a drawing into every existing cell ----
        doc_tree = etree.fromstring(zin.read("word/document.xml"))
        tcs = doc_tree.findall(".//{%s}tc" % W)
        for idx, tc in enumerate(tcs, start=1):
            p = tc.find("{%s}p" % W)
            if p is None:
                p = etree.SubElement(tc, "{%s}p" % W)
            pPr = p.find("{%s}pPr" % W)
            if pPr is None:
                pPr = etree.SubElement(p, "{%s}pPr" % W)
            jc = pPr.find("{%s}jc" % W)
            if jc is None:
                jc = etree.SubElement(pPr, "{%s}jc" % W)
            jc.set("{%s}val" % W, "center")
            p.append(etree.fromstring(drawing_xml(idx, newrid, cx, cy)))
        new_doc = etree.tostring(doc_tree, xml_declaration=True, encoding="UTF-8", standalone=True)

        # ---- ensure png is registered in content types ----
        ct_bytes = zin.read("[Content_Types].xml")
        if b'Extension="png"' not in ct_bytes:
            ct_tree = etree.fromstring(ct_bytes)
            CTS = "http://schemas.openxmlformats.org/package/2006/content-types"
            d = etree.SubElement(ct_tree, "{%s}Default" % CTS)
            d.set("Extension", "png")
            d.set("ContentType", "image/png")
            new_ct = etree.tostring(ct_tree, xml_declaration=True, encoding="UTF-8", standalone=True)
        else:
            new_ct = ct_bytes

        # ---- write output zip (template + injected parts + media) ----
        zout = zipfile.ZipFile(out, "w", zipfile.ZIP_DEFLATED)
        for item in zin.infolist():
            data = zin.read(item.filename)
            if item.filename == "word/document.xml":
                data = new_doc
            elif item.filename == "word/_rels/document.xml.rels":
                data = new_rels
            elif item.filename == "[Content_Types].xml":
                data = new_ct
            zout.writestr(item, data)
        zout.writestr("word/media/label.png", png_bytes)
        zout.close()
        zin.close()
        ok += 1

    print("DONE ->", args.out, "| files:", ok)


if __name__ == "__main__":
    main()
