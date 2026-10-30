#!/usr/bin/env python3
"""Render each PDF's first page to a PNG, named by the PDF base name.

Backends (auto, in priority order):
  1. --pdf-cli <exe>   (e.g. workbuddy-pdf.exe)  -- preferred, pixel-perfect
  2. PyMuPDF (fitz)    -- pure Python, cross-platform
  3. pdf2image         -- needs poppler installed

Output: <OUT>/<base>.png  for every <base>.pdf in <SRC>.
"""
import os
import glob
import argparse
import subprocess
import sys


def render_with_pdfcli(pdf, out_png, dpi, pdfcli):
    """workbuddy-pdf.exe writes <dstdir>/page-0001.png; rename to out_png."""
    tmp = out_png + ".tmpdir"
    os.makedirs(tmp, exist_ok=True)
    r = subprocess.run(
        [pdfcli, pdf, tmp, "--format", "png", "--dpi", str(dpi), "--quiet"],
        capture_output=True,
        text=True,
    )
    if r.returncode != 0:
        raise RuntimeError(r.stderr or "pdf-cli exited %d" % r.returncode)
    src = os.path.join(tmp, "page-0001.png")
    if not os.path.exists(src):
        raise RuntimeError("pdf-cli produced no page-0001.png")
    os.replace(src, out_png)
    for f in os.listdir(tmp):
        try:
            os.remove(os.path.join(tmp, f))
        except OSError:
            pass
    try:
        os.rmdir(tmp)
    except OSError:
        pass


def render_with_fitz(pdf, out_png, dpi):
    import fitz  # PyMuPDF

    doc = fitz.open(pdf)
    page = doc.load_page(0)
    pix = page.get_pixmap(dpi=dpi)
    pix.save(out_png)


def render_with_pdf2image(pdf, out_png, dpi):
    from pdf2image import convert_from_path

    imgs = convert_from_path(pdf, dpi=dpi, first_page=1, last_page=1)
    imgs[0].save(out_png, "PNG")


def main():
    ap = argparse.ArgumentParser(description="Render PDFs (first page) to PNGs.")
    ap.add_argument("--src", required=True, help="folder of source PDFs")
    ap.add_argument("--out", required=True, help="output PNG folder")
    ap.add_argument("--dpi", type=int, default=600)
    ap.add_argument(
        "--pdf-cli",
        default=None,
        help="path to workbuddy-pdf.exe or any pdf->png CLI; if omitted, use fitz/pdf2image",
    )
    args = ap.parse_args()

    os.makedirs(args.out, exist_ok=True)
    pdfs = sorted(glob.glob(os.path.join(args.src, "*.pdf")))
    print("PDFs to render:", len(pdfs))

    for pdf in pdfs:
        base = os.path.splitext(os.path.basename(pdf))[0]
        out_png = os.path.join(args.out, base + ".png")
        try:
            if args.pdf_cli and os.path.exists(args.pdf_cli):
                render_with_pdfcli(pdf, out_png, args.dpi, args.pdf_cli)
            else:
                try:
                    render_with_fitz(pdf, out_png, args.dpi)
                except ImportError:
                    render_with_pdf2image(pdf, out_png, args.dpi)
        except Exception as e:  # noqa: BLE001
            print("FAIL", pdf, "->", e, file=sys.stderr)

    print("DONE ->", args.out)


if __name__ == "__main__":
    main()
