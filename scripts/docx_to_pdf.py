#!/usr/bin/env python3
"""Convert generated .docx label sheets to PDF.

This is the final, optional step of the pipeline: turn each finished Word sheet
into a print-ready PDF.

Backends (auto, in priority order):
  1. LibreOffice  -- portable, headless; works on Linux / Windows / macOS
                    install from https://libreoffice.org (or your package manager)
  2. Microsoft Word (Windows COM) -- needs `pip install pywin32` AND Word installed

Usage:
  python scripts/docx_to_pdf.py --src STYX_label_word_docs
  python scripts/docx_to_pdf.py --file one.docx --out pdf_out
  python scripts/docx_to_pdf.py --src STYX_label_word_docs --backend libreoffice
"""
import os
import glob
import shutil
import argparse
import subprocess
import sys


def find_libreoffice():
    name = shutil.which("soffice") or shutil.which("libreoffice")
    if name:
        return name
    for p in (
        r"C:\Program Files\LibreOffice\program\soffice.exe",
        r"C:\Program Files (x86)\LibreOffice\program\soffice.exe",
        "/usr/bin/soffice",
        "/opt/libreoffice/program/soffice",
        "/Applications/LibreOffice.app/Contents/MacOS/soffice",
    ):
        if os.path.exists(p):
            return p
    return None


def convert_libreoffice(soffice, docx, out_dir):
    # LibreOffice writes <basename>.pdf into out_dir
    subprocess.run(
        [soffice, "--headless", "--convert-to", "pdf", "--outdir", out_dir, docx],
        check=True,
        capture_output=True,
        text=True,
    )


def convert_word_com(docx, out_pdf):
    import win32com.client

    word = win32com.client.Dispatch("Word.Application")
    word.Visible = False
    try:
        doc = word.Documents.Open(os.path.abspath(docx))
        try:
            doc.SaveAs(os.path.abspath(out_pdf), FileFormat=17)  # 17 = wdFormatPDF
        finally:
            doc.Close()
    finally:
        word.Quit()


def main():
    ap = argparse.ArgumentParser(description="Convert .docx label sheets to PDF.")
    ap.add_argument("--src", help="folder of .docx files to convert")
    ap.add_argument("--file", help="convert a single .docx instead of a folder")
    ap.add_argument("--out", help="output PDF folder (default: alongside each .docx)")
    ap.add_argument(
        "--backend",
        choices=["auto", "libreoffice", "word"],
        default="auto",
        help="conversion backend (default: auto-detect)",
    )
    args = ap.parse_args()

    if args.file:
        docxs = [args.file]
    elif args.src:
        docxs = sorted(glob.glob(os.path.join(args.src, "*.docx")))
    else:
        ap.error("provide --src (folder) or --file (single .docx)")

    if not docxs:
        print("No .docx found.", file=sys.stderr)
        sys.exit(1)
    print("docx to convert:", len(docxs))

    # --- pick backend ---
    soffice = find_libreoffice() if args.backend in ("auto", "libreoffice") else None
    if args.backend == "libreoffice":
        if not soffice:
            print("LibreOffice not found. Install it or use --backend word.", file=sys.stderr)
            sys.exit(2)
        backend = "lo"
    elif args.backend == "word":
        backend = "word"
    else:  # auto
        backend = "lo" if soffice else "word"

    if backend == "word":
        try:
            import win32com.client  # noqa: F401
        except ImportError:
            print("pywin32 not installed. Run `pip install pywin32` (MS Word also required).", file=sys.stderr)
            sys.exit(3)
    print("backend:", "LibreOffice" if backend == "lo" else "Microsoft Word (COM)")

    ok = 0
    for dx in docxs:
        out_dir = args.out or os.path.dirname(os.path.abspath(dx))
        os.makedirs(out_dir, exist_ok=True)
        out_pdf = os.path.join(out_dir, os.path.splitext(os.path.basename(dx))[0] + ".pdf")
        try:
            if backend == "lo":
                convert_libreoffice(soffice, dx, out_dir)
            else:
                convert_word_com(dx, out_pdf)
        except Exception as e:  # noqa: BLE001
            print("FAIL", dx, "->", e, file=sys.stderr)
            continue
        if not os.path.exists(out_pdf):
            print("FAIL (no pdf produced)", dx, file=sys.stderr)
            continue
        ok += 1
    print("DONE ->", args.out or "(alongside source)", "| pdf:", ok, "/", len(docxs))


if __name__ == "__main__":
    main()
