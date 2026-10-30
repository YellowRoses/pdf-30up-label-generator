---
name: pdf-30up-label-generator
description: Batch-convert a folder of PDFs into printable 30-up (Avery 5160 style) Word label sheets. Renders each PDF to a PNG and injects it into every cell of a label template, preserving the template's exact cell geometry. Use for "PDF into 30-up Word", "label sheet generation", "FNSKU / product label to Word", or any task where each sheet must show the same image repeated across a grid of label cells.
---

# PDF → 30-up Word Label Generator

## When to use
- User has a folder of PDFs (product labels, FNSKU, barcodes, price tags, etc.) and wants each turned into a Word document where one sheet contains a grid of identical labels (commonly 30-up = 3 columns × 10 rows on Letter/A4).
- The key requirement is to **preserve the template's cell layout exactly** while dropping the image into every cell.

## Workflow
1. **Render** each PDF's first page to a PNG → `scripts/render_pdfs.py`.
   Backends (auto, in order): `--pdf-cli` exe (preferred, e.g. `workbuddy-pdf.exe`) → PyMuPDF (`fitz`) → `pdf2image` (poppler).
2. **Inject** each PNG into ALL cells of the template → `scripts/build_labels.py`.
   The script copies the template per PDF, embeds the PNG once, and inserts an inline `<w:drawing>` into every `<w:tc>` it finds — so it works whether the template is a single 30-cell table **or several tables (incl. floating text-box tables)**.
3. Output: one `.docx` per source PDF in the output folder, named after the PDF.
4. **(Optional) Convert to PDF** (`scripts/docx_to_pdf.py`) — turn the finished Word sheets into print-ready PDFs.
   Backends (auto): **LibreOffice** (headless, portable) → **Microsoft Word COM** (Windows, needs `pywin32` + Word). Ensure one is installed before running.

## Critical lessons (learned the hard way)
- ❌ **Do NOT build a new table** (e.g. `python-docx add_table(10,3)`) to fake the 30-up layout. It will not match the template's exact geometry and the cells will *look enlarged*. Always **inject into the template's existing cells**.
- ✅ The template's cells keep their own width/row-height (`tcW`, `trH`). We only append a drawing; cell size is never modified.
- ✅ To "fill" the cell, set the image's display size to the cell interior (e.g. `2.602″ × 0.992″` for a `2.625″ × 1.0″` cell). Do **NOT** set `noChangeAspect` on `cNvGraphicFramePr`, or the image won't stretch.
- Embed each image **once** in `word/media/` and reference it via a single relationship; deep-copy the `<w:drawing>` into every cell.

## Usage
```bash
python scripts/render_pdfs.py --src "STYX (new SKU code) Label" --out pdfpng --dpi 600 \
    --pdf-cli "C:/Users/Administrator/.workbuddy/external-packages/pdf-cli/workbuddy-pdf.exe"

python scripts/build_labels.py --template amazon_fnsku_label_template_30up.docx \
    --src "STYX (new SKU code) Label" --png-dir pdfpng --out STYX_label_word_docs \
    --imgw-in 2.602 --imgh-in 0.992
```
No template handy? Generate a demo one:
```bash
python examples/make_sample_template.py demo_template.docx
```

Convert the finished Word sheets to PDF:
```bash
python scripts/docx_to_pdf.py --src STYX_label_word_docs
# or a single file
python scripts/docx_to_pdf.py --file STYX_label_word_docs/AP15500\ v2.docx --out pdf_out
```

## Validate
Each output `.docx` should have: `document.xml` well-formed, N drawings = N cells, all `r:embed` pointing to one `media/label.png`, image extent = the configured size, and cell `tcW`/`trH` byte-identical to the template.
