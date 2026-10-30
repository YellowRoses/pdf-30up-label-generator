---
name: pdf-30up-label-generator
description: Batch-convert a folder of PDFs into printable 30-up (Avery 5160 style) Word label sheets. Renders each PDF to a PNG and injects it into every cell of a label template, preserving the template's exact cell geometry. Use for "PDF into 30-up Word", "label sheet generation", "FNSKU / product label to Word", or any task where each sheet must show the same image repeated across a grid of label cells.
---

# PDF → 30-up Word Label Generator

## When to use
- User has a folder of PDFs (product labels, FNSKU, barcodes, price tags, etc.) and wants each turned into a Word document where one sheet contains a grid of identical labels (commonly 30-up = 3 columns × 10 rows on Letter/A4).
- The key requirement is to **preserve the template's cell layout exactly** while dropping the image into every cell.

## Step 0 — Ask the user first (ALWAYS, before running anything)

Never assume these values — collect them from the user, even mid-conversation:

1. **Source PDF folder** — where the label PDFs live.
2. **Word template** — the user's own `.docx`, or the bundled `examples/amazon_fnsku_label_template_30up.docx` (30-up, Avery 5160 geometry, ready to use). **The layout MUST follow this template** — never rebuild the grid yourself (see Critical lessons).
3. **Label display size** — width × height as printed inside each cell. The user may give **cm or inches**; convert to inches for the scripts (`cm ÷ 2.54`). Verified reference: Amazon FNSKU labels display at **6.61 cm × 2.52 cm** (= `2.602″ × 0.992″`) inside a 2.625″ × 1.0″ cell. Pass via `--imgw-in` / `--img-h-in`.
4. **Output folder** (optional) — sensible default is fine.

If the user doesn't know their label size, measure a printed label or check the template's cell interior, then **confirm the numbers before running**.

## Workflow
1. **Render** each PDF's first page to a PNG → `scripts/render_pdfs.py`.
   Backends (auto, in order): `--pdf-cli` exe (preferred, e.g. `workbuddy-pdf.exe`) → PyMuPDF (`fitz`) → `pdf2image` (poppler).
2. **Inject** each PNG into ALL cells of the **user's template from Step 0** → `scripts/build_labels.py`.
   The script copies the template per PDF, embeds the PNG once, and inserts an inline `<w:drawing>` into every `<w:tc>` it finds — so it works whether the template is a single 30-cell table **or several tables (incl. floating text-box tables)**.
3. Output: one `.docx` per source PDF in the output folder, named after the PDF.
4. **(Optional) Convert to PDF** — two paths, pick by environment:
   - **No Word/LibreOffice available (or PDFs are all you need)** → `scripts/png_to_pdf.py`: rebuilds the 30-up PDF straight from the step-1 PNGs, zero dependencies (pure stdlib), byte-compatible geometry with the Word sheet. Accepts flat `<base>.png`, numbered `<NN>/page-0001.png`, or named `<base>/page-0001.png` PNG layouts. Supports `--limit N` for smoke tests and `--cols/--rows/--img-w-in/--img-h-in` overrides.
   - **Word sheets needed too** → `scripts/docx_to_pdf.py`. Backends (auto): **LibreOffice** (headless, portable) → **Microsoft Word COM** (Windows, needs `pywin32` + Word). Ensure one is installed before running.

## Critical lessons (learned the hard way)
- ❌ **Do NOT build a new table** (e.g. `python-docx add_table(10,3)`) to fake the 30-up layout. It will not match the template's exact geometry and the cells will *look enlarged*. Always **inject into the template's existing cells**.
- ✅ The template's cells keep their own width/row-height (`tcW`, `trH`). We only append a drawing; cell size is never modified.
- ✅ To "fill" the cell, set the image's display size to the cell interior (e.g. `2.602″ × 0.992″` for a `2.625″ × 1.0″` cell). Do **NOT** set `noChangeAspect` on `cNvGraphicFramePr`, or the image won't stretch.
- Embed each image **once** in `word/media/` and reference it via a single relationship; deep-copy the `<w:drawing>` into every cell.

## Usage
Run only after Step 0 collected the template and label size from the user.

```bash
python scripts/render_pdfs.py --src "<user's PDF folder>" --out pdfpng --dpi 600 \
    --pdf-cli "C:/Users/Administrator/.workbuddy/external-packages/pdf-cli/workbuddy-pdf.exe"

# --template: user's own template, or the bundled examples/amazon_fnsku_label_template_30up.docx
# --imgw-in/--imgh-in: the size collected in Step 0 (cm ÷ 2.54); reference: 2.602 x 0.992
python scripts/build_labels.py \
    --template examples/amazon_fnsku_label_template_30up.docx \
    --src "<user's PDF folder>" --png-dir pdfpng --out STYX_label_word_docs \
    --imgw-in 2.602 --imgh-in 0.992
```
No template of your own? The bundled `examples/amazon_fnsku_label_template_30up.docx` is a real 30-up template — use it directly. A synthetic demo can also be generated:
```bash
python examples/make_sample_template.py demo_template.docx
```

Convert the finished Word sheets to PDF:
```bash
python scripts/docx_to_pdf.py --src STYX_label_word_docs
# or a single file
python scripts/docx_to_pdf.py --file STYX_label_word_docs/AP15500\ v2.docx --out pdf_out
```

Or skip Word entirely and build the PDFs straight from the PNGs (zero dependencies) — pass the same Step 0 size:
```bash
python scripts/png_to_pdf.py --src "<user's PDF folder>" --png-dir pdfpng --out pdf_out \
    --img-w-in 2.602 --img-h-in 0.992
python scripts/png_to_pdf.py --src labels --png-dir pdfpng --out pdf_out --limit 2  # smoke test
```

## Validate
Each output `.docx` should have: `document.xml` well-formed, N drawings = N cells, all `r:embed` pointing to one `media/label.png`, image extent = the configured size, and cell `tcW`/`trH` byte-identical to the template.
