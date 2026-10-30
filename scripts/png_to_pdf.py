#!/usr/bin/env python3
"""Build print-ready 30-up label PDFs directly from PNG renders.

Zero-dependency path: no Word / LibreOffice / PyMuPDF / pip packages needed.
Each source PDF becomes one Letter PDF with COLS x ROWS copies of its label
image laid out in a grid with dashed cell borders, matching the Word sheet's
geometry (Avery 5160 style).

Why skip the .docx? The Word sheets are just 30 identical images per page, so
we rebuild the equivalent PDF straight from the source render. Use this path
when no Word/LibreOffice is available, or when you only need PDFs anyway.

PNG lookup (per source PDF, in order):
  1. <png-dir>/<base>.png                  (flat, as written by render_pdfs.py)
  2. <png-dir>/<index:02d>/page-0001.png   (numbered dirs, pdf-cli batch output)
  3. <png-dir>/<base>/page-0001.png        (named dirs)

The label image is embedded ONCE per PDF (as a reused XObject) and referenced
COLS*ROWS times, keeping file sizes small.

Usage:
  python scripts/png_to_pdf.py --src "STYX (new SKU code) Label" --png-dir pdfpng --out pdf_out
  python scripts/png_to_pdf.py --src labels --png-dir pdfpng --out pdf_out --limit 2  # smoke test
"""
import os
import glob
import struct
import zlib
import argparse

# --- default template geometry (points, 72pt = 1 inch) ---
# Letter page; margins/cells match the verified 30-up Word template.
PAGE_W, PAGE_H = 612.0, 792.0            # Letter
MARGIN_L, MARGIN_T = 13.248, 14.616      # 0.184" / 0.203"
CELL_W, CELL_H = 189.0, 71.5             # 2.625" / ~0.993"
IMG_W, IMG_H = 187.344, 71.424           # 2.602" / 0.992" (matches Word sheet)
TARGET_W = 780                           # downscale label to ~300 DPI before embed


def decode_png(path):
    d = open(path, 'rb').read()
    assert d[:8] == b'\x89PNG\r\n\x1a\n', 'not a PNG'
    pos, idat = 8, b''
    W = H = ct = interlace = None
    while pos < len(d):
        ln = struct.unpack('>I', d[pos:pos + 4])[0]
        typ = d[pos + 4:pos + 8]
        data = d[pos + 8:pos + 8 + ln]
        if typ == b'IHDR':
            W, H, _bd, ct = struct.unpack('>IIBB', data[:10])
            interlace = data[12]
        elif typ == b'IDAT':
            idat += data
        elif typ == b'IEND':
            break
        pos += 12 + ln
    assert interlace == 0, 'interlaced PNG unsupported'
    assert ct in (2, 6), f'unsupported colortype {ct}'
    raw = zlib.decompress(idat)
    bpp = 4 if ct == 6 else 3
    stride = W * bpp
    out = bytearray(H * stride)
    prev = bytearray(stride)
    p = 0

    def paeth(a, b, c):
        pp = a + b - c
        pa, pb, pc = abs(pp - a), abs(pp - b), abs(pp - c)
        if pa <= pb and pa <= pc:
            return a
        if pb <= pc:
            return b
        return c

    for y in range(H):
        f = raw[p]; p += 1
        line = bytearray(raw[p:p + stride]); p += stride
        if f == 1:
            for i in range(bpp, stride):
                line[i] = (line[i] + line[i - bpp]) & 255
        elif f == 2:
            for i in range(stride):
                line[i] = (line[i] + prev[i]) & 255
        elif f == 3:
            for i in range(stride):
                a = line[i - bpp] if i >= bpp else 0
                line[i] = (line[i] + ((a + prev[i]) >> 1)) & 255
        elif f == 4:
            for i in range(stride):
                a = line[i - bpp] if i >= bpp else 0
                c = prev[i - bpp] if i >= bpp else 0
                line[i] = (line[i] + paeth(a, prev[i], c)) & 255
        out[y * stride:(y + 1) * stride] = line
        prev = line

    if ct == 6:
        rgb = bytearray(W * H * 3)
        for i in range(W * H):
            r, g, b, a = out[i * 4:i * 4 + 4]
            rgb[i * 3]     = (r * a + (255 - a) * 255) // 255
            rgb[i * 3 + 1] = (g * a + (255 - a) * 255) // 255
            rgb[i * 3 + 2] = (b * a + (255 - a) * 255) // 255
        return W, H, bytes(rgb)
    return W, H, bytes(out)


def downscale(rgb, W, H, tw):
    if tw >= W:
        return rgb, W, H
    th = max(1, round(H * tw / W))
    out = bytearray(tw * th * 3)
    fx, fy = W / tw, H / th
    for y in range(th):
        y0, y1 = int(y * fy), max(int((y + 1) * fy), int(y * fy) + 1)
        for x in range(tw):
            x0, x1 = int(x * fx), max(int((x + 1) * fx), int(x * fx) + 1)
            r = g = b = n = 0
            for yy in range(y0, y1):
                base = yy * W * 3
                for xx in range(x0, x1):
                    i = base + xx * 3
                    r += rgb[i]; g += rgb[i + 1]; b += rgb[i + 2]; n += 1
            o = (y * tw + x) * 3
            out[o], out[o + 1], out[o + 2] = r // n, g // n, b // n
    return bytes(out), tw, th


def build_pdf(rgb, W, H, out_path, cols, rows, img_w, img_h):
    # grid (top-based coordinates)
    grid = []
    for row in range(rows):
        for col in range(cols):
            xL = MARGIN_L + col * CELL_W
            yT = MARGIN_T + row * CELL_H
            grid.append((xL, yT, CELL_W, CELL_H))
    if MARGIN_L + cols * CELL_W > PAGE_W or MARGIN_T + rows * CELL_H > PAGE_H:
        print(f"WARN: {cols}x{rows} grid overflows the page; check geometry")

    img = zlib.compress(rgb, 9)
    ops = []
    for (xL, yT, cw, ch) in grid:
        ix = xL + (cw - img_w) / 2
        iy = yT + (ch - img_h) / 2
        yb = PAGE_H - (iy + img_h)                     # bottom-left y for cm
        ops.append(f"q {img_w:.3f} 0 0 {img_h:.3f} {ix:.3f} {yb:.3f} cm /Im0 Do Q")
        bx, by = xL, PAGE_H - (yT + ch)                # dashed border rect
        ops.append(f"0.6 w [3 2] 0 d {bx:.3f} {by:.3f} {cw:.3f} {ch:.3f} re S")
    content = ("\n".join(ops)).encode('latin-1')

    objs = {}
    objs[1] = b"<< /Type /Catalog /Pages 2 0 R >>"
    objs[2] = b"<< /Type /Pages /Kids [3 0 R] /Count 1 >>"
    objs[3] = (b"<< /Type /Page /Parent 2 0 R /MediaBox [0 0 612 792] "
               b"/Resources << /XObject << /Im0 5 0 R >> >> /Contents 4 0 R >>")
    objs[4] = (b"<< /Length " + str(len(content)).encode() + b" >>\nstream\n" +
               content + b"\nendstream")
    objs[5] = (b"<< /Type /XObject /Subtype /Image /Width " + str(W).encode() +
               b" /Height " + str(H).encode() +
               b" /ColorSpace /DeviceRGB /BitsPerComponent 8 /Filter /FlateDecode"
               b" /Length " + str(len(img)).encode() + b" >>\nstream\n" +
               img + b"\nendstream")

    out = bytearray(b"%PDF-1.4\n")
    offsets = {}
    for i in range(1, 6):
        offsets[i] = len(out)
        out += (f"{i} 0 obj\n").encode() + objs[i] + b"\nendobj\n"
    xref_pos = len(out)
    out += b"xref\n0 6\n0000000000 65535 f \n"
    for i in range(1, 6):
        out += ("%010d 00000 n \n" % offsets[i]).encode()
    out += (f"trailer\n<< /Size 6 /Root 1 0 R >>\nstartxref\n{xref_pos}\n%%EOF").encode()
    open(out_path, 'wb').write(out)


def find_png(png_dir, base, idx):
    for cand in (
        os.path.join(png_dir, base + ".png"),                   # flat (render_pdfs.py)
        os.path.join(png_dir, f"{idx:02d}", "page-0001.png"),   # numbered dirs (pdf-cli)
        os.path.join(png_dir, base, "page-0001.png"),           # named dirs
    ):
        if os.path.exists(cand):
            return cand
    return None


def main():
    ap = argparse.ArgumentParser(
        description="Build 30-up label PDFs directly from PNG renders (zero dependencies).")
    ap.add_argument("--src", required=True, help="folder of source PDFs (names drive output naming)")
    ap.add_argument("--png-dir", required=True, help="folder with rendered label PNGs")
    ap.add_argument("--out", required=True, help="output PDF folder")
    ap.add_argument("--cols", type=int, default=3, help="grid columns (default: 3)")
    ap.add_argument("--rows", type=int, default=10, help="grid rows (default: 10)")
    ap.add_argument("--img-w-in", type=float, default=2.602,
                    help="label display width in inches (default: 2.602)")
    ap.add_argument("--img-h-in", type=float, default=0.992,
                    help="label display height in inches (default: 0.992)")
    ap.add_argument("--target-width", type=int, default=TARGET_W,
                    help=f"downscale label to this pixel width before embed (default: {TARGET_W})")
    ap.add_argument("--limit", type=int, default=None,
                    help="only process the first N source PDFs (smoke tests)")
    args = ap.parse_args()

    os.makedirs(args.out, exist_ok=True)
    pdfs = sorted(glob.glob(os.path.join(args.src, "*.pdf")))
    if args.limit:
        pdfs = pdfs[:args.limit]
    print("source PDFs:", len(pdfs))

    img_w, img_h = args.img_w_in * 72, args.img_h_in * 72
    ok = 0
    for i, pdf in enumerate(pdfs):
        base = os.path.splitext(os.path.basename(pdf))[0]
        png = find_png(args.png_dir, base, i)
        if not png:
            print("MISSING PNG for", base); continue
        W, H, rgb = decode_png(png)
        rgb, W, H = downscale(rgb, W, H, args.target_width)
        out = os.path.join(args.out, base + ".pdf")
        build_pdf(rgb, W, H, out, args.cols, args.rows, img_w, img_h)
        ok += 1
    print("DONE ->", args.out, "| pdf written:", ok)


if __name__ == "__main__":
    main()
