#!/usr/bin/env python3
"""usage: preview.py p02 [p13 ...]
Builds out/<name>.docx from pages/<name>.txt and writes one stitched preview image per page
to ../preview/第<N>页_预览.png (blank margins and page footers trimmed)."""
import os
import sys

import pymupdf
from PIL import Image

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, HERE)
import lib  # noqa: E402

PREVIEW_DIR = os.path.join(HERE, "..", "preview")
DPI = 140


def content_rows(im, footer_frac=0.9):
    """(top, bottom) of rows that contain ink, ignoring the page-number footer"""
    g = im.convert("L")
    w, h = g.size
    limit = int(h * footer_frac)
    px = g.load()
    rows = [y for y in range(0, limit, 2) if any(px[x, y] < 200 for x in range(0, w, 3))]
    return (rows[0], rows[-1]) if rows else None


def stitch(pdf_path, out_png):
    parts = []
    for pg in pymupdf.open(pdf_path):
        pix = pg.get_pixmap(dpi=DPI)
        im = Image.frombytes("RGB", (pix.width, pix.height), pix.samples)
        rr = content_rows(im)
        if rr:
            top, bot = rr
            parts.append(im.crop((0, max(0, top - 30), im.width, min(im.height, bot + 30))))
    width = max(p.width for p in parts)
    gap = 24
    total = sum(p.height for p in parts) + gap * (len(parts) - 1)
    sheet = Image.new("RGB", (width, total), "white")
    y = 0
    for k, p in enumerate(parts):
        sheet.paste(p, (0, y))
        y += p.height
        if k < len(parts) - 1:
            for x in range(60, width - 60, 12):  # dashed separator where a Word page break falls
                sheet.paste((200, 200, 200), (x, y + gap // 2, x + 6, y + gap // 2 + 2))
            y += gap
    sheet.save(out_png)
    return out_png


if __name__ == "__main__":
    os.makedirs(PREVIEW_DIR, exist_ok=True)
    for name in sys.argv[1:]:
        src = os.path.join(HERE, "pages", name + ".txt")
        docx = os.path.join(HERE, "out", name + ".docx")
        lib.write_docx(lib.render_blocks(lib.parse_page(open(src, encoding="utf-8").read())), docx)
        lib.render_preview(docx, os.path.join(HERE, "out", name))
        num = int(name.lstrip("p"))
        png = stitch(os.path.splitext(docx)[0] + ".pdf", os.path.join(PREVIEW_DIR, "第%d页_预览.png" % num))
        print(png)
