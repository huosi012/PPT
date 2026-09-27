#!/usr/bin/env python3
"""usage: build.py page1.txt [page2.txt ...] -o out.docx [--preview PREFIX] [--pagebreak]"""
import argparse
import os
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import lib  # noqa: E402

ap = argparse.ArgumentParser()
ap.add_argument("pages", nargs="+")
ap.add_argument("-o", "--out", required=True)
ap.add_argument("--preview")
ap.add_argument("--pagebreak", action="store_true", help="page break between source files")
ap.add_argument("--title", default="线性代数笔记")
a = ap.parse_args()

paras = []
for k, fn in enumerate(a.pages):
    if k and a.pagebreak:
        paras.append('<w:p><w:r><w:br w:type="page"/></w:r></w:p>')
    paras += lib.render_blocks(lib.parse_page(open(fn, encoding="utf-8").read()))
lib.write_docx(paras, a.out, a.title)
print("wrote", a.out, "paragraphs:", len(paras))
if a.preview:
    for p in lib.render_preview(a.out, a.preview):
        print("preview", p)
