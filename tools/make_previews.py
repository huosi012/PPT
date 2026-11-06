# -*- coding: utf-8 -*-
"""把 render_preview.sh 渲染出的逐页 PNG 压缩为 preview/ 下的 JPG，并拼一张总览图。
用法：python3 tools/make_previews.py <渲染目录>
"""
import glob
import os
import sys

from PIL import Image

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
src = sys.argv[1]
out = os.path.join(ROOT, 'preview')
os.makedirs(out, exist_ok=True)
for f in glob.glob(os.path.join(out, '*.jpg')):
    os.remove(f)
files = sorted(glob.glob(os.path.join(src, 'slide-*.png')))
thumbs = []
for i, f in enumerate(files, 1):
    im = Image.open(f).convert('RGB')
    im = im.resize((1280, round(im.height * 1280 / im.width)), Image.LANCZOS)
    im.save(os.path.join(out, 'slide-%02d.jpg' % i), quality=86, optimize=True)
    thumbs.append(im)
cols, tw = 5, 384
th = round(thumbs[0].height * tw / thumbs[0].width)
rows = (len(thumbs) + cols - 1) // cols
pad = 12
sheet = Image.new('RGB', (cols * tw + (cols + 1) * pad, rows * th + (rows + 1) * pad), (233, 237, 242))
for i, im in enumerate(thumbs):
    sheet.paste(im.resize((tw, th), Image.LANCZOS), (pad + (i % cols) * (tw + pad), pad + (i // cols) * (th + pad)))
sheet.save(os.path.join(out, 'overview.jpg'), quality=88, optimize=True)
print('previews:', len(thumbs))
