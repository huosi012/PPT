"""给 PDF 每页铺水印并输出图片版 PDF。
用法：python3 -I wm.py <in.pdf> <watermark.png> <out.pdf> [dpi=200] [opacity=0.55] [cols=2] [rows=3] [scale=0.46]
"""
import sys, os, glob, subprocess, tempfile, img2pdf
from PIL import Image
src, wmf, out = sys.argv[1:4]
dpi = int(sys.argv[4]) if len(sys.argv) > 4 else 200
op = float(sys.argv[5]) if len(sys.argv) > 5 else 0.55
cols = int(sys.argv[6]) if len(sys.argv) > 6 else 2
rows = int(sys.argv[7]) if len(sys.argv) > 7 else 3
scale = float(sys.argv[8]) if len(sys.argv) > 8 else 0.46
tmp = tempfile.mkdtemp(dir=os.path.dirname(os.path.abspath(out)))
subprocess.run(['pdftoppm', '-r', str(dpi), '-png', src, f'{tmp}/p'], check=True)
wm0 = Image.open(wmf).convert('RGBA')
a = wm0.getchannel('A').point(lambda v: 0 if v < 12 else int(v * op))   # 去掉边缘噪点，再整体淡化
wm0.putalpha(a)
jpgs = []
for f in sorted(glob.glob(f'{tmp}/p-*.png'), key=lambda x: int(x.rsplit('-', 1)[1][:-4])):
    pg = Image.open(f).convert('RGBA')
    W, H = pg.size
    w = int(W * scale); wm = wm0.resize((w, int(w * wm0.height / wm0.width)), Image.LANCZOS)
    layer = Image.new('RGBA', pg.size, (0, 0, 0, 0))
    for r in range(rows):
        for c in range(cols):
            cx = W * (c + 0.5) / cols; cy = H * (r + 0.5) / rows
            layer.alpha_composite(wm, (int(cx - wm.width / 2), int(cy - wm.height / 2)))
    pg.alpha_composite(layer)
    j = f[:-4] + '.jpg'
    pg.convert('RGB').save(j, quality=88, dpi=(dpi, dpi), optimize=True)
    jpgs.append(j)
with open(out, 'wb') as fh:
    fh.write(img2pdf.convert(jpgs, layout_fun=img2pdf.get_fixed_dpi_layout_fun((dpi, dpi))))
for f in glob.glob(f'{tmp}/*'): os.remove(f)
os.rmdir(tmp)
print(out, len(jpgs), 'pages', round(os.path.getsize(out) / 1e6, 1), 'MB')
