"""美化后自检：渲染 PDF、出总览图、核对原文是否缺失、核对表格结构与图片。

用法：python3 check.py 原文件.docx 新文件.docx [输出目录]
- 文字核对忽略标点/空白/编号/分隔符（美化会拆条目、改标签），只报告真正找不到的段落
- 结构核对：主表网格列数、行数、每行跨列/纵向合并/嵌套表是否一致（预期有变化时自行判断）
依赖：soffice（需 libreoffice-writer）、pdftoppm（poppler-utils）、Pillow
"""
import sys, os, re, zipfile, subprocess, glob, shutil
from lxml import etree

W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
A = 'http://schemas.openxmlformats.org/drawingml/2006/main'
ns = {'w': W}; q = lambda n: f'{{{W}}}{n}'


def load(path):
    z = zipfile.ZipFile(path)
    return etree.fromstring(z.read('word/document.xml'))


def main_rows(doc):
    body = doc.find('w:body', ns)
    t = max(body.findall('w:tbl', ns), key=lambda t: len(t.findall('w:tr', ns)))
    return t, t.findall('w:tr', ns)


def sig(tr):
    out = []
    for tc in tr.findall('w:tc', ns):
        pr = tc.find('w:tcPr', ns)
        gs = pr.find('w:gridSpan', ns) if pr is not None else None
        vm = pr.find('w:vMerge', ns) if pr is not None else None
        out.append((int(gs.get(q('val'))) if gs is not None else 1,
                    None if vm is None else (vm.get(q('val')) or 'continue'),
                    len(tc.findall('.//w:tbl', ns))))
    return out


STRIP = re.compile(r'[\s、；;。，,：:【】“”"（）()|｜·　▍■\d\.．①②③④⑤]')


def main():
    src, new = sys.argv[1], sys.argv[2]
    out = sys.argv[3] if len(sys.argv) > 3 else os.path.splitext(new)[0] + '_check'
    os.makedirs(out, exist_ok=True)
    a, b = load(src), load(new)
    ta, ra = main_rows(a); tb, rb = main_rows(b)
    ga = [g.get(q('w')) for g in ta.find('w:tblGrid', ns)]
    gb = [g.get(q('w')) for g in tb.find('w:tblGrid', ns)]
    print(f'网格列：{len(ga)} → {len(gb)}（{"一致" if ga == gb else "不同"}）；行数：{len(ra)} → {len(rb)}')
    if len(ra) == len(rb):
        diff = [i for i, (x, y) in enumerate(zip(ra, rb)) if sig(x) != sig(y)]
        print('结构不同的行：', diff or '无')
    # 文字完整性
    alltext = STRIP.sub('', ''.join(x.text or '' for x in b.iter(q('t'))))
    miss = []
    for p in ta.iter(q('p')):
        t = STRIP.sub('', ''.join(x.text or '' for x in p.iter(q('t'))))
        if len(t) >= 2 and t not in alltext: miss.append(t[:60])
    print(f'原文段落在新文件中找不到：{len(miss)}')
    for m in miss: print('  -', m)
    # 图片数
    na = len(list(a.iter(f'{{{A}}}blip'))); nb = len(list(b.iter(f'{{{A}}}blip')))
    print(f'图片（含图标）：{na} → {nb}；浮动图：{len(list(b.iter("{http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing}anchor")))}')
    # 渲染
    shutil.copy(new, f'{out}/new.docx')
    subprocess.run(['soffice', '--headless', '--convert-to', 'pdf', '--outdir', out, f'{out}/new.docx'],
                   capture_output=True)
    pdf = f'{out}/new.pdf'
    if not os.path.exists(pdf):
        print('渲染失败：确认已安装 libreoffice-writer'); return
    for f in glob.glob(f'{out}/pg-*.jpg'): os.remove(f)
    subprocess.run(['pdftoppm', '-jpeg', '-r', '50', pdf, f'{out}/pg'])
    from PIL import Image
    fs = sorted(glob.glob(f'{out}/pg-*.jpg')); ims = [Image.open(f) for f in fs]
    w, h = ims[0].size; cols = 5; rows = (len(ims) + cols - 1) // cols
    sheet = Image.new('RGB', (w * cols, h * rows), 'white')
    for k, im in enumerate(ims): sheet.paste(im, ((k % cols) * w, (k // cols) * h))
    sheet.save(f'{out}/overview.jpg')
    print(f'页数：{len(ims)}；总览图：{out}/overview.jpg')


if __name__ == '__main__':
    main()
