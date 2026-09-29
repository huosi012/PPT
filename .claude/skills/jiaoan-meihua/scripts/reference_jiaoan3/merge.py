# 合并图示：保留原文与原图，加入新图示；与原图重复的用新图替换；原浮动图改为嵌入式
# 由 v2.py 在 page3.py 之后 exec（沿用 rows、rels、pic_id、para、run 等）
import struct

NEWIMG = '/tmp/lo2/n/word/media'
SVGIMG = '/tmp/lo2/svg'
WPNS = 'http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing'
ANS = 'http://schemas.openxmlformats.org/drawingml/2006/main'

rid_target = dict(re.findall(r'Id="(rId\d+)"[^>]*Target="media/([^"]+)"', rels))
rid_target.update({b: a for a, b in re.findall(r'Target="media/([^"]+)"[^>]*Id="(rId\d+)"', rels)})

def png_size(path):
    with open(path, 'rb') as f:
        head = f.read(24)
    return struct.unpack('>II', head[16:24])

def add_media(src, name):
    global rid_n, rels
    rid_n += 1
    shutil.copy(src, f'{D}/word/media/{name}')
    rels = rels.replace('</Relationships>',
        f'<Relationship Id="rId{rid_n}" Type="{R}/image" Target="media/{name}"/></Relationships>')
    return f'rId{rid_n}'

def cell_cap(cell):
    w = int(cell.find('w:tcPr/w:tcW', ns).get(q('w')))
    mar = cell.find('w:tcPr/w:tcMar', ns)
    side = 0
    if mar is not None:
        for s in ('left', 'right'):
            m = mar.find(f'w:{s}', ns)
            side += int(m.get(q('w'))) if m is not None else 108
    else:
        side = 216
    return int((w - side - 40) * 635)

def inline_xml(rid, cx, cy):
    pic_id[0] += 1
    i = pic_id[0]
    return (f'<w:r><w:drawing><wp:inline distT="0" distB="0" distL="0" distR="0">'
            f'<wp:extent cx="{cx}" cy="{cy}"/><wp:docPr id="{i}" name="图示 {i}"/>'
            '<a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">'
            f'<pic:pic><pic:nvPicPr><pic:cNvPr id="{i}" name="图示{i}.png"/><pic:cNvPicPr/></pic:nvPicPr>'
            f'<pic:blipFill><a:blip r:embed="{rid}"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill>'
            f'<pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="{cx}" cy="{cy}"/></a:xfrm>'
            '<a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr></pic:pic>'
            '</a:graphicData></a:graphic></wp:inline></w:drawing></w:r>')

def figure(rid, cx, cy, cap, caption=None):
    if cx > cap:
        cy = int(cy * cap / cx); cx = cap
    xml = para([inline_xml(rid, cx, cy)], jc='center', before=60, after=20 if caption else 60)
    if caption:
        xml += para([run(caption, 16, '666666')], jc='center', after=80)
    return [n for n in etree.fromstring(f'<x {NSDECL}>{xml}</x>')]

def new_figure(src, name, cell, caption=None, width_emu=None):
    rid = add_media(src, name)
    pw, ph = png_size(src)
    cap = cell_cap(cell)
    cx = min(width_emu or cap, cap)
    return figure(rid, cx, int(cx * ph / pw), cap, caption)

def drawings_by_media(root):
    out = {}
    for dr in root.iter(q('drawing')):
        blip = next(dr.iter(f'{{{ANS}}}blip'), None)
        if blip is None: continue
        name = rid_target.get(blip.get(f'{{{R}}}embed'))
        if name: out[name] = dr
    return out

def host_para(node):
    while node is not None and node.tag != q('p'):
        node = node.getparent()
    return node

def is_empty(p):
    return (not ''.join(x.text or '' for x in p.iter(q('t'))).strip()
            and p.find('.//w:drawing', ns) is None and p.find('.//w:pict', ns) is None)

def detach(dr):
    """移除一张图（所在 run）；返回其所在段落"""
    r = dr.getparent()
    while r.tag != q('r'): r = r.getparent()
    p = host_para(r)
    r.getparent().remove(r)
    return p

def find_para(cell, prefix):
    for p in cell.iter(q('p')):
        t = re.sub(r'\s+', '', ''.join(x.text or '' for x in p.iter(q('t'))))
        if t.startswith(prefix): return p
    raise KeyError(prefix)

def insert_after(anchor, nodes):
    for n in reversed(nodes):
        anchor.addnext(n)

def extent(dr):
    e = dr[0].find(f'{{{WPNS}}}extent')
    return int(e.get('cx')), int(e.get('cy'))

def cleanup(p):
    if p is not None and p.getparent() is not None and is_empty(p):
        cell = p.getparent()
        if len([x for x in cell if x.tag == q('p')]) > 1:
            cell.remove(p)

D_ = drawings_by_media(tbl)
tc_ = lambda ri, ci: rows[ri].findall('w:tc', ns)[ci]

# —— 教学流程：原流程图 → 新流程图 ——
cell = tc_(19, 1)
p = detach(D_['image1.png'])
insert_after(p, new_figure(f'{NEWIMG}/image20.png', 'fig_flow.png', cell, width_emu=5688000))
cleanup(p)

# —— 理论探究：原“投影法”导图 → 新分类图 ——
cell = tc_(27, 1)
p = detach(D_['image4.png'])
insert_after(p, new_figure(f'{NEWIMG}/image21.png', 'fig_fenlei.png', cell, '投影法分类'))
cleanup(p)

# —— 示范讲解 · 教学内容 ——
c1 = tc_(28, 1)
p = detach(D_['image5.png']); cleanup(p)                    # 图2-1 → 新三投影面体系图
insert_after(find_para(c1, '的交线OX'), new_figure(f'{NEWIMG}/image24.png', 'fig_tixi.png', c1, '图2-1　三投影面体系'))
host = find_para(c1, '三视图的形成：')                       # 图2-2a/b：浮动 → 嵌入
figs = []
for name in ('image6.png', 'image7.png'):
    dr = D_[name]; rid = next(dr.iter(f'{{{ANS}}}blip')).get(f'{{{R}}}embed'); cx, cy = extent(dr)
    detach(dr); figs += figure(rid, cx, cy, cell_cap(c1))
insert_after(host, figs)
dr = D_['image9.png']; rid = next(dr.iter(f'{{{ANS}}}blip')).get(f'{{{R}}}embed'); cx, cy = extent(dr)
host = detach(dr)                                            # 图2-2c：浮动 → 嵌入，放在说明文字之后
insert_after(host, figure(rid, cx, cy, cell_cap(c1)))
insert_after(find_para(c1, '俯、左视图宽相等'),
             new_figure(f'{NEWIMG}/image25.png', 'fig_sandeng.png', c1, '三视图的投影对应关系'))
insert_after(find_para(c1, '左视图反映物体的前、后和上、下'),
             new_figure(f'{SVGIMG}/fangwei.png', 'fig_fangwei.png', c1, '三视图的方位对应'))
host = detach(D_['image10.png'])                             # 直角弯板 → 新图
insert_after(host, new_figure(f'{NEWIMG}/image26.png', 'fig_wanban.png', c1, '带方槽与切角的直角弯板'))

# —— 示范讲解 · 教师活动 ——
c2 = tc_(28, 2)
for old, new, name, caption in (('image11.png', 'image22.png', 'fig_zhongxin.png', '中心投影与正投影'),
                                ('image13.png', 'image23.png', 'fig_texing.png', '正投影的三个基本特性'),
                                ('image19.png', 'image27.png', 'fig_sanshitu.png', '直角弯板三视图')):
    p = detach(D_[old])
    insert_after(p, new_figure(f'{NEWIMG}/{new}', name, c2, caption))
    cleanup(p)
for dup in ('image15.png', 'image18.png'):                   # 与教学内容列新图重复
    cleanup(detach(D_[dup]))

# —— 总结评价：知识框架图 ——
c = tc_(29, 1)
last = [p for p in c.findall('w:p', ns)][-1]
insert_after(last, new_figure(f'{SVGIMG}/kuangjia.png', 'fig_kuangjia.png', c, '本课知识框架',
                              width_emu=int(cell_cap(c) * 0.92)))

# —— 示范讲解各列：连续空行合并为一行，首尾空行去掉 ——
for ci in range(1, 5):
    cell = tc_(28, ci)
    ps = [x for x in cell if x.tag == q('p')]
    prev_empty = True
    for p in ps:
        if is_empty(p):
            if prev_empty: cell.remove(p)
            prev_empty = True
        else:
            prev_empty = False
    ps = [x for x in cell if x.tag == q('p')]
    while len(ps) > 1 and is_empty(ps[-1]):
        cell.remove(ps[-1]); ps.pop()

# 教学流程行不跨页（避免底部留白溢出到下一页形成空条）
_tp = rows[19].find('w:trPr', ns)
if _tp.find('w:cantSplit', ns) is None:
    _tp.insert(0, etree.Element(q('cantSplit')))

# 第 4 页（课前准备 / 情境导入 / 理论探究）均摊余量，让“示范讲解”从下一页顶部开始
E4 = int(os.environ.get('FILL_P4', '0'))
for ri in (26, 27):
    for c in rows[ri].findall('w:tc', ns)[1:]:
        mar = c.find('w:tcPr/w:tcMar', ns)
        for side in ('top', 'bottom'):
            m = mar.find(f'w:{side}', ns)
            m.set(q('w'), str(int(m.get(q('w'))) + E4))
for ri in (23,):
    _tp = rows[ri].find('w:trPr', ns)
    if _tp.find('w:cantSplit', ns) is None:
        _tp.insert(0, etree.Element(q('cantSplit')))
# 情境导入 / 理论探究 / 示范讲解 允许跨页，由内容自然衔接，避免整行下移留出大片空白
for ri in (26, 27, 28):
    _tp = rows[ri].find('w:trPr', ns)
    _cs = _tp.find('w:cantSplit', ns)
    if _cs is not None: _tp.remove(_cs)

# ================= 重画粗糙原图并替换 =================
def replace_img(cell, old, src, name, caption=None, after_prefix=None):
    p = detach(D_[old])
    nodes = new_figure(src, name, cell, caption)
    insert_after(find_para(cell, after_prefix) if after_prefix else p, nodes)
    cleanup(p)

c1 = tc_(28, 1)
# 图2-2a / 2-2b：原图 7、6（8 与 7 重复，删去），放在“三视图的形成：”之后；图2-2c：原图 9
old_figs = [n for n in c1.iter(q('drawing'))]
for dr in list(c1.iter(q('drawing'))):
    blip = next(dr.iter(f'{{{ANS}}}blip'), None)
    if blip is not None and rid_target.get(blip.get(f'{{{R}}}embed')) in ('image6.png', 'image7.png', 'image8.png', 'image9.png'):
        cleanup(detach(dr))
insert_after(find_para(c1, '三视图的形成：'),
             new_figure(f'{SVGIMG}/fig_A.png', 'fig_2_2a.png', c1, '图2-2a　物体在三投影面体系中的投影')
             + new_figure(f'{SVGIMG}/fig_B.png', 'fig_2_2b.png', c1, '图2-2b　三面投影与投影面展开'))
insert_after(find_para(c1, '定正面不动'),
             new_figure(f'{SVGIMG}/fig_C.png', 'fig_2_2c.png', c1, '图2-2c　三视图'))

c2 = tc_(28, 2)
replace_img(c2, 'image12.png', f'{SVGIMG}/fig_E.png', 'fig_zhengxie.png', '正投影与斜投影')
replace_img(c2, 'image14.png', f'{SVGIMG}/fig_F.png', 'fig_danmian.png', '单面投影不能确定形状')
replace_img(c2, 'image16.png', f'{SVGIMG}/fig_D.png', 'fig_zhankai.png', '投影面展开')

# ================= 第 1 页：学习内容三步配小图 =================
inner = rows[6].findall('w:tc', ns)[1].find('w:tbl', ns)
for tcx in inner.iter(q('tc')):
    t = re.sub(r'\s+', '', ''.join(tcx.itertext()))
    for key, src, nm in (('01投影法分类', f'{NEWIMG}/image21.png', 'thumb1.png'),
                         ('02三投影面体系', f'{NEWIMG}/image24.png', 'thumb2.png'),
                         ('03三视图投影对应关系', f'{SVGIMG}/fig_C.png', 'thumb3.png')):
        if t.startswith(key):
            ps = tcx.findall('w:p', ns)
            rid = add_media(src, nm)
            pw, ph = png_size(src)
            hh = 800000                                   # 缩略图统一高度 ≈ 2.2 cm
            cx, cy = int(hh * pw / ph), hh
            cap = cell_cap(tcx)
            if cx > cap: cy = int(cy * cap / cx); cx = cap
            ps[1].addnext(etree.fromstring(f'<x {NSDECL}>' + para([inline_xml(rid, cx, cy)], jc='center', before=60, after=40) + '</x>')[0])

# ================= 第 2 页：重点 / 难点 左文右图（原内容格按主表网格拆为 文字 5 列 + 图 2 列） =================
for ri, src, nm, caption in ((13, f'{NEWIMG}/image25.png', 'fig_zhongdian.png', '三等投影规律'),
                             (14, f'{SVGIMG}/fig_B.png', 'fig_nandian.png', '投影面展开')):
    c_txt = rows[ri].findall('w:tc', ns)[1]
    pr = c_txt.find('w:tcPr', ns)
    pr.find('w:tcW', ns).set(q('w'), '6009')
    pr.find('w:gridSpan', ns).set(q('val'), '5')
    img_cell = etree.fromstring(f'<x {NSDECL}>' + tc(3599, '', 'FFFFFF', span=2, bdr=ALL_THIN,
                                mar={'top': 80, 'left': 100, 'bottom': 80, 'right': 100}) + '</x>')[0]
    c_txt.addnext(img_cell)
    for n_ in new_figure(src, nm, img_cell, caption, width_emu=int(cell_cap(img_cell) * 0.9)):
        img_cell.append(n_)
    # 图高受限：保持行高合理
    for dr in img_cell.iter(q('drawing')):
        e = dr[0].find(f'{{{WPNS}}}extent'); cx_, cy_ = int(e.get('cx')), int(e.get('cy'))
        mh = 1250000
        if cy_ > mh:
            nx = int(cx_ * mh / cy_); e.set('cx', str(nx)); e.set('cy', str(mh))
            ext = next(dr.iter(f'{{{ANS}}}ext')); ext.set('cx', str(nx)); ext.set('cy', str(mh))

# 第 4 页（教学流程 + 第一阶段 课前准备）均摊余量
E5 = int(os.environ.get('FILL_P5', '0'))
for ri in (19, 23):
    for c in rows[ri].findall('w:tc', ns)[1:]:
        mar = c.find('w:tcPr/w:tcMar', ns)
        for side in ('top', 'bottom'):
            m = mar.find(f'w:{side}', ns)
            m.set(q('w'), str(int(m.get(q('w'))) + E5))
