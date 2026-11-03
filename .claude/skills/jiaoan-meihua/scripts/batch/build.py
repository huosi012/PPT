"""教案蓝金美化（通用模板版，按行首标签定位，不改主表行列/合并结构）。
用法：python3 -I build.py <源解包目录> <输出解包目录> <配置.py>
环境变量：ROW_MARK=1 时每行加不可见定位码（仅 fill2 测量用）；FILL_PLAN='[[起始行,结束行,增量twips],...]'  REFLECT_H=末页反思行高
"""
import sys, os, re, shutil, struct, json
TOOLS = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, TOOLS)
sys.path.insert(0, os.path.join(TOOLS, '..'))          # 本 skill 的 scripts/docx_kit.py
from docx_kit import *
from lxml import etree

SRC, DST, CFGF = sys.argv[1:4]
ICONS = os.path.join(TOOLS, '..', '..', 'assets', 'icons')
_cfg = {}
exec(open(CFGF, encoding='utf8').read(), _cfg)
CFG = _cfg['CFG']
if os.path.exists(DST): shutil.rmtree(DST)
shutil.copytree(SRC, DST)
d = Doc(DST, icon_dir=ICONS)
body = d.body
tbl = d.main_table()
rows = tbl.findall('w:tr', ns)
grid = [int(g.get(q('w'))) for g in tbl.find('w:tblGrid', ns)]
TW = sum(grid)
WPNS = 'http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing'
ANS = 'http://schemas.openxmlformats.org/drawingml/2006/main'
MCNS = 'http://schemas.openxmlformats.org/markup-compatibility/2006'
NOTBL = borders('tblBorders', {s: NONE for s in ('top', 'left', 'bottom', 'right', 'insideH', 'insideV')})
NAVY_B = borders('tcBorders', {s: ('single', 4, C.NAVY) for s in SIDES})
LABEL_CONT_B = borders('tcBorders', {'top': NONE, 'left': GOLD_BAR, 'bottom': THIN, 'right': THIN})


# ---------------- 基础 ----------------
_SKIP = None
def ptext(p):
    """段落正文文字：不含图片/文本框/VML 里的文字"""
    out = []
    for x in p.iter(q('t')):
        a, inside = x.getparent(), False
        while a is not None and a is not p:
            ln = etree.QName(a).localname
            if ln in ('drawing', 'pict', 'AlternateContent', 'txbxContent'): inside = True; break
            a = a.getparent()
        if not inside: out.append(x.text or '')
    return ''.join(out)

def ntext(node):
    return re.sub(r'\s+', '', ''.join(x.text or '' for x in node.iter(q('t'))))

def cells(r):
    return r.findall('w:tc', ns)

def width(c):
    return int(c.find('w:tcPr/w:tcW', ns).get(q('w')))

def lines(c):
    """单元格直接段落的文字（去首尾空白，去空行）"""
    return [ptext(p).strip() for p in c.findall('w:p', ns) if ptext(p).strip()]

def find(prefix, start=0):
    for i in range(start, len(rows)):
        if ntext(cells(rows[i])[0]).startswith(prefix): return i
    raise KeyError(prefix)

def is_label(c):
    shd = c.find('w:tcPr/w:shd', ns)
    return shd is not None and shd.get(q('fill')) in ('C6D9F1', '0070C0')

def trh(r, val, rule='atLeast'):
    trpr = r.find('w:trPr', ns)
    if trpr is None:
        trpr = etree.Element(q('trPr')); r.insert(0, trpr)
    h = trpr.find('w:trHeight', ns)
    if h is None: h = etree.SubElement(trpr, q('trHeight'))
    h.set(q('val'), str(val)); h.set(q('hRule'), rule)

def keep_next(r):
    for p in r.iter(q('p')):
        ppr = p.find('w:pPr', ns)
        if ppr is None: ppr = etree.Element(q('pPr')); p.insert(0, ppr)
        put(ppr, 'keepNext', PORD, etree.Element(q('keepNext')))

def strip_num(t):
    return re.sub(r'^\s*(\d+\s*[\.．、]|[①-⑩]|（\d+）|\(\d+\))\s*', '', t)

def rich(s, sz=20, color=C.TEXT):
    """**粗体藏蓝** [[金色标签]] {{深红粗体}}"""
    out = []
    for part in re.split(r'(\*\*.+?\*\*|\[\[.+?\]\]|\{\{.+?\}\})', s):
        if not part: continue
        if part.startswith('**'): out.append(run(part[2:-2], sz, C.NAVY, b=True))
        elif part.startswith('[['): out.append(chip(part[2:-2], sz - 1))
        elif part.startswith('{{'): out.append(run(part[2:-2], sz, C.RED, b=True))
        else: out.append(run(part, sz, color))
    return out

def png_size(path):
    with open(path, 'rb') as f: h = f.read(24)
    return struct.unpack('>II', h[16:24])

def fig_xml(src, name, max_w, caption=None, max_h=None):
    pw, ph = png_size(src)
    rid = d.add_media(src, name)
    cx = max_w; cy = int(cx * ph / pw)
    if max_h and cy > max_h: cy = max_h; cx = int(cy * pw / ph)
    xml = para([d.inline(rid, cx, cy)], jc='center', before=60, after=20 if caption else 60)
    if caption: xml += para([run(caption, 16, '666666')], jc='center', after=60)
    return xml

def cap_emu(c, side_mar=None):
    pr = c.find('w:tcPr', ns)
    mar = pr.find('w:tcMar', ns)
    side = 0
    for s in ('left', 'right'):
        m = mar.find(f'w:{s}', ns) if mar is not None else None
        side += int(m.get(q('w'))) if m is not None else 108
    if side_mar is not None: side = side_mar
    return int((width(c) - side - 60) * 635)


# ================= 1. 页首横幅（标题段 → 横幅小表，标题文字保留） =================
title_p = body.find('w:p', ns)
title = ptext(title_p).strip() or '《机械制图》教学设计'
LW = 1700
banner = (f'<w:tbl><w:tblPr><w:tblW w:w="{TW}" w:type="dxa"/><w:jc w:val="center"/>{NOTBL}'
          '<w:tblLayout w:type="fixed"/><w:tblCellMar><w:left w:w="0" w:type="dxa"/><w:right w:w="0" w:type="dxa"/></w:tblCellMar></w:tblPr>'
          f'<w:tblGrid><w:gridCol w:w="{LW}"/><w:gridCol w:w="{TW - LW}"/></w:tblGrid>'
          '<w:tr><w:trPr><w:trHeight w:val="1760" w:hRule="exact"/><w:jc w:val="center"/></w:trPr>'
          + tc(LW, para([d.icon('ruler-measure-FFFFFF', 54)], jc='center'), fill=C.NAVY)
          + tc(TW - LW,
               para([run(CFG.get('eyebrow_a', '机械类专业核心基础课'), 20, C.GOLD, b=True, spacing=20),
                     run('  ·  ' + CFG.get('eyebrow_b', '理实一体化教学设计'), 20, 'B4C3DA', spacing=20)], after=60)
               + para([run(title, 48, 'FFFFFF', b=True, spacing=30)], after=60)
               + para([run(CFG['subtitle'], 24, 'D6DEEB', spacing=10)]),
               fill=C.NAVY, mar={'left': 60})
          + '</w:tr><w:tr><w:trPr><w:trHeight w:val="70" w:hRule="exact"/><w:jc w:val="center"/></w:trPr>'
          + tc(TW, para([run('', 2)], line=20), fill=C.GOLD, span=2) + '</w:tr></w:tbl>')
body.replace(title_p, el(banner))
body.find('w:tbl', ns).addnext(el(para([run('', 8)], line=240)))


# ================= 2. 分区标题栏（居中） =================
SEC = {'一、': ('info-circle-FFFFFF', '01'), '二、': ('chalkboard-FFFFFF', '02'),
       '三、': ('presentation-FFFFFF', '03'), '四、': ('message-2-FFFFFF', '04')}
sec_rows = []
for i, r in enumerate(rows):
    cs = cells(r)
    t = ntext(cs[0])
    if len(cs) == 1 and t[:2] in SEC and is_label(cs[0]):
        ic, num = SEC[t[:2]]
        set_tcpr(cs[0], fill=C.NAVY, bdr_xml=NAVY_B, valign='center')
        set_paras(cs[0], para([d.icon(ic, 15, lower=6), run(f'  {num}', 26, C.GOLD, b=True),
                               run(f'  {t[2:]}', 26, 'FFFFFF', b=True, spacing=40)], jc='center'))
        trh(r, 560); sec_rows.append(i)


# ================= 3. 授课信息 =================
for key in ('模块名称', '任务名称', '授课时数', '授课班级'):
    r = rows[find(key)]
    for c in cells(r):
        if is_label(c):
            set_tcpr(c, fill=C.LABEL, bdr_xml=ALL_THIN, valign='center')
            set_paras(c, para([run(ntext(c), 22, C.NAVY, b=True)], jc='center'))
            continue
        ls = [re.sub(r'\s{2,}', '  ', x) for x in lines(c)]
        if key == '模块名称' and len(ls) == 2 and len(ls[0]) <= 5 and len(ls[0]) + len(ls[1]) <= 22:
            ls = [ls[0] + '  ' + ls[1]]          # “第五章 / 组合体的绘制与识读”合为一行
        set_tcpr(c, fill='FFFFFF', bdr_xml=ALL_THIN, valign='center')
        if not ls:
            xml = para([run('待填写', 21, C.HINT)], jc='center')
        elif len(ls) == 1 and re.fullmatch(r'(\d+)\s*(学时|课时)', ls[0]):
            m = re.fullmatch(r'(\d+)\s*(学时|课时)', ls[0])
            xml = para([run(m.group(1), 28, C.NAVY, b=True), run(' ' + m.group(2), 22, C.TEXT)], jc='center')
        elif len(ls) == 1 and ls[0] in ('理实一体', '理论', '实践', '实训'):
            xml = para([chip(ls[0], 21, C.NAVY, 'FFFFFF')], jc='center')
        else:
            xml = ''.join(para([run(x, 22, C.TEXT)], jc='center', after=20) for x in ls)
        set_paras(c, xml)
    trh(r, 560)


# ================= 4. 大类 / 小类标签 =================
def group_label(c, icon_name, text, note=None):
    ls = [text[:2], text[2:]] if len(text) == 4 else [text]
    set_tcpr(c, fill=C.LABEL, bdr_xml=LABEL_B, valign='center')
    top = para([run(note, 16, '8A99B3')], jc='center', after=200) if note else ''
    set_paras(c, top + para([d.icon(icon_name, 22)], jc='center', after=60)
              + ''.join(para([run(x, 24, C.NAVY, b=True, spacing=40)], jc='center') for x in ls))

def group_continue(c):
    set_tcpr(c, fill=C.LABEL, bdr_xml=LABEL_CONT_B)

def single_label(c, icon_name, text):
    set_tcpr(c, fill=C.LABEL, bdr_xml=LABEL_B, valign='center')
    ls = [text[:2], text[2:]] if len(text) == 4 else [text]
    set_paras(c, para([d.icon(icon_name, 18)], jc='center', after=40)
              + ''.join(para([run(x, 22, C.NAVY, b=True, spacing=20)], jc='center') for x in ls))

def sub_label(c, text):
    set_tcpr(c, fill=C.CARD, bdr_xml=ALL_THIN, valign='center')
    set_paras(c, para([run(text, 22, C.NAVY, b=True)], jc='center'))

CONTENT_MAR = {'top': 90, 'left': 180, 'bottom': 90, 'right': 180}
def content(c, xml, valign='center', mar=CONTENT_MAR):
    set_tcpr(c, fill='FFFFFF', bdr_xml=ALL_THIN, valign=valign, mar=mar)
    set_paras(c, xml)


# ================= 5. 内容分析：卡片（贴合外层单元格） =================
def build_cards(c, cards):
    CW = width(c)
    N = 12
    COL = CW // N
    G = [COL] * (N - 1) + [CW - COL * (N - 1)]
    MAR = {'top': 110, 'left': 200, 'bottom': 120, 'right': 200}
    def edges(top=THIN, left=THIN, bottom=THIN, right=THIN):
        return borders('tcBorders', {'top': top, 'left': left, 'bottom': bottom, 'right': right})
    def cell(i0, span, xml, fill=C.CARD, bdr_xml=None, mar=MAR):
        return tc(sum(G[i0:i0 + span]), xml, fill, span=span, valign='top', bdr=bdr_xml, mar=mar)
    def head(text, color=C.NAVY, icon_name=None):
        lead = [d.icon(icon_name, 14, lower=5), run('  ', 20)] if icon_name else [run('▍', 22, C.GOLD if color == C.NAVY else color)]
        return para(lead + [run(' ' + text if not icon_name else text, 24, color, b=True)], after=80)
    def blocks(bl, color=C.TEXT):
        out = ''
        for kind, s in bl:
            if kind == 'lead': out += para([run(s, 19, C.SUB)], after=40, line=300)
            elif kind == 'b': out += para([run('■\t', 14, '8A99B3', pos=2)] + rich(s, 20, color), jc='both',
                                          line=300, after=30, ind_left=240, hanging=240)
            else: out += para(rich(s, 20, color), jc='both', line=300, after=30)
        return out
    def card_xml(cd):
        red = cd.get('red')
        return head(cd['title'], C.RED if red else C.NAVY, 'flag-3-9B1C1C' if red else None) + blocks(cd['blocks'])
    N_ = NONE
    rows_xml = []
    # 第 1 行：定位 | 标准
    a, b = cards['pos'], cards['std']
    rows_xml.append(cell(0, 6, card_xml(a), bdr_xml=edges(top=N_, left=N_)) +
                    cell(6, 6, card_xml(b), bdr_xml=edges(top=N_, left=N_, right=N_)))
    # 第 2 行：学习内容（标题 + 横排步骤）
    L = cards['learn']
    rows_xml.append(cell(0, 12, head(L['title']) + blocks(L.get('blocks', [])),
                         bdr_xml=edges(left=N_, right=N_, bottom=N_), mar={**MAR, 'bottom': 40}))
    steps = L['steps']; n = len(steps)
    spans = {1: [12], 2: [6, 6], 3: [4, 4, 4], 4: [3, 3, 3, 3], 5: [2, 3, 2, 3, 2], 6: [2] * 6}[n]
    has_img = any(stp.get('img') for stp in steps)
    sx, ix, i0 = '', '', 0
    for k, (stp, sp) in enumerate(zip(steps, spans)):
        x = para([run(f'{k + 1:02d}', 26, C.GOLD_D, b=True)], after=20) + para(rich(stp['t'], 19), line=290)
        lr = dict(left=N_ if k == 0 else THIN, right=N_ if k == n - 1 else THIN)
        sx += cell(i0, sp, x, bdr_xml=edges(top=N_, bottom=N_ if has_img else THIN, **lr),
                   mar={'top': 40, 'left': 200, 'bottom': 40 if has_img else 140, 'right': 160})
        if has_img:                          # 缩略图单独成行、垂直居中，各列图片对齐
            w = int((sum(G[i0:i0 + sp]) - 360) * 635)
            fx = (fig_xml(stp['img'], f'thumb_{k}.png', w, max_h=int(L.get('img_h', 2.54) / 2.54 * 914400))
                  if stp.get('img') else para([run('', 10)]))
            ix += tc(sum(G[i0:i0 + sp]), fx, C.CARD, span=sp, valign='center', bdr=edges(top=N_, **lr),
                     mar={'top': 0, 'left': 200, 'bottom': 100, 'right': 160})
        i0 += sp
    rows_xml.append(sx)
    if has_img: rows_xml.append(ix)
    if L.get('figure'):                      # 学习内容配图（通栏）
        F = L['figure']
        fx = fig_xml(F['img'], 'fig_learn.png', int((CW - 400) * 635), F.get('caption'), max_h=int(F.get('h', 5) / 2.54 * 914400))
        rows_xml.append(cell(0, 12, fx, bdr_xml=edges(left=N_, right=N_), mar={'top': 60, 'left': 200, 'bottom': 60, 'right': 200}))
    # 第 3 行：四新 | 思政
    a, b = cards['new'], cards['szh']; b = {**b, 'red': True}
    rows_xml.append(cell(0, 6, card_xml(a), bdr_xml=edges(left=N_, bottom=N_)) +
                    cell(6, 6, card_xml(b), fill=C.WARM,
                         bdr_xml=borders('tcBorders', {'top': ('single', 12, C.RED), 'left': THIN, 'bottom': N_, 'right': N_})))
    t = (f'<w:tbl><w:tblPr><w:tblW w:w="{CW}" w:type="dxa"/><w:tblInd w:w="0" w:type="dxa"/>{NOTBL}'
         '<w:tblLayout w:type="fixed"/><w:tblCellMar><w:left w:w="0" w:type="dxa"/><w:right w:w="0" w:type="dxa"/></w:tblCellMar></w:tblPr>'
         '<w:tblGrid>' + ''.join(f'<w:gridCol w:w="{g}"/>' for g in G) + '</w:tblGrid>'
         + ''.join(f'<w:tr><w:trPr><w:cantSplit/></w:trPr>{x}</w:tr>' for x in rows_xml) + '</w:tbl>')
    set_tcpr(c, fill='FFFFFF', bdr_xml=ALL_THIN, valign='top', mar={'top': 0, 'left': 0, 'bottom': 0, 'right': 0})
    set_paras(c, '')
    for n_ in els(t + para([run('', 2)], line=20)): c.append(n_)

ri = find('内容分析')
c0, c1 = cells(rows[ri])
group_label(c0, 'file-analytics-1F3864', '内容分析')
build_cards(c1, CFG['cards'])


# ================= 6. 学情 / 目标 / 重点难点 / 方法 / 资源 =================
def split_semicolon(s):
    return [p.strip() for p in re.split(r'(?<=；)', s) if p.strip()]

def split_inline_numbers(t):
    parts = re.split(r'(?:^|(?<=；)|(?<=。))\s*(\d+)\s*[\.．、]\s*', t)
    out = []
    if parts and parts[0].strip(): out.append(parts[0].strip())
    for i in range(1, len(parts) - 1, 2): out.append(parts[i + 1].strip())
    return [x for x in out if x]

def bullets(items):
    return ''.join(para([run('■\t', 14, '8A99B3', pos=2)] + rich(x, 21), jc='both', ind_left=240, hanging=240,
                        line=320, after=40) for x in items)

def numbered(items, hl=()):
    out = ''
    for i, x in enumerate(items, 1):
        x = strip_num(x)
        for h in hl: x = x.replace(h, f'[[{h.strip("“”")}]]')
        out += para([run(f'{i}\t', 21, C.GOLD_D, b=True)] + rich(x, 21), jc='both', ind_left=300, hanging=300,
                    line=320, after=40)
    return out

def items_of(c):
    ls = lines(c)
    if len(ls) == 1 and re.search(r'[；;]\s*\d+\s*[\.．、]', ls[0]):
        return split_inline_numbers(ls[0])
    out = []
    for x in ls:
        if len(ls) == 1 and '；' in x.rstrip('；。'):
            out += split_semicolon(x)
        else:
            out.append(x)
    return out

ri = find('学情分析')
for k in range(3):
    cs = cells(rows[ri + k])
    if k == 0: group_label(cs[0], 'users-group-1F3864', '学情分析', note=CFG.get('cont_note'))
    else: group_continue(cs[0])
    sub_label(cs[1], ntext(cs[1]))
    content(cs[2], bullets([y for x in lines(cs[2]) for y in split_semicolon(x)]))

ri = find('教学目标')
for k in range(3):
    cs = cells(rows[ri + k])
    if k == 0: group_label(cs[0], 'target-1F3864', '教学目标')
    else: group_continue(cs[0])
    sub_label(cs[1], ntext(cs[1]))
    content(cs[2], numbered(items_of(cs[2])))

for key, ic, hl in (('教学重点', 'star-1F3864', CFG.get('key_hl', ())), ('教学难点', 'alert-triangle-1F3864', ())):
    c0, c1 = cells(rows[find(key)])
    single_label(c0, ic, key)
    content(c1, numbered(items_of(c1), hl))
    if key == '教学难点' and CFG.get('diff_fig'):      # 难点图解
        F = CFG['diff_fig']
        for n_ in els(fig_xml(F['img'], 'fig_nandian.png', int(cap_emu(c1) * 0.9), F.get('caption'),
                              max_h=int(F.get('h', 5) / 2.54 * 914400))):
            c1.append(n_)

def split_dun(raw):
    items, buf, depth = [], '', 0
    for ch in raw:
        if ch in '（(': depth += 1
        if ch in '）)': depth -= 1
        if ch in '、，,' and depth == 0: items.append(buf); buf = ''
        else: buf += ch
    items.append(buf)
    return [x.strip() for x in items if x.strip()]

for key, ic in (('教学方法', 'bulb-1F3864'), ('学习方法', 'users-group-1F3864')):
    c0, c1 = cells(rows[find(key)])
    single_label(c0, ic, key)
    rs = []
    for i, x in enumerate(split_dun(''.join(lines(c1)).rstrip('。；'))):
        if i: rs.append(run('　', 21))
        rs.append(chip(x, 21, C.LABEL, C.NAVY))
    content(c1, para(rs, line=420))

c0, c1 = cells(rows[find('教学资源')])
single_label(c0, 'books-1F3864', '教学资源')
xml = ''
for x in lines(c1):
    m = re.match(r'【(.+?)】(.*)', x)
    if m:
        xml += para([chip(m.group(1).replace(' ', ''), 19, C.NAVY, 'FFFFFF'), run('　', 21)] + rich(m.group(2).strip(), 21),
                    jc='both', line=320, after=60)
    else:
        xml += para(rich(x, 21), jc='both', line=320, after=60)
content(c1, xml)


# ================= 7. 考核方式（保留原嵌套评价表，改为贴合样式） =================
def restyle_assess(c):
    inner = c.find('w:tbl', ns)
    cw = width(c)
    # 嵌套表宽度 = 外层单元格宽度，按比例缩放列宽
    g = inner.find('w:tblGrid', ns); gc = g.findall('w:gridCol', ns)
    old = [int(x.get(q('w'))) for x in gc]; f = cw / sum(old)
    new = [round(v * f) for v in old]; new[-1] += cw - sum(new)
    for x, v in zip(gc, new): x.set(q('w'), str(v))
    tpr = inner.find('w:tblPr', ns)
    for tag in ('tblW', 'tblInd', 'tblBorders', 'tblCellMar', 'jc'):
        o = tpr.find(f'w:{tag}', ns)
        if o is not None: tpr.remove(o)
    tpr.append(el(f'<w:tblW w:w="{cw}" w:type="dxa"/>'))
    tpr.append(el('<w:tblInd w:w="0" w:type="dxa"/>'))
    tpr.append(el(NOTBL))
    tpr.append(el('<w:tblCellMar><w:left w:w="0" w:type="dxa"/><w:right w:w="0" w:type="dxa"/></w:tblCellMar>'))
    # 按 schema 顺序重排 tblPr
    order = ['tblStyle', 'tblpPr', 'tblOverlap', 'bidiVisual', 'tblStyleRowBandSize', 'tblStyleColBandSize', 'tblW', 'jc',
             'tblCellSpacing', 'tblInd', 'tblBorders', 'shd', 'tblLayout', 'tblCellMar', 'tblLook']
    kids = sorted(list(tpr), key=lambda e: order.index(etree.QName(e).localname) if etree.QName(e).localname in order else 99)
    for k_ in list(tpr): tpr.remove(k_)
    for k_ in kids: tpr.append(k_)
    intro = [p for p in c if p.tag == q('p') and ptext(p).strip() and c.index(p) < c.index(inner)]
    trs = inner.findall('w:tr', ns)
    for ri_, tr in enumerate(trs):
        tcs = tr.findall('w:tc', ns)
        col = 0
        for j, x in enumerate(tcs):
            pr = x.find('w:tcPr', ns)
            gs = pr.find('w:gridSpan', ns); span = int(gs.get(q('val'))) if gs is not None else 1
            pr.find('w:tcW', ns).set(q('w'), str(sum(new[col:col + span])))
            last_col = col + span >= len(new)
            b = {'top': (THIN if (ri_ > 0 or intro) else NONE), 'left': NONE if col == 0 else THIN,
                 'bottom': NONE if ri_ == len(trs) - 1 else THIN, 'right': NONE if last_col else THIN}
            col += span
            ls = lines(x)
            if ri_ == 0:
                set_tcpr(x, fill=C.LABEL, bdr_xml=borders('tcBorders', b), valign='center', mar={'top': 60, 'left': 100, 'bottom': 60, 'right': 100})
                set_paras(x, ''.join(para([run(t, 20, C.NAVY, b=True)], jc='center', line=280) for t in ls) or para([run('', 20)]))
            elif j == 0:
                set_tcpr(x, fill=C.CARD, bdr_xml=borders('tcBorders', b), valign='center', mar={'top': 60, 'left': 100, 'bottom': 60, 'right': 100})
                set_paras(x, ''.join(para([run(t, 20, C.NAVY, b=True)], jc='center', line=280) for t in ls) or para([run('', 20)]))
            else:
                set_tcpr(x, fill='FFFFFF', bdr_xml=borders('tcBorders', b), valign='center', mar={'top': 60, 'left': 140, 'bottom': 60, 'right': 120})
                set_paras(x, ''.join(para([run(t, 20, C.TEXT)], line=280, after=20) for t in ls) or para([run('', 20)]))
    # 外层：说明文字 + 贴边的评价表
    for p in [p for p in c if p.tag == q('p')]: c.remove(p)
    pos = 0
    for t in [ptext(p).strip() for p in intro]:
        c.insert(pos + 1, el(para(rich(t, 21), jc='both', line=320, before=80, after=80, ind_left=160)))
        pos += 1
    set_tcpr(c, fill='FFFFFF', bdr_xml=ALL_THIN, valign='top', mar={'top': 0, 'left': 0, 'bottom': 0, 'right': 0})
    c.append(el(para([run('', 2)], line=20)))

ri = find('考核方式')
c0, c1 = cells(rows[ri])
single_label(c0, 'clipboard-check-1F3864', '考核方式')
restyle_assess(c1)


# ================= 8. 图片工具：浮动 → 嵌入，按单元格宽度缩放 =================
def drawing_size(dr):
    e = dr[0].find(f'{{{WPNS}}}extent')
    return int(e.get('cx')), int(e.get('cy'))

WPG = 'http://schemas.microsoft.com/office/word/2010/wordprocessingGroup'
def set_drawing_size(dr, cx, cy):
    e = dr[0].find(f'{{{WPNS}}}extent'); e.set('cx', str(cx)); e.set('cy', str(cy))
    # 只改顶层图形（图片 / 组合 / 形状）自身的外框尺寸；组合内子图形保持原坐标，由 chExt 自动缩放
    gd = dr.find(f'.//{{{ANS}}}graphicData')
    top = gd[0] if gd is not None and len(gd) else None
    if top is not None:
        for pr in top:
            if etree.QName(pr).localname in ('spPr', 'grpSpPr'):
                x = pr.find(f'{{{ANS}}}xfrm/{{{ANS}}}ext')
                if x is not None: x.set('cx', str(cx)); x.set('cy', str(cy))

def to_inline(dr):
    a = dr[0]
    if etree.QName(a).localname == 'inline': return dr
    new = etree.Element(f'{{{WPNS}}}inline')
    for k in ('distT', 'distB', 'distL', 'distR'): new.set(k, '0')
    for tag in ('extent', 'docPr', 'cNvGraphicFramePr'):
        x = a.find(f'{{{WPNS}}}{tag}')
        if x is not None: new.append(x)
    new.append(a.find(f'{{{ANS}}}graphic'))
    dr.remove(a); dr.append(new)
    return dr

def img_paras(drs, cap):
    """一组图片放成若干行；能并排的并排，超宽的等比缩小"""
    sized = []
    for dr in drs:
        to_inline(dr)
        cx, cy = drawing_size(dr)
        if cx > cap: cy = int(cy * cap / cx); cx = cap
        sized.append([dr, cx, cy])
    rows_, cur, w = [], [], 0
    for s in sized:
        if len(cur) == 1 and w + s[1] + 60000 > cap and (w + s[1] + 60000) * 0.82 <= cap:
            f = (cap - 2 * 60000) / (cur[0][1] + s[1])     # 两张小图略缩后并排
            for t_ in (cur[0], s): t_[1], t_[2] = int(t_[1] * f), int(t_[2] * f)
            cur.append(s); rows_.append(cur); cur, w = [], 0
            continue
        if cur and w + s[1] + 60000 > cap: rows_.append(cur); cur, w = [], 0
        cur.append(s); w += s[1] + 60000
    if cur: rows_.append(cur)
    out = []
    for rw in rows_:
        p = el('<w:p><w:pPr><w:snapToGrid w:val="0"/><w:spacing w:before="60" w:after="60" w:line="240" w:lineRule="auto"/>'
               '<w:ind w:firstLine="0"/><w:jc w:val="center"/></w:pPr></w:p>')
        for dr, cx, cy in rw:
            set_drawing_size(dr, cx, cy)
            r_ = etree.SubElement(p, q('r')); r_.append(dr)
            if len(rw) > 1:
                t_ = etree.SubElement(r_, q('t')); t_.text = ' '
                t_.set('{http://www.w3.org/XML/1998/namespace}space', 'preserve')
        out.append(p)
    return out


# ================= 9. 教学流程 =================
ri = find('教学流程')
c0, c1 = cells(rows[ri])
single_label(c0, 'route-1F3864', '教学流程')
seq = []
for p in [p for p in c1 if p.tag == q('p')]:
    drs = [x for x in p.iter(q('drawing'))]
    t = ptext(p).strip()
    if t: seq.append(el(para(rich(t, 21), jc='both', line=320, before=40, after=60)))
    if drs:
        cap_ = cap_emu(c1, 160)
        if CFG.get('flow_h'):                   # 流程图限高（为了与上方各行同页）
            cx_, cy_ = drawing_size(drs[0])
            cap_ = min(cap_, int(CFG['flow_h'] / 2.54 * 914400 * cx_ / cy_))
        seq += img_paras(drs, cap_)
set_tcpr(c1, fill='FFFFFF', bdr_xml=ALL_THIN, valign='center', mar={'top': 80, 'left': 80, 'bottom': 80, 'right': 80})
for p in [p for p in c1 if p.tag == q('p')]: c1.remove(p)
for n_ in seq: c1.append(n_)


# ================= 10. 三、教学实施 =================
r_impl = find('三、') if False else next(i for i in sec_rows if ntext(cells(rows[i])[0]).endswith('教学实施'))
r_end = next(i for i in sec_rows if ntext(cells(rows[i])[0]).endswith('教学反思'))
TYPO = CFG.get('typo', [])
BODY_SZ = 19
INTENT_BG, INTENT_FG = C.INTENT_BG, C.INTENT_FG

def P_body(text, color=C.TEXT):
    m = re.match(r'^([^：:，。；、（）()]{2,14})([：:])(.*)$', text)
    if m and m.group(3).strip():
        return para([run(m.group(1) + m.group(2), BODY_SZ, C.NAVY, b=True)] + rich(tagify(m.group(3)), BODY_SZ, color),
                    jc='left', line=290, after=40)
    return para(rich(tagify(text), BODY_SZ, color), jc='left', line=290, after=40)

def P_item(num, text, color=C.TEXT):
    w = 300 if len(num) <= 2 else 440
    m = re.match(r'^([^：:，。；、（）()]{2,14})([：:])(.*)$', text)
    if m and m.group(3).strip():
        body_runs = [run(m.group(1) + m.group(2), BODY_SZ, C.NAVY, b=True)] + rich(tagify(m.group(3)), BODY_SZ, color)
    else:
        body_runs = rich(tagify(text), BODY_SZ, color)
    return para([run(num + '\t', BODY_SZ, C.GOLD_D, b=True)] + body_runs, jc='left', line=290, after=40,
                ind_left=w, hanging=w)

def P_h2(text):
    return para([run('▍', 20, C.GOLD), run(text, 21, C.NAVY, b=True)], line=260, before=100, after=50)

def P_h3(num, text):
    return para(([run(num + ' ', 20, C.GOLD_D, b=True)] if num else []) + [run(text, 20, C.NAVY, b=True)],
                line=260, before=60, after=30)

def tagify(s):
    s = re.sub(r'【(重点|难点|重点·难点|重难点)】', lambda m: f'[[{m.group(1)}]]', s)
    return re.sub(r'★\s*(重点|难点)', lambda m: f'[[★{m.group(1)}]]', s)

def clean(t):
    t = t.replace('　', ' ').strip()
    t = re.sub(r'\s{2,}', ' ', t)
    for a, b in TYPO: t = t.replace(a, b)
    return t

HEAD_RE = r'^([一二三四五六七八九十]、|第\d+节|第[一二三四五六七八九十]+节)'
NUM_RE = r'^(\d+\s*[\.．、]|（\d+）|\(\d+\)|[①②③④⑤⑥⑦⑧⑨⑩])'

def rebuild_cell(c, intent=False):
    color = INTENT_FG if intent else C.TEXT
    cap = cap_emu(c)
    seq = []
    for ch in [x for x in c if x.tag in (q('p'), q('tbl'))]:
        if ch.tag == q('tbl') or (ch.find('.//w:pict', ns) is not None and ch.find(f'.//{{{MCNS}}}AlternateContent') is None):
            seq.append(('keep', ch)); continue
        # 兼容包装（图形组）：取 Choice 里的 drawing，丢弃 VML 回退
        for ac in list(ch.iter(f'{{{MCNS}}}AlternateContent')):
            dr_ = ac.find(f'{{{MCNS}}}Choice/w:drawing', ns)
            par = ac.getparent()
            if dr_ is not None: par.replace(ac, dr_)
            else: par.remove(ac)
        drs = list(ch.iter(q('drawing')))
        t = clean(ptext(ch))
        if t: seq.append(('text', t))
        if drs:
            for dr in drs: dr.getparent().remove(dr)
            seq.append(('imgs', drs))
    # 被图片截断的句子接回：text(无句读结尾) + imgs + text → text + imgs
    i = 0
    while i + 2 < len(seq):
        a, b, c_ = seq[i], seq[i + 1], seq[i + 2]
        if (a[0] == 'text' and b[0] == 'imgs' and c_[0] == 'text' and not re.search(r'[。；：:？！）)]$', a[1])
                and not re.match(HEAD_RE + '|' + NUM_RE, c_[1])):
            seq[i] = ('text', a[1] + c_[1]); seq[i + 2] = b; del seq[i + 1]
        else:
            i += 1
    # 硬换行合并
    merged = []
    for kind, v in seq:
        if (kind == 'text' and merged and merged[-1][0] == 'text'
                and not re.search(r'[。；：:？?！!）)】]$', merged[-1][1])
                and not re.match(HEAD_RE + '|' + NUM_RE + '|【|任务\d|第[一二三四五]步|例', v)
                and not re.match(HEAD_RE + '|' + NUM_RE + '|【|任务\d', merged[-1][1])
                and len(merged[-1][1]) > 14):
            merged[-1] = ('text', merged[-1][1] + v)
        else:
            merged.append((kind, v))
    out = []
    for kind, v in merged:
        if kind == 'keep': out.append(v); continue
        if kind == 'imgs': out += img_paras(v, cap); continue
        # 一段内连写的（n）条目拆开
        segs = re.split(r'\s+(?=（\d+）)', v)
        for s in segs:
            s = s.strip()
            if not s: continue
            if re.match(HEAD_RE, s):
                out.append(el(P_h2(s))); continue
            if re.fullmatch(r'【[^】]{2,12}】', s) or re.fullmatch(r'任务\s*\d+', s):
                out.append(el(P_h3('', s.strip('【】')))); continue
            m = re.match(r'^(\d+)\s*[\.．、]\s*(.+)$', s)
            if m and len(m.group(2)) <= 12 and not re.search(r'[，。；：]', m.group(2)):
                out.append(el(P_h3(m.group(1), m.group(2)))); continue
            # 段内 “1.…；2.…”
            if re.search(r'[；;。]\s*\d+\s*[\.．、]', s) and re.match(r'^\d+\s*[\.．、]', s):
                nums = re.findall(r'(?:^|(?<=；)|(?<=;)|(?<=。))\s*(\d+)\s*[\.．、]', s)
                for n_, it in zip(nums, split_inline_numbers(s)):
                    out.append(el(P_item(n_, it, color)))
                continue
            m2 = re.match(r'^(\d+)\s*[\.．、]\s*(.+)$', s)
            m3 = re.match(r'^（(\d+)）\s*(.+)$', s) or re.match(r'^\((\d+)\)\s*(.+)$', s)
            m4 = re.match(r'^([①②③④⑤⑥⑦⑧⑨⑩])\s*(.+)$', s)
            if m2: out.append(el(P_item(m2.group(1), m2.group(2), color)))
            elif m3: out.append(el(P_item(f'({m3.group(1)})', m3.group(2), color)))
            elif m4: out.append(el(P_item(m4.group(1), m4.group(2), color)))
            else: out.append(el(P_body(s, color)))
    for ch in [x for x in c if x.tag in (q('p'), q('tbl'))]: c.remove(ch)
    for n_ in out:
        ppr = n_.find('w:pPr', ns) if n_.tag == q('p') else None
        if ppr is not None and ppr.find('w:numPr', ns) is not None and n_ not in [v for k, v in merged if k == 'keep']:
            ppr.remove(ppr.find('w:numPr', ns))
        c.append(n_)
    if not len(c) or c[-1].tag != q('p'):
        c.append(el(para([run('', 2)], line=20)))

HEAD_ICONS = {'教学环节': 'list-details-1F3864', '教学内容': 'book-2-1F3864', '教师活动': 'chalkboard-1F3864',
              '学生活动': 'users-1F3864', '设计意图': 'bulb-1F3864'}

def parse_step_name(c):
    raw = '\n'.join(lines(c))
    flat = re.sub(r'\s+', '', raw)
    tags = re.findall(r'【([^】]+)】', flat)
    flat2 = re.sub(r'【[^】]+】', '', flat)
    m = re.match(r'^(.*?)[（(]([^）)]*)[）)](.*)$', flat2)
    if m:
        name, inner, rest = m.group(1), m.group(2), m.group(3)
        bits = re.split(r'[·・]', inner)
        time = [b for b in bits if re.search(r'\d', b)]
        tags += [b for b in bits if not re.search(r'\d', b) and b]
        name = name + rest
    else:
        name, time = flat2, []
    return name, '·'.join(time), [t for t in tags if t]

def label_step(c, num, name, time, tags):
    set_tcpr(c, fill=C.CARD, bdr_xml=ALL_THIN, valign='center', mar={'top': 80, 'left': 60, 'bottom': 80, 'right': 60})
    xml = para([run(num, 26, C.GOLD_D, b=True)], jc='center', line=260, after=40) if num else ''
    if len(name) > 6 and re.search(r'[·・]', name): parts = re.split(r'(?<=[·・])', name)
    elif len(name) > 6: parts = [name[i:i + 4] for i in range(0, len(name), 4)]
    else: parts = [name]
    xml += ''.join(para([run(x, 21, C.NAVY, b=True)], jc='center', line=280) for x in parts)
    if time: xml += para([run(time, 17, C.SUB)], jc='center', line=240, before=40)
    if tags:
        rs = []
        for k, t in enumerate(tags):
            if k: rs.append(run(' ', 17))
            rs.append(chip(t.replace('·', '·'), 17))
        xml += para(rs, jc='center', line=280, before=60)
    set_paras(c, xml)

step_no = 0
in_mid = False
impl_rows = list(range(r_impl + 1, r_end))
phase_rows, head_rows, step_rows = [], [], []
for ri in impl_rows:
    r = rows[ri]; cs = cells(r)
    txt = ntext(r)
    if len(cs) == 1:
        phase_rows.append(ri)
        in_mid = '课中' in txt; step_no = 0
        m = re.match(r'(第[一二三四五]阶段)(.*)', ntext(cs[0]))
        stage, name = (m.group(1), m.group(2)) if m else ('', ntext(cs[0]))
        set_tcpr(cs[0], fill='FFFFFF', valign='center', mar={'top': 60, 'left': 200, 'bottom': 60, 'right': 200},
                 bdr_xml=borders('tcBorders', {'top': THIN, 'left': GOLD_BAR, 'bottom': ('single', 8, C.NAVY), 'right': THIN}))
        set_paras(cs[0], para([run(stage, 22, C.GOLD_D, b=True, spacing=20), run('　|　', 22, 'B4C3DA'),
                               run(name, 24, C.NAVY, b=True, spacing=20)], keep=True))
        trh(r, 500)
        continue
    if txt.startswith('教学环节'):
        head_rows.append(ri)
        for c in cs:
            nm = ntext(c)
            set_tcpr(c, fill='F3EEE2' if nm == '设计意图' else C.LABEL, bdr_xml=ALL_THIN, valign='center')
            set_paras(c, para([d.icon(HEAD_ICONS.get(nm, 'list-details-1F3864'), 12, lower=3),
                               run(' ' + nm, 21, C.NAVY, b=True)], jc='center', keep=True))
        continue
    step_rows.append(ri)
    step_no += 1
    name, time, tags = parse_step_name(cs[0])
    label_step(cs[0], f'{step_no:02d}' if in_mid else '', name, time, tags)
    for k, c in enumerate(cs[1:], 1):
        intent = (k == len(cs) - 1)
        set_tcpr(c, fill=INTENT_BG if intent else 'FFFFFF', bdr_xml=ALL_THIN, valign='top',
                 mar={'top': 80, 'left': 110, 'bottom': 80, 'right': 110})
        rebuild_cell(c, intent=intent)

# 知识框架图（总结评价·教学内容列末尾）
if CFG.get('framework_img'):
    sr = next(ri for ri in step_rows if '总结' in ntext(cells(rows[ri])[0]))
    c = cells(rows[sr])[1]
    for n_ in els(fig_xml(CFG['framework_img'], 'fig_kuangjia.png', int(cap_emu(c) * 0.98))):
        c.insert(len(c) - 1 if c[-1].tag == q('p') and not ptext(c[-1]).strip() else len(c), n_)


# ================= 11. 四、教学反思 =================
REFLECT = {'授课实效': 'chart-bar-1F3864', '特色亮点': 'sparkles-1F3864', '不足改进': 'tool-1F3864'}
refl_rows = []
for ri in range(r_end + 1, len(rows)):
    cs = cells(rows[ri])
    nm = ntext(cs[0])
    single_label(cs[0], REFLECT.get(nm, 'tool-1F3864'), nm)
    for c in cs[1:]:
        ls = lines(c)
        set_tcpr(c, fill='FFFFFF', bdr_xml=ALL_THIN, valign='top', mar=CONTENT_MAR)
        set_paras(c, ''.join(para(rich(x, 20), line=300) for x in ls) if ls else para([run('（课后填写）', 20, C.HINT)], line=300))
    trh(rows[ri], int(os.environ.get('REFLECT_H', '1400')))
    refl_rows.append(ri)


# ================= 12. 分页控制、左边框统一 =================
# 原模板的行高是按 14 号字设的，字号缩小后会留出大片空白：内容行最小行高统一放低（排满由 fill 控制）
for i, r in enumerate(rows):
    if i in sec_rows or i in refl_rows: continue
    trpr = r.find('w:trPr', ns)
    h = trpr.find('w:trHeight', ns) if trpr is not None else None
    if h is not None and int(h.get(q('val'), '0')) > 560:
        h.set(q('val'), '400'); h.set(q('hRule'), 'atLeast')
for ri in sec_rows + phase_rows + head_rows:
    keep_next(rows[ri])
WPNS = 'http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing'
def est_h(r):
    """行高粗估（twip）：取最高的单元格；文字按每字 200twip 宽、行距 300 估算"""
    best = 0
    for tc_ in r.findall('w:tc', ns):
        w_ = max(600, int(tc_.find('w:tcPr/w:tcW', ns).get(q('w'))) - 240)
        h_ = 200
        for p_ in tc_.iter(q('p')):
            drs = list(p_.iter(q('drawing')))
            if drs:
                h_ += max(int(e.get('cy')) for dr in drs for e in dr.iter(f'{{{WPNS}}}extent')) / 635 + 140
            else:
                n_ = len(ptext(p_).strip())
                h_ += max(1, -(-n_ * 200 // w_)) * 300 + 40
        best = max(best, h_)
    return best
PAGE_H = 16838 - 2 * 454
for i, r in enumerate(rows):
    # 能在一页内放下的行一律不跨页（LibreOffice 拆分含图行时会丢内容）；确实超过一页的大行才允许跨页
    big = i in step_rows and est_h(r) > 0.45 * PAGE_H
    trpr = r.find('w:trPr', ns)
    cs_ = trpr.find('w:cantSplit', ns) if trpr is not None else None
    if big:
        if cs_ is not None: trpr.remove(cs_)
    else:
        cant_split(r)
unify_left_edge(rows)

# 页面留白均摊（由 fill 程序给出）
plan = json.loads(os.environ.get('FILL_PLAN', '[]'))
def pad_rows(a, b, E):
    for r in rows[a:b + 1]:
        for c in cells(r):
            pr = c.find('w:tcPr', ns)
            vm = pr.find('w:vMerge', ns)
            if vm is not None and vm.get(q('val')) != 'restart': continue
            mar = pr.find('w:tcMar', ns)
            if mar is None:
                set_tcpr(c, mar={'top': E, 'left': 108, 'bottom': E, 'right': 108}); continue
            for side in ('top', 'bottom'):
                m = mar.find(f'w:{side}', ns)
                if m is None:
                    cur = {etree.QName(x).localname: int(x.get(q('w'))) for x in mar}
                    cur[side] = 0
                    cur = {k: cur.get(k, 108) for k in ('top', 'left', 'bottom', 'right')}
                    cur[side] = E
                    set_tcpr(c, mar=cur); mar = c.find('w:tcPr/w:tcMar', ns); continue
                m.set(q('w'), str(int(m.get(q('w'))) + E))
for a, b, E in plan:
    if E: pad_rows(a, b, E)

# 表格后多余空段只留一个
after = [p for p in body.findall('w:p', ns) if body.index(p) > body.index(tbl)]
for p in after[1:]:
    if not ptext(p).strip() and p.find('.//w:drawing', ns) is None and p.find('.//w:sectPr', ns) is None:
        body.remove(p)


# ================= 13. 页边距收窄、表格外扩 =================
SIDE, TB = CFG.get('margin_side', 284), CFG.get('margin_tb', 454)
PW = 11906
m = body.find('w:sectPr/w:pgMar', ns)
for k, v in (('top', TB), ('bottom', TB), ('left', SIDE), ('right', SIDE), ('header', 227), ('footer', 227)):
    m.set(q(k), str(v))
NEW = PW - 2 * SIDE
f = NEW / TW
for t in d.tree.iter(q('tbl')):
    tw = t.find('w:tblPr/w:tblW', ns)
    total = round(int(tw.get(q('w'))) * f) if tw is not None and tw.get(q('type')) == 'dxa' else None
    if total: tw.set(q('w'), str(total))
    gc = t.find('w:tblGrid', ns).findall('w:gridCol', ns)
    vals = [round(int(x.get(q('w'))) * f) for x in gc]
    if total: vals[-1] += total - sum(vals)
    for x, v in zip(gc, vals): x.set(q('w'), str(v))
    for tr in t.findall('w:tr', ns):
        i0 = 0
        for c in tr.findall('w:tc', ns):
            pr = c.find('w:tcPr', ns); gs = pr.find('w:gridSpan', ns)
            n = int(gs.get(q('val'))) if gs is not None else 1
            pr.find('w:tcW', ns).set(q('w'), str(sum(vals[i0:i0 + n]))); i0 += n

if os.environ.get('ROW_MARK') == '1':           # 仅供 fill 测量：每行首段末尾加 1 磅白色定位码
    for i_, r_ in enumerate(rows):
        p_ = r_.find('w:tc', ns).find('.//w:p', ns)
        p_.append(el(f'<w:r><w:rPr><w:color w:val="FFFFFF"/><w:sz w:val="2"/><w:szCs w:val="2"/></w:rPr><w:t>QZX{i_:03d}QZX</w:t></w:r>'))
d.save()
json.dump({'sec': sec_rows, 'phase': phase_rows, 'head': head_rows, 'step': step_rows, 'refl': refl_rows,
           'impl': [r_impl, r_end], 'n': len(rows)}, open(DST + '.rows.json', 'w'))
print('ok rows', len(rows), 'steps', step_rows)
