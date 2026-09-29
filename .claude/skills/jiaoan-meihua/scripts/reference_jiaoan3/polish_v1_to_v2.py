"""在 教案3v1 基础上美化“03 教学实施”，并统一表格最左侧边框。只改格式与明显笔误，不动行列结构。
用法：python3 polish.py <unpacked_dir>
"""
import sys, re
head = open('/tmp/lo1/v2.py', encoding='utf8').read().split('# ---------- load ----------')[0]
exec(head)          # D、rels、run、para、tc、borders、icon、配色等

tree = etree.fromstring(docxml.encode('utf8'))
body = tree.find('w:body', ns)
tbl = max(body.findall('w:tbl', ns), key=lambda t: len(t.findall('w:tr', ns)))
rows = tbl.findall('w:tr', ns)
SIDES = ('top', 'left', 'bottom', 'right')
GOLD_BAR = ('single', 24, GOLD)
INTENT_BG, INTENT_FG = 'FBF8F1', '5C5346'
PORD = ['pStyle', 'keepNext', 'keepLines', 'pageBreakBefore', 'framePr', 'widowControl', 'numPr',
        'suppressLineNumbers', 'pBdr', 'shd', 'tabs', 'suppressAutoHyphens', 'kinsoku', 'wordWrap',
        'overflowPunct', 'topLinePunct', 'autoSpaceDE', 'autoSpaceDN', 'bidi', 'adjustRightInd', 'snapToGrid',
        'spacing', 'ind', 'contextualSpacing', 'mirrorIndents', 'suppressOverlap', 'jc', 'textDirection',
        'textAlignment', 'textboxTightWrap', 'outlineLvl', 'divId', 'cnfStyle', 'rPr', 'sectPr', 'pPrChange']
TCPR = ['cnfStyle', 'tcW', 'gridSpan', 'hMerge', 'vMerge', 'tcBorders', 'shd', 'noWrap', 'tcMar',
        'textDirection', 'tcFitText', 'vAlign', 'hideMark']

def put(parent, name, order, node):
    old = parent.find(f'w:{name}', ns)
    if old is not None: parent.remove(old)
    idx = order.index(name); pos = len(parent)
    for i, ch in enumerate(parent):
        ln = etree.QName(ch).localname
        if ln in order and order.index(ln) > idx: pos = i; break
    parent.insert(pos, node)

def set_fill(cell, fill):
    put(cell.find('w:tcPr', ns), 'shd', TCPR, el(f'<w:shd w:val="clear" w:color="auto" w:fill="{fill}"/>'))

def set_side(cell, side, spec):
    pr = cell.find('w:tcPr', ns)
    b = pr.find('w:tcBorders', ns)
    if b is None:
        b = el('<w:tcBorders/>'); put(pr, 'tcBorders', TCPR, b)
    old = b.find(f'w:{side}', ns)
    if old is not None: b.remove(old)
    node = el(f'<w:{side} w:val="{spec[0]}" w:sz="{spec[1]}" w:space="0" w:color="{spec[2]}"/>')
    order = ['top', 'start', 'left', 'bottom', 'end', 'right', 'insideH', 'insideV', 'tl2br', 'tr2bl']
    pos = len(b)
    for i, ch in enumerate(b):
        if order.index(etree.QName(ch).localname) > order.index(side): pos = i; break
    b.insert(pos, node)

def ptext(p):
    return ''.join(x.text or '' for x in p.iter(q('t')))

def has_img(p):
    return p.find('.//w:drawing', ns) is not None

def is_caption(p):
    sz = p.find('.//w:r/w:rPr/w:sz', ns)
    col = p.find('.//w:r/w:rPr/w:color', ns)
    return (sz is not None and sz.get(q('val')) == '16' and col is not None and col.get(q('val')) == '666666')

# ================= 1. 表格最左侧边框统一为金色细条 =================
for r in rows:
    first = r.find('w:tc', ns)
    set_side(first, 'left', GOLD_BAR)

# ================= 2. 03 教学实施 =================
def idx_of(prefix):
    for i, r in enumerate(rows):
        if re.sub(r'\s+', '', ''.join(r.itertext())).startswith(prefix): return i
    raise KeyError(prefix)

r_start = idx_of('03教学实施')
r_end = idx_of('04教学反思')
impl = rows[r_start + 1:r_end]

BODY_SZ, H_SZ = 19, 21
def body_runs(text, color=TEXT):
    return [run(text, BODY_SZ, color)]

def P_body(text, color=TEXT):
    return para(body_runs(text, color), jc='left', line=290, after=40)

def P_item(num, text, color=TEXT):
    w = 300 if len(num) <= 2 else 420
    return para([run(num + '\t', BODY_SZ, GOLD_D, b=True)] + body_runs(text, color), jc='left', line=290, after=40,
                ind_left=w, hanging=w)

def P_h2(text):
    return para([run('▍', 20, GOLD), run(text, H_SZ, NAVY, b=True)], line=260, before=120, after=60)

def P_h3(num, text):
    return para([run(num + ' ', H_SZ - 1, GOLD_D, b=True), run(text, H_SZ - 1, NAVY, b=True)], line=260, before=60, after=30)

TYPO = [('府视图', '俯视图'), ('提问学生一下问题', '提问学生以下问题'), ('学生 学以致用', '学生学以致用'),
        ('分析物体确定主视图投射方向图11所示。', '分析物体，确定主视图投射方向。'),
        ('推导 “', '推导“'), ('” 投影', '”投影'), ('动画 + 实物', '动画+实物'), ('投影面上图2-2', '投影面上（图2-2）')]

def clean(t):
    t = t.replace('　', ' ').strip()
    t = re.sub(r'\s{2,}', ' ', t)
    for a, b in TYPO: t = t.replace(a, b)
    return t

def split_inline_numbers(t):
    """“1. A；2. B；3. C” → [('1', A), ('2', B), ('3', C)]"""
    parts = re.split(r'(?:^|(?<=；))\s*(\d+)[\.．、]\s*', t)
    out = []
    if parts and parts[0].strip():
        out.append((None, parts[0].strip()))
    for i in range(1, len(parts) - 1, 2):
        out.append((parts[i], parts[i + 1].strip()))
    return out

def rebuild_cell(cell, intent=False, auto_seq=None):
    """把单元格段落改写为统一样式；保留图片段与图注段原样（仅规范间距）"""
    color = INTENT_FG if intent else TEXT
    items = []            # (kind, payload)
    old = [ch for ch in cell if ch.tag in (q('p'), q('tbl'))]
    for p in old:
        if p.tag == q('tbl') or has_img(p) or is_caption(p):
            items.append(('keep', p)); continue
        t = clean(ptext(p))
        if not t: continue
        numbered = p.find('w:pPr/w:numPr', ns) is not None
        if auto_seq and numbered and auto_seq.get(t[:6]):
            t = auto_seq[t[:6]] + t
        items.append(('text', t))
    # 合并被硬换行拆开的句子（上一段不以句读结尾且下一段不是标题/编号）
    merged = []
    for kind, v in items:
        if (kind == 'text' and merged and merged[-1][0] == 'text'
                and not re.search(r'[。；：:？?！!）)]$', merged[-1][1])
                and not re.match(r'^(\d+[\.．、]|（\d+）|[①②③④⑤]|[一二三四五六七八九十]、|第\d+节)', merged[-1][1])
                and not re.match(r'^([一二三四五六七八九十]、|第\d+节|\d+[\.．、]|（\d+）|[①②③④⑤])', v)
                and len(merged[-1][1]) > 12):
            merged[-1] = ('text', merged[-1][1] + v)
        else:
            merged.append((kind, v))
    xml_nodes = []
    for kind, v in merged:
        if kind == 'keep':
            xml_nodes.append(v); continue
        t = v
        if re.match(r'^([一二三四五六七八九十]、|第\d+节)', t):
            xml_nodes.append(P_h2(t.replace('  ', ' ')))
            continue
        m = re.match(r'^(\d+)[\.．、]\s*(.+)$', t)
        if m and len(m.group(2)) <= 10 and not re.search(r'[，。；]', m.group(2)):
            xml_nodes.append(P_h3(m.group(1), m.group(2)))
            continue
        segs = split_inline_numbers(t) if re.search(r'；\s*\d+[\.．]', t) or re.match(r'^\d+[\.．、]', t) else [(None, t)]
        for num, s in segs:
            m2 = re.match(r'^（(\d+)）\s*(.+)$', s)
            m3 = re.match(r'^([①②③④⑤])\s*(.+)$', s)
            if num:
                xml_nodes.append(P_item(num, s, color))
            elif m2:
                xml_nodes.append(P_item(f'({m2.group(1)})', m2.group(2), color))
            elif m3:
                xml_nodes.append(P_item(m3.group(1), m3.group(2), color))
            elif re.match(r'^作图步骤', s):
                xml_nodes.append(P_h3('', s))
            else:
                xml_nodes.append(P_body(s, color))
    for p in old:
        cell.remove(p)
    for n in xml_nodes:
        if isinstance(n, str):
            cell.append(el(n))
        else:
            ppr = n.find('w:pPr', ns)
            if ppr is not None:
                num = ppr.find('w:numPr', ns)
                if num is not None: ppr.remove(num)
            cell.append(n)
    if not cell.findall('w:p', ns):
        cell.append(el(para([run('', BODY_SZ)])))

# 自动编号原来显示的序号（按渲染结果核对）：以段首 6 字为键补成文字编号
AUTO = {
    '投影法的基本': '1. ', '生活中影子现': '2. ', '了解什么是正': '3. ',
    '投影法分类：': '1. ',
    '积聚性': '2．', '类似性': '3．',
    '三投影面体系': '1. ',
    '投影面展开，': '3. ',
    '分析物体，确': '（1）', '画直角弯板的': '（2）', '画方槽的三面': '（3）', '画右边切角的': '（4）',
}
# 教师活动首段原文编号错位（1、2、2），这里按 1–4 重排：先把两段拆成 4 项
def fix_teacher_first(cell):
    ps = cell.findall('w:p', ns)
    for p in ps[:3]:
        t = ptext(p)
        if t.startswith('三投影面体系构成'):
            for r_ in p.findall('.//w:t', ns):
                r_.text = r_.text.replace('三投影面体系构成；2. 物体向三个投影面投影；', '1. 三投影面体系构成；2. 物体向三个投影面投影；')
        if t.startswith('投影面展开'):
            for r_ in p.findall('.//w:t', ns):
                r_.text = r_.text.replace('投影面展开，得到', '3. 投影面展开，得到')
            nm = p.find('w:pPr/w:numPr', ns)
            if nm is not None: nm.getparent().remove(nm)
        nm = p.find('w:pPr/w:numPr', ns)
        if t.startswith('三投影面体系构成') and nm is not None: nm.getparent().remove(nm)

def label_step(cell, num, name):
    set_fill(cell, CARD)
    lines = [name[i:i + 4] for i in range(0, len(name), 4)] if len(name) > 4 else [name]
    xml = (para([run(num, 26, GOLD_D, b=True)], jc='center', line=260, after=40) if num else '') + \
          ''.join(para([run(x, 21, NAVY, b=True)], jc='center', line=280) for x in lines)
    for p in cell.findall('w:p', ns): cell.remove(p)
    for n in etree.fromstring(f'<x {NSDECL}>{xml}</x>'): cell.append(n)

HEAD_ICONS = {'教学环节': 'list-details-1F3864', '教学内容': 'book-2-1F3864', '教师活动': 'chalkboard-1F3864',
              '学生活动': 'users-1F3864', '设计意图': 'bulb-1F3864'}
step_no = 0
in_mid = False
for r in impl:
    cs = r.findall('w:tc', ns)
    txt = re.sub(r'\s+', '', ''.join(r.itertext()))
    if len(cs) == 1:                                   # 阶段行
        in_mid = '课中' in txt
        step_no = 0
        continue
    if txt.startswith('教学环节'):                     # 表头行
        for c in cs:
            name = re.sub(r'\s+', '', ''.join(c.itertext()))
            for p in c.findall('w:p', ns): c.remove(p)
            c.append(el(para([icon(HEAD_ICONS.get(name, 'list-details-1F3864'), 12, lower=3),
                              run(' ' + name, 21, NAVY, b=True)], jc='center', keep=True)))
            set_fill(c, LABEL)
            if name == '设计意图': set_fill(c, 'F3EEE2')
        continue
    # 环节行
    name = re.sub(r'\s+', '', ''.join(cs[0].itertext()))
    step_no += 1
    label_step(cs[0], f'{step_no:02d}' if in_mid else '', name)
    for k, c in enumerate(cs[1:], 1):
        if k == 2 and name.startswith('示范讲解'): fix_teacher_first(c)
        intent = (k == len(cs) - 1)
        rebuild_cell(c, intent=intent, auto_seq=AUTO)
        if intent: set_fill(c, INTENT_BG)

open(doc_path, 'wb').write(etree.tostring(tree, xml_declaration=True, encoding='UTF-8', standalone=True))
open(rels_path, 'w', encoding='utf8').write(rels)
print('rows', r_start, r_end, 'icons', len(icon_rid))
