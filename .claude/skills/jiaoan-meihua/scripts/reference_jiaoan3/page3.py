# 第 3 页起（教学资源之后）：考核方式 / 教学流程 / 三、教学实施 / 四、教学反思
# 只改格式：行列、合并、嵌套表、图片、自动编号全部保留；由 v2.py 在 page2_v2.py 之后 exec

RORD = ['rStyle', 'rFonts', 'b', 'bCs', 'i', 'iCs', 'caps', 'smallCaps', 'strike', 'dstrike', 'outline',
        'shadow', 'emboss', 'imprint', 'noProof', 'snapToGrid', 'vanish', 'webHidden', 'color', 'spacing',
        'w', 'kern', 'position', 'sz', 'szCs', 'highlight', 'u', 'effect', 'bdr', 'shd', 'fitText',
        'vertAlign', 'rtl', 'cs', 'em', 'lang', 'eastAsianLayout', 'specVanish', 'oMath']
PORD = ['pStyle', 'keepNext', 'keepLines', 'pageBreakBefore', 'framePr', 'widowControl', 'numPr',
        'suppressLineNumbers', 'pBdr', 'shd', 'tabs', 'suppressAutoHyphens', 'kinsoku', 'wordWrap',
        'overflowPunct', 'topLinePunct', 'autoSpaceDE', 'autoSpaceDN', 'bidi', 'adjustRightInd', 'snapToGrid',
        'spacing', 'ind', 'contextualSpacing', 'mirrorIndents', 'suppressOverlap', 'jc', 'textDirection',
        'textAlignment', 'textboxTightWrap', 'outlineLvl', 'divId', 'cnfStyle', 'rPr', 'sectPr', 'pPrChange']

def put(parent, name, order, attrs=None):
    old = parent.find(f'w:{name}', ns)
    if old is not None: parent.remove(old)
    node = etree.Element(q(name))
    for k, v in (attrs or {}).items(): node.set(q(k), str(v))
    idx = order.index(name); pos = len(parent)
    for i, ch in enumerate(parent):
        ln = etree.QName(ch).localname
        if ln in order and order.index(ln) > idx: pos = i; break
    parent.insert(pos, node)
    return node

def ensure(parent, name, first=True):
    node = parent.find(f'w:{name}', ns)
    if node is None:
        node = etree.Element(q(name))
        parent.insert(0, node) if first else parent.append(node)
    return node

def style_rpr(rpr, sz, color, bold=None):
    put(rpr, 'rFonts', RORD, {'ascii': HEI, 'hAnsi': HEI, 'eastAsia': HEI, 'cs': HEI})
    col = rpr.find('w:color', ns)
    keep = col is not None and col.get(q('val')) not in (None, 'auto', '000000') and color is None
    if not keep:
        put(rpr, 'color', RORD, {'val': color or TEXT})
    if bold is not None:
        for b in ('b', 'bCs'):
            old = rpr.find(f'w:{b}', ns)
            if old is not None: rpr.remove(old)
        if bold:
            put(rpr, 'b', RORD); put(rpr, 'bCs', RORD)
    put(rpr, 'sz', RORD, {'val': sz}); put(rpr, 'szCs', RORD, {'val': sz})

def restyle(node, sz=20, color=None, bold=None, jc='left', line=None, after=40):
    """就地改写 node（单元格或嵌套表）内所有段落与文字的格式"""
    line = line or round(sz * 10 * 1.5)
    for p in node.iter(q('p')):
        ppr = ensure(p, 'pPr')
        has_img = p.find('.//w:drawing', ns) is not None or p.find('.//w:pict', ns) is not None
        put(ppr, 'snapToGrid', PORD, {'val': 0})
        if has_img:
            put(ppr, 'spacing', PORD, {'before': 40, 'after': 40, 'line': 240, 'lineRule': 'auto'})
        else:
            put(ppr, 'spacing', PORD, {'before': 0, 'after': after, 'line': line, 'lineRule': 'exact'})
        if ppr.find('w:numPr', ns) is None:
            put(ppr, 'ind', PORD, {'left': 0, 'firstLine': 0})
        old = ppr.find('w:jc', ns)
        if jc or (old is not None and old.get(q('val')) in ('both', 'distribute')):
            put(ppr, 'jc', PORD, {'val': jc or 'left'})
        style_rpr(ensure(ppr, 'rPr', first=False), sz, color, bold)
        for r in p.findall('.//w:r', ns):
            style_rpr(ensure(r, 'rPr'), sz, color, bold)

def cells(r):
    return r.findall('w:tc', ns)

def label_cell(cell, icon_name, text):
    single_label(cell, icon_name, text)

def plain_content(cell, **kw):
    set_tcpr(cell, fill='FFFFFF', bdr_xml=ALL_THIN, valign='top', mar=CONTENT_MAR)
    restyle(cell, **kw)

# —— 考核方式（18 行）：原嵌套评价表拆成主表的 4 行（表头 + 课前/课中/课后），列与主表网格对齐 ——
# 主表网格：1506 | 2429 | 70 | 1875 | 795 | 840 | 1800 | 1799
KH_COLS = [(2429, 1), (70 + 1875, 2), (795 + 840, 2), (1800 + 1799, 2)]   # 时段 / 内容 / 主体 / 方式
old_row = rows[18]
inner = cells(old_row)[1].find('w:tbl', ns)
data = [[[''.join(p.itertext()).strip() for p in tc.findall('w:p', ns)] for tc in r.findall('w:tc', ns)]
        for r in inner.findall('w:tr', ns)]
data = [[[x for x in c if x] for c in r] for r in data]

KH_MAR = {'top': 80, 'left': 140, 'bottom': 80, 'right': 140}
def kh_lines(lines, color=TEXT, bold=False, jc='left'):
    return ''.join(para([run(x, 20, color, b=bold)], jc=jc, line=280, after=30) for x in lines) or para([run('', 20)])

kh_rows = []
for k, r in enumerate(data):
    if k == 0:
        first = tc(1506, para([icon('clipboard-check-1F3864', 14, lower=4), run(' 考核方式', 22, NAVY, b=True)], jc='center'),
                   fill=LABEL, vmerge='restart', bdr=LABEL_B)
    else:
        first = tc(1506, para([run('', 20)]), fill=LABEL, vmerge='continue', bdr=LABEL_CONT_B)
    xml = first
    for j, ((w, span), lines) in enumerate(zip(KH_COLS, r)):
        if k == 0:
            xml += tc(w, kh_lines(lines, NAVY, True, 'center'), fill=CARD, span=span, bdr=ALL_THIN, mar=KH_MAR)
        elif j == 0:
            xml += tc(w, kh_lines(lines, NAVY, True, 'center'), fill='FFFFFF', span=span, bdr=ALL_THIN, mar=KH_MAR)
        else:
            xml += tc(w, kh_lines(lines, jc='center' if j == 2 else 'left'), fill='FFFFFF', span=span,
                      bdr=ALL_THIN, mar=KH_MAR)
    height = 440 if k == 0 else 700
    kh_rows.append(el(f'<w:tr><w:trPr><w:cantSplit/><w:trHeight w:val="{height}" w:hRule="atLeast"/>'
                      f'<w:jc w:val="center"/></w:trPr>{xml}</w:tr>'))
for nr in reversed(kh_rows):
    old_row.addnext(nr)
tbl.remove(old_row)

# —— 教学流程（19 行）：流程图居中 ——
c0, c1 = cells(rows[19])
label_cell(c0, 'route-1F3864', '教学流程')
set_tcpr(c1, fill='FFFFFF', bdr_xml=ALL_THIN, valign='center', mar={'top': 80, 'left': 80, 'bottom': 80, 'right': 80})
for p in c1.findall('w:p', ns):
    ppr = ensure(p, 'pPr'); put(ppr, 'jc', PORD, {'val': 'center'})
    put(ppr, 'snapToGrid', PORD, {'val': 0}); put(ppr, 'spacing', PORD, {'before': 0, 'after': 0, 'line': 240, 'lineRule': 'auto'})

# —— 三、教学实施 / 四、教学反思：分区标题栏 ——
section_row(rows[20], 'presentation-FFFFFF', '03', '教学实施')
section_row(rows[33], 'message-2-FFFFFF', '04', '教学反思')

# —— 阶段行：白底 + 左侧金色粗边 + 金色阶段名 + 藏蓝标题 ——
def phase_row(row):
    c = cells(row)[0]
    txt = ''.join(c.itertext()).strip()
    stage, _, name = txt.partition(' ')
    set_tcpr(c, fill='FFFFFF', valign='center', mar={'top': 60, 'left': 200, 'bottom': 60, 'right': 200},
             bdr_xml=borders('tcBorders', {'top': THIN, 'left': ('single', 36, GOLD), 'bottom': ('single', 8, NAVY), 'right': THIN}))
    set_paras(c, para([run(stage, 22, GOLD_D, b=True, spacing=20), run('　|　', 22, 'B4C3DA'),
                       run(name.strip(), 24, NAVY, b=True, spacing=20)], keep=True))
    trpr = row.find('w:trPr', ns)
    h = trpr.find('w:trHeight', ns)
    if h is not None: h.set(q('val'), '520')

def header_row(row):
    for c in cells(row):
        set_tcpr(c, fill=LABEL, bdr_xml=ALL_THIN, valign='center')
        txt = ''.join(c.itertext()).strip()
        set_paras(c, para([run(txt, 21, NAVY, b=True)], jc='center', keep=True))

def step_row(row):
    cs = cells(row)
    name = ''.join(cs[0].itertext()).strip()
    set_tcpr(cs[0], fill=CARD, bdr_xml=ALL_THIN, valign='center')
    # 环节名较长时两字一行
    lines = [name[i:i + 4] for i in range(0, len(name), 4)] if len(name) > 4 else [name]
    set_paras(cs[0], ''.join(para([run(x, 21, NAVY, b=True)], jc='center') for x in lines))
    for c in cs[1:]:
        set_tcpr(c, fill='FFFFFF', bdr_xml=ALL_THIN, valign='top', mar={'top': 80, 'left': 110, 'bottom': 80, 'right': 110})
        restyle(c, sz=19, line=290, after=30)
        # 单元格内嵌套表（示范讲解中的图表）：统一细线
        for t2 in c.findall('.//w:tbl', ns):
            for tc2 in t2.iter(q('tc')):
                set_tcpr(tc2, bdr_xml=ALL_THIN)

for ri in (21, 24, 30):
    phase_row(rows[ri])
for ri in (22, 25, 31):
    header_row(rows[ri])
for ri in (23, 26, 27, 28, 29, 32):
    step_row(rows[ri])

# 阶段行与表头不跨页、并与下一行同页；内容行允许跨页（示范讲解一行跨多页）
for ri in (20, 21, 22, 24, 25, 30, 31, 33):
    trpr = rows[ri].find('w:trPr', ns)
    if trpr.find('w:cantSplit', ns) is None:
        trpr.insert(0, etree.Element(q('cantSplit')))

# —— 教学反思（34-36 行）：留白待填 ——
for ri, ic in ((34, 'chart-bar-1F3864'), (35, 'sparkles-1F3864'), (36, 'tool-1F3864')):
    c0, c1 = cells(rows[ri])
    label_cell(c0, ic, ''.join(c0.itertext()).strip())
    set_tcpr(c1, fill='FFFFFF', bdr_xml=ALL_THIN, valign='top', mar=CONTENT_MAR)
    set_paras(c1, para([run('（课后填写）', 20, HINT)], line=300))
    h = rows[ri].find('w:trPr/w:trHeight', ns)
    if h is None:
        h = etree.SubElement(ensure(rows[ri], 'trPr'), q('trHeight'))
    h.set(q('val'), os.environ.get('REFLECT_H', '4200')); h.set(q('hRule'), 'atLeast')

# —— 表格后的空段：只保留一个，避免多出空白页 ——
tail_ps = [p for p in body.findall('w:p', ns) if body.index(p) > body.index(tbl)]
for p in tail_ps[1:]:
    if not ''.join(p.itertext()).strip() and p.find('.//w:drawing', ns) is None and p.find('.//w:sectPr', ns) is None:
        body.remove(p)

# 分区标题栏与后续行同页；教学反思四行整体不拆开
for ri in (20, 33, 34, 35):
    for p in rows[ri].iter(q('p')):
        put(ensure(p, 'pPr'), 'keepNext', PORD)
for ri in (34, 35, 36):
    trpr = rows[ri].find('w:trPr', ns)
    if trpr.find('w:cantSplit', ns) is None:
        trpr.insert(0, etree.Element(q('cantSplit')))

# 第 3 页（教学资源 / 考核方式 / 教学流程）均摊余量，使本页填满
E3 = int(os.environ.get('FILL_P3', '0'))
for r_ in [rows[14], rows[15], rows[16], rows[17], *kh_rows]:
    mar = r_.findall('w:tc', ns)[-1].find('w:tcPr/w:tcMar', ns)
    for side in ('top', 'bottom'):
        m = mar.find(f'w:{side}', ns)
        m.set(q('w'), str(int(m.get(q('w'))) + E3))
