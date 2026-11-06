# 第 2 页（蓝金 v2）：学情分析 / 教学目标 / 重点难点 / 方法 / 资源
# 只改单元格内的格式，不动行列与合并关系；由 v2.py 在保存前 exec

TCPR_ORDER = ['cnfStyle', 'tcW', 'gridSpan', 'hMerge', 'vMerge', 'tcBorders', 'shd', 'noWrap',
              'tcMar', 'textDirection', 'tcFitText', 'vAlign', 'hideMark']

def set_tcpr(cell, fill=None, bdr_xml=None, valign=None, mar=None):
    pr = cell.find('w:tcPr', ns)
    def put(name, xml):
        old = pr.find(f'w:{name}', ns)
        if old is not None: pr.remove(old)
        node = el(xml)
        idx = TCPR_ORDER.index(name); pos = len(pr)
        for i, ch in enumerate(pr):
            ln = etree.QName(ch).localname
            if ln in TCPR_ORDER and TCPR_ORDER.index(ln) > idx: pos = i; break
        pr.insert(pos, node)
    if bdr_xml: put('tcBorders', bdr_xml)
    if fill: put('shd', f'<w:shd w:val="clear" w:color="auto" w:fill="{fill}"/>')
    if mar: put('tcMar', '<w:tcMar>' + ''.join(f'<w:{k} w:w="{v}" w:type="dxa"/>' for k, v in mar.items()) + '</w:tcMar>')
    if valign: put('vAlign', f'<w:vAlign w:val="{valign}"/>')

def set_paras(cell, xml):
    for p in cell.findall('w:p', ns):
        cell.remove(p)
    for node in etree.fromstring(f'<x {NSDECL}>{xml}</x>'):
        cell.append(node)

def texts(cell):
    return [s for s in (''.join(p.itertext()).strip() for p in cell.findall('w:p', ns)) if s]

CONTENT_MAR = {'top': 90, 'left': 180, 'bottom': 90, 'right': 180}
BODY2 = dict(line=320, after=40)
LABEL_CONT_B = borders('tcBorders', {'top': NONE, 'left': GOLD_BAR, 'bottom': THIN, 'right': THIN})

def group_label(cell, icon_name, text, note=None):
    """大类标签（第二级）：浅蓝灰底 + 左侧金色细边 + 藏蓝图标与文字"""
    lines = [text[:2], text[2:]] if len(text) == 4 else [text]
    set_tcpr(cell, fill=LABEL, bdr_xml=LABEL_B, valign='center')
    top = para([run(note, 16, '8A99B3')], jc='center', after=200) if note else ''
    set_paras(cell, top + para([icon(icon_name, 22)], jc='center', after=60)
              + ''.join(para([run(x, 24, NAVY, b=True, spacing=40)], jc='center') for x in lines))

def group_continue(cell):
    set_tcpr(cell, fill=LABEL, bdr_xml=LABEL_CONT_B)

def single_label(cell, icon_name, text):
    set_tcpr(cell, fill=LABEL, bdr_xml=LABEL_B, valign='center')
    set_paras(cell, para([icon(icon_name, 14, lower=4), run(' ' + text, 22, NAVY, b=True)], jc='center'))

def sub_label(cell, text):
    """小类标签（第三级）：卡片底色，无图标"""
    set_tcpr(cell, fill=CARD, bdr_xml=ALL_THIN, valign='center')
    set_paras(cell, para([run(text, 22, NAVY, b=True)], jc='center'))

def rich(s, highlight=False):
    s = (s.replace('三视图 “', '三视图“').replace('” 投影', '”投影').replace('” 三等', '”三等')
          .replace('落实 “', '落实“'))
    out = []
    for part in re.split(r'(“长对正、高平齐、宽相等”|“宽相等”)', s):
        if not part: continue
        if highlight and part.startswith('“'):
            out.append(chip(part.strip('“”'), 20))
        else:
            out.append(run(part, 21, TEXT))
    return out

def bullet_list(items):
    return ''.join(para([run('■\t', 14, '8A99B3', pos=2)] + rich(x), jc='both', ind_left=240, hanging=240, **BODY2)
                   for x in items)

def numbered(items, highlight=False):
    out = ''
    for i, x in enumerate(items, 1):
        x = re.sub(r'^\d+\.\s*', '', x)
        out += para([run(f'{i}\t', 21, GOLD_D, b=True)] + rich(x, highlight), jc='both',
                    ind_left=300, hanging=300, **BODY2)
    return out

def content(cell, xml):
    set_tcpr(cell, fill='FFFFFF', bdr_xml=ALL_THIN, valign='center', mar=CONTENT_MAR)
    set_paras(cell, xml)

def split_semicolon(s):
    return [p.strip() for p in re.split(r'(?<=；)', s) if p.strip()]

# —— 学情分析（7-9 行），顶部加“续”提示 ——
for k, name in enumerate(['知识和技能基础', '认知和实践能力', '学习特点']):
    c0, c1, c2 = rows[7 + k].findall('w:tc', ns)
    if k == 0: group_label(c0, 'users-group-1F3864', '学情分析', note='02 教学设计 · 续')
    else: group_continue(c0)
    sub_label(c1, name)
    content(c2, bullet_list(split_semicolon(texts(c2)[0])))

# —— 教学目标（10-12 行） ——
for k, name in enumerate(['知识目标', '能力目标', '素质目标']):
    c0, c1, c2 = rows[10 + k].findall('w:tc', ns)
    if k == 0: group_label(c0, 'target-1F3864', '教学目标')
    else: group_continue(c0)
    sub_label(c1, name)
    content(c2, numbered(texts(c2)))

# —— 重点 / 难点：只在“教学重点”里高亮三等规律 ——
for ri, ic, hl in ((13, 'star-1F3864', True), (14, 'alert-triangle-1F3864', False)):
    c0, c1 = rows[ri].findall('w:tc', ns)
    single_label(c0, ic, texts(c0)[0])
    content(c1, numbered([x.strip() for x in texts(c1)], highlight=hl))

# —— 教学方法 / 学习方法：方法名做成浅蓝灰标签 ——
def split_dun(raw):
    items, buf, depth = [], '', 0
    for ch in raw:
        if ch in '（(': depth += 1
        if ch in '）)': depth -= 1
        if ch == '、' and depth == 0: items.append(buf); buf = ''
        else: buf += ch
    items.append(buf)
    return [x.strip() for x in items if x.strip()]

for ri, ic in ((15, 'bulb-1F3864'), (16, 'users-group-1F3864')):
    c0, c1 = rows[ri].findall('w:tc', ns)
    single_label(c0, ic, texts(c0)[0])
    rs = []
    for i, x in enumerate(split_dun(texts(c1)[0].rstrip('。'))):
        if i: rs.append(run('　', 21))
        rs.append(chip(x, 21, LABEL, NAVY))
    content(c1, para(rs, line=400))

# —— 教学资源：【标签】做成藏蓝小标签 ——
c0, c1 = rows[17].findall('w:tc', ns)
single_label(c0, 'books-1F3864', texts(c0)[0])
res = ''
for x in texts(c1):
    m = re.match(r'【(.+?)】(.*)', x)
    tag, rest = (m.group(1).replace(' ', ''), m.group(2)) if m else ('', x)
    res += para(([chip(tag, 19, NAVY, 'FFFFFF'), run('　', 21)] if tag else []) + rich(rest),
                jc='both', **BODY2)
content(c1, res)

# 第 2 页各行（学情分析 → 学习方法）加上均摊的留白，使本页填满
for r in rows[7:14]:
    mar = r.findall('w:tc', ns)[-1].find('w:tcPr/w:tcMar', ns)
    for side in ('top', 'bottom'):
        m = mar.find(f'w:{side}', ns)
        m.set(q('w'), str(int(m.get(q('w'))) + E2))

# 行不跨页：页面在行与行之间断开
for r in rows[6:18]:
    trpr = r.find('w:trPr', ns)
    if trpr.find('w:cantSplit', ns) is None:
        trpr.insert(0, etree.Element(q('cantSplit')))
