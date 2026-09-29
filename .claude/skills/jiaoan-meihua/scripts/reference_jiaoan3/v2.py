"""教案3 第1、2页（蓝金 v2）：藏蓝主色 + 金色强调 + 浅蓝灰标签 + 中性灰正文。
用法：python3 redesign.py <unpacked_dir>
"""
import sys, os, shutil, re
from lxml import etree

D = sys.argv[1]
W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
ns = {'w': W}
q = lambda n: f'{{{W}}}{n}'
NSDECL = (f'xmlns:w="{W}" xmlns:r="{R}" '
          'xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" '
          'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" '
          'xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture"')

# ---------- palette & fonts ----------
NAVY, GOLD, GOLD_D = '1F3864', 'C9A227', 'A8841A'      # 主色 / 强调色 / 深金（编号文字）
LABEL, CARD, LINE = 'EEF3F9', 'F7F9FC', 'D5DEEA'       # 标签底 / 卡片底 / 细线
TEXT, SUB, HINT = '333333', '666666', 'AAAAAA'         # 正文 / 说明 / 占位
CHIP_BG, CHIP_FG = 'F6EDD2', '7A5C00'                  # 金色关键词标签
RED, WARM = '9B1C1C', 'FAF6F0'                         # 课程思政：深红标题 / 暖灰底
BLUE = NAVY
HEI = '微软雅黑'

# ---------- images ----------
rels_path = f'{D}/word/_rels/document.xml.rels'
rels = open(rels_path, encoding='utf8').read()
rid_n = max(int(x) for x in re.findall(r'Id="rId(\d+)"', rels))
doc_path = f'{D}/word/document.xml'
docxml = open(doc_path, encoding='utf8').read()
pic_id = [max(int(x) for x in re.findall(r'docPr id="(\d+)"', docxml))]
icon_rid = {}

def icon_rel(name):
    global rid_n, rels
    if name not in icon_rid:
        rid_n += 1
        fn = f'p1icon_{name}.png'
        shutil.copy(f'/tmp/lo1/icons/{name}.png', f'{D}/word/media/{fn}')
        rels = rels.replace('</Relationships>',
            f'<Relationship Id="rId{rid_n}" Type="{R}/image" Target="media/{fn}"/></Relationships>')
        icon_rid[name] = f'rId{rid_n}'
    return icon_rid[name]

def icon(name, pt, lower=0):
    """inline icon run, pt = edge length in points, lower = half-points to drop below baseline"""
    rid = icon_rel(name)
    pic_id[0] += 1
    i = pic_id[0]
    emu = int(pt * 12700)
    pos = f'<w:rPr><w:position w:val="-{lower}"/></w:rPr>' if lower else ''
    return (f'<w:r>{pos}<w:drawing><wp:inline distT="0" distB="0" distL="0" distR="0">'
            f'<wp:extent cx="{emu}" cy="{emu}"/><wp:docPr id="{i}" name="图标 {i}"/>'
            '<a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">'
            f'<pic:pic><pic:nvPicPr><pic:cNvPr id="{i}" name="{name}.png"/><pic:cNvPicPr/></pic:nvPicPr>'
            f'<pic:blipFill><a:blip r:embed="{rid}"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill>'
            f'<pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="{emu}" cy="{emu}"/></a:xfrm>'
            '<a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr></pic:pic>'
            '</a:graphicData></a:graphic></wp:inline></w:drawing></w:r>')

# ---------- text helpers ----------
def run(text, sz=21, color=TEXT, b=False, font=HEI, shd=None, spacing=None, pos=None):
    rpr = f'<w:rFonts w:ascii="{font}" w:hAnsi="{font}" w:eastAsia="{font}" w:cs="{font}"/>'
    if b: rpr += '<w:b/><w:bCs/>'
    rpr += f'<w:color w:val="{color}"/>'
    if spacing: rpr += f'<w:spacing w:val="{spacing}"/>'
    if pos: rpr += f'<w:position w:val="{pos}"/>'
    rpr += f'<w:sz w:val="{sz}"/><w:szCs w:val="{sz}"/>'
    if shd:  # 底色标签：同色边框撑出内边距，避免首尾空格底色错位
        rpr += (f'<w:bdr w:val="single" w:sz="18" w:space="0" w:color="{shd}"/>'
                f'<w:shd w:val="clear" w:color="auto" w:fill="{shd}"/>')
    text = text.replace('&', '&amp;').replace('<', '&lt;')
    return f'<w:r><w:rPr>{rpr}</w:rPr><w:t xml:space="preserve">{text}</w:t></w:r>'

def para(runs, jc='left', before=0, after=0, line=None, ind_left=0, keep=False, hanging=0):
    """文字段落用“固定值”行距（与字体无关，Word/WPS/LibreOffice 排版一致）；含图标的段落用单倍行距"""
    body_xml = ''.join(runs)
    if '<w:drawing' in body_xml:
        rule = f' w:line="{line or 240}" w:lineRule="auto"/>'
    else:
        sizes = [int(x) for x in re.findall(r'<w:sz w:val="(\d+)"', body_xml)] or [21]
        exact = max(20, round(max(sizes) * 10 * 1.3 * (line or 240) / 240))
        rule = f' w:line="{exact}" w:lineRule="exact"/>'
    sp = f'<w:spacing w:before="{before}" w:after="{after}"' + rule
    ind = (f'<w:ind w:left="{ind_left}" w:hanging="{hanging}"/>' if hanging else
           f'<w:ind w:left="{ind_left}" w:firstLine="0"/>' if ind_left else '<w:ind w:firstLine="0"/>')
    kn = '<w:keepNext/>' if keep else ''
    return (f'<w:p><w:pPr>{kn}<w:snapToGrid w:val="0"/>{sp}{ind}<w:jc w:val="{jc}"/></w:pPr>'
            + body_xml + '</w:p>')

# 页面填满：把每页剩余高度均摊到该页各行的上下内边距（由 fill_pages 测量后传入，单位 twips）
E1 = int(os.environ.get('FILL_P1', '0'))
E2 = int(os.environ.get('FILL_P2', '0'))

def borders(tag, spec):
    """spec: dict side -> (val, sz, color)"""
    return f'<w:{tag}>' + ''.join(
        f'<w:{s} w:val="{v[0]}" w:sz="{v[1]}" w:space="0" w:color="{v[2]}"/>' for s, v in spec.items()) + f'</w:{tag}>'

def tc(width, content, fill=None, span=1, valign='center', bdr=None, mar=None, vmerge=None):
    pr = f'<w:tcW w:w="{width}" w:type="dxa"/>'
    if span > 1: pr += f'<w:gridSpan w:val="{span}"/>'
    if vmerge: pr += f'<w:vMerge w:val="{vmerge}"/>'
    if bdr: pr += bdr
    if fill: pr += f'<w:shd w:val="clear" w:color="auto" w:fill="{fill}"/>'
    if mar: pr += '<w:tcMar>' + ''.join(f'<w:{k} w:w="{v}" w:type="dxa"/>' for k, v in mar.items()) + '</w:tcMar>'
    pr += f'<w:vAlign w:val="{valign}"/>'
    return f'<w:tc><w:tcPr>{pr}</w:tcPr>{content}</w:tc>'

def el(xml):
    return etree.fromstring(f'<x {NSDECL}>{xml}</x>')[0]

NONE = ('nil', 0, 'auto')
THIN = ('single', 4, LINE)
ALL_THIN = borders('tcBorders', {s: THIN for s in ('top', 'left', 'bottom', 'right')})

# ---------- load ----------
tree = etree.fromstring(docxml.encode('utf8'))
body = tree.find('w:body', ns)
title = body.find('w:p', ns)
tbl = body.find('w:tbl', ns)
rows = tbl.findall('w:tr', ns)
TW = 11114
SIDES = ('top', 'left', 'bottom', 'right')
NOTBL = borders('tblBorders', {s: NONE for s in ('top','left','bottom','right','insideH','insideV')})
GOLD_BAR = ('single', 24, GOLD)
LABEL_B = borders('tcBorders', {'top': THIN, 'left': GOLD_BAR, 'bottom': THIN, 'right': THIN})

def chip(text, sz=21, bg=CHIP_BG, fg=CHIP_FG):
    return run(text, sz, fg, b=True, shd=bg)

def bar(color=GOLD):
    return run('▍', 22, color)

# ---------- A. 页首横幅：藏蓝底 + 金色细线 ----------
banner = f'''<w:tbl><w:tblPr><w:tblW w:w="{TW}" w:type="dxa"/><w:jc w:val="center"/>{NOTBL}
<w:tblLayout w:type="fixed"/><w:tblCellMar><w:left w:w="0" w:type="dxa"/><w:right w:w="0" w:type="dxa"/></w:tblCellMar></w:tblPr>
<w:tblGrid><w:gridCol w:w="1700"/><w:gridCol w:w="{TW-1700}"/></w:tblGrid>
<w:tr><w:trPr><w:trHeight w:val="1760" w:hRule="exact"/><w:jc w:val="center"/></w:trPr>
{tc(1700, para([icon('ruler-measure-FFFFFF', 54)], jc='center'), fill=NAVY)}
{tc(TW-1700,
    para([run('机械类专业核心基础课', 20, GOLD, b=True, spacing=20), run('  ·  理实一体化教学设计', 20, 'B4C3DA', spacing=20)], after=60)
    + para([run('《机械制图》教学设计', 48, 'FFFFFF', b=True, spacing=30)], after=60)
    + para([run('第二章  正投影作图基础', 24, 'D6DEEB', spacing=10)]),
    fill=NAVY, mar={'left': 60})}
</w:tr>
<w:tr><w:trPr><w:trHeight w:val="70" w:hRule="exact"/><w:jc w:val="center"/></w:trPr>
{tc(TW, para([run('', 2)]), fill=GOLD, span=2)}
</w:tr></w:tbl>'''
body.replace(title, el(banner))
body.find('w:tbl', ns).addnext(el(para([run('', 8)], line=240)))

# ---------- B. 分区标题栏（第一级：藏蓝） ----------
def section_row(row, icon_name, num, text):
    for tcx in row.findall('w:tc', ns):
        row.remove(tcx)
    row.find('w:trPr/w:trHeight', ns).set(q('val'), '560')
    cell = tc(TW, para([icon(icon_name, 15, lower=6), run(f'  {num}', 26, GOLD, b=True),
                        run(f'  {text}', 26, 'FFFFFF', b=True, spacing=40)], ind_left=160),
              fill=NAVY, span=8, bdr=borders('tcBorders', {s: ('single', 4, NAVY) for s in SIDES}))
    row.append(el(cell))

section_row(rows[0], 'info-circle-FFFFFF', '01', '授课信息')
section_row(rows[5], 'chalkboard-FFFFFF', '02', '教学设计')

# ---------- C. 授课信息（第二级：浅蓝灰标签，无图标） ----------
def label(text, width, span=1):
    return tc(width, para([run(text, 22, NAVY, b=True)], jc='center'), fill=LABEL, span=span, bdr=ALL_THIN)

def value(width, paras, span=1):
    return tc(width, ''.join(paras), span=span, bdr=ALL_THIN, fill='FFFFFF')

def v(text, **kw):
    return para([run(text, 22, TEXT)], jc='center', **kw)

def hint():
    return para([run('待填写', 21, HINT)], jc='center')

info = [
    [label('模块名称', 1506), value(4374, [v('第二章  正投影作图基础')], 3),
     label('授课时间', 1635, 2), value(3599, [v('2026-2027 学年第二学期')], 2)],
    [label('任务名称', 1506),
     value(4374, [v('第1节  投影法概述', after=40), v('第2节  三视图的形成及其对应关系')], 3),
     label('授课地点', 1635, 2), value(3599, [hint()], 2)],
    [label('授课时数', 1506), value(4374, [para([run('2', 28, NAVY, b=True), run(' 学时', 22, TEXT)], jc='center')], 3),
     label('授课类型', 1635, 2), value(3599, [para([chip('理实一体', 21, NAVY, 'FFFFFF')], jc='center')], 2)],
    [label('授课班级', 1506), value(9608, [hint()], 7)],
]
for row, cells in zip(rows[1:5], info):
    for tcx in row.findall('w:tc', ns):
        row.remove(tcx)
    row.find('w:trPr/w:trHeight', ns).set(q('val'), '560')
    for c in cells:
        row.append(el(c))

# ---------- D. 内容分析：三排卡片，长短均衡 ----------
row = rows[6]
old_label, old_content = row.findall('w:tc', ns)
row.find('w:trPr/w:trHeight', ns).set(q('val'), '2308')

def group_label_xml(width, icon_name, text, vmerge=None, note=None):
    lines = [text[:2], text[2:]] if len(text) == 4 else [text]
    top = para([run(note, 16, '8A99B3')], jc='center', after=160) if note else ''
    return tc(width, top + para([icon(icon_name, 22)], jc='center', after=60)
              + ''.join(para([run(x, 24, NAVY, b=True, spacing=40)], jc='center') for x in lines),
              fill=LABEL, vmerge=vmerge, bdr=LABEL_B)

lab = group_label_xml(1506, 'file-analytics-1F3864', '内容分析', vmerge='restart')

# 卡片表与外层单元格“贴合”：外层单元格边距为 0、不留空段；卡片表外沿不画线（由外层单元格边框收口），
# 卡片之间只用一条细线分隔；学习内容的三步作为同一张表的三列，不再三层嵌套。
CW = 9608
COL = CW // 6
GRID = [COL] * 5 + [CW - COL * 5]
MAR = {'top': 120 + E1, 'left': 200, 'bottom': 130 + E1, 'right': 200}
BODY = dict(line=300, after=0)

def edges(top=THIN, left=THIN, bottom=THIN, right=THIN):
    return borders('tcBorders', {'top': top, 'left': left, 'bottom': bottom, 'right': right})

def card_head(text, color=NAVY):
    return para([bar(GOLD if color == NAVY else color), run(' ' + text, 24, color, b=True)], after=80)

def body_p(runs, **kw):
    a = dict(BODY); a.update(kw)
    return para(runs, jc='both', **a)

def t(text, color=TEXT, **kw):
    return run(text, 20, color, **kw)

def cell(span, content, fill=CARD, bdr_xml=None, mar=MAR, valign='top'):
    return tc(sum(GRID[:span]), content, fill, span=span, valign=valign, bdr=bdr_xml, mar=mar)

pos = card_head('课程定位') + body_p([t(
    '从岗课赛证角度，机械制图是机械类专业核心基础课，机械加工、设备装配、产品测绘岗位均要求从业人员'),
    t('读懂三视图', NAVY, b=True), t('、能够'), t('按照投影规律绘制零件视图', NAVY, b=True), t('。')])
std = card_head('对接标准') + body_p([t('教学对接：', SUB)], after=40) + ''.join(
    body_p([run('■\t', 14, '8A99B3', pos=2), t(x)], after=30, ind_left=240, hanging=240) for x in
    ['机械行业制图国家标准', '1+X 机械制图职业技能等级证书考核要点', '职业院校技能大赛零部件测绘与 CAD 赛项基础要求'])
learn_head = card_head('学习内容') + body_p([t('本次课学习内容：', SUB)])

def step(n, text, extra=''):
    return para([run(n, 26, GOLD_D, b=True)], after=20) + para([t(text)], line=300) + extra

digi = card_head('数字拓展') + body_p([t('引入现代三维建模软件（'), t('中望CAD', NAVY, b=True),
    t('）三维模型转二维视图案例，对比'), t('传统手绘投影', NAVY, b=True), t('与'),
    t('数字化投影', NAVY, b=True), t('表达。')])
szh = (para([icon('flag-3-9B1C1C', 14, lower=5), run('  课程思政', 24, RED, b=True)], after=80)
       + body_p([t('结合大国工匠严谨细致的绘图识图案例，培养严谨规范的工程思维，树立工程质量意识，体会'),
                 t('图纸是工程语言', RED, b=True), t('，图纸的一丝误差会造成产品报废，树立'),
                 t('工匠精神', RED, b=True), t('。')]))

N = NONE
STEP_MAR = {'top': 60 + E1, 'left': 200, 'bottom': 140 + E1, 'right': 200}
rows_xml = [
    # 定位 | 标准（外沿上、左、右不画线）
    cell(3, pos, bdr_xml=edges(top=N, left=N)) + cell(3, std, bdr_xml=edges(top=N, right=N, left=N)),
    # 学习内容标题（与下方三步连成一块，中间不画线）
    cell(6, learn_head, bdr_xml=edges(left=N, right=N, bottom=N), mar={**MAR, 'bottom': 60}),
    # 三步：同一张表的三列，列间细线
    cell(2, step('01', '投影法分类、正投影基本特性'), bdr_xml=edges(top=N, left=N, right=N), mar=STEP_MAR)
    + cell(2, step('02', '三投影面体系建立、三视图形成过程'), bdr_xml=edges(top=N, right=N), mar=STEP_MAR)
    + cell(2, step('03', '三视图投影对应关系',
                   para([chip('长对正', 19), run(' ', 19), chip('高平齐', 19), run(' ', 19), chip('宽相等', 19)], before=80)),
           bdr_xml=edges(top=N, right=N), mar=STEP_MAR),
    # 数字拓展 | 课程思政（外沿下、左、右不画线）
    cell(3, digi, bdr_xml=edges(left=N, bottom=N)) + cell(3, szh, fill=WARM, bdr_xml=edges(left=N, right=N, bottom=N)),
]
cards = (f'<w:tbl><w:tblPr><w:tblW w:w="{CW}" w:type="dxa"/><w:tblInd w:w="0" w:type="dxa"/>{NOTBL}'
         '<w:tblLayout w:type="fixed"/><w:tblCellMar><w:left w:w="0" w:type="dxa"/><w:right w:w="0" w:type="dxa"/></w:tblCellMar></w:tblPr>'
         '<w:tblGrid>' + ''.join(f'<w:gridCol w:w="{g}"/>' for g in GRID) + '</w:tblGrid>'
         + ''.join(f'<w:tr><w:trPr><w:cantSplit/></w:trPr>{r}</w:tr>' for r in rows_xml) + '</w:tbl>')
# Word 要求单元格以段落结尾：用 1 磅高的空段收尾，视觉上不留缝
tail = para([run('', 2)], line=20)
content = tc(9608, cards + tail, span=7, valign='top', bdr=ALL_THIN,
             mar={'top': 0, 'left': 0, 'bottom': 0, 'right': 0})
row.replace(old_label, el(lab))
row.replace(old_content, el(content))

exec(open('/tmp/lo1/page2_v2.py', encoding='utf8').read())
exec(open('/tmp/lo1/page3.py', encoding='utf8').read())
exec(open('/tmp/lo1/merge.py', encoding='utf8').read())

# ---------- save ----------
open(doc_path, 'wb').write(etree.tostring(tree, xml_declaration=True, encoding='UTF-8', standalone=True))
open(rels_path, 'w', encoding='utf8').write(rels)
print('icons:', len(icon_rid), 'last docPr', pic_id[0])
