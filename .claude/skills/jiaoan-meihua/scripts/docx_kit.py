"""教案美化通用工具：直接改 word/document.xml（解包后的目录），只动格式、不动行列结构。

用法（在自己的处理脚本里）：
    import sys; sys.path.insert(0, '<skill>/scripts')
    from docx_kit import Doc, C
    d = Doc('unpacked_dir', icon_dir='<skill>/assets/icons')
    rows = d.main_rows()
    ... 调用下面的函数 ...
    d.save()

设计要点（为什么这样做）见 references/pitfalls.md。
"""
import os, re, shutil, struct
from lxml import etree

W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
A = 'http://schemas.openxmlformats.org/drawingml/2006/main'
WP = 'http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing'
ns = {'w': W}
q = lambda n: f'{{{W}}}{n}'
NSDECL = (f'xmlns:w="{W}" xmlns:r="{R}" xmlns:wp="{WP}" xmlns:a="{A}" '
          'xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture"')


class C:
    """蓝金配色（用户最终选定）。一主一辅一中性：藏蓝=第一级，浅蓝灰=第二级，白=正文；金色只做点缀。"""
    NAVY, GOLD, GOLD_D = '1F3864', 'C9A227', 'A8841A'
    LABEL, CARD, LINE = 'EEF3F9', 'F7F9FC', 'D5DEEA'
    TEXT, SUB, HINT = '333333', '666666', 'AAAAAA'
    CHIP_BG, CHIP_FG = 'F6EDD2', '7A5C00'
    RED, WARM = '9B1C1C', 'FAF6F0'
    INTENT_BG, INTENT_FG = 'FBF8F1', '5C5346'
    FONT = '微软雅黑'


NONE = ('nil', 0, 'auto')
THIN = ('single', 4, C.LINE)
GOLD_BAR = ('single', 24, C.GOLD)
SIDES = ('top', 'left', 'bottom', 'right')

PORD = ['pStyle', 'keepNext', 'keepLines', 'pageBreakBefore', 'framePr', 'widowControl', 'numPr',
        'suppressLineNumbers', 'pBdr', 'shd', 'tabs', 'suppressAutoHyphens', 'kinsoku', 'wordWrap',
        'overflowPunct', 'topLinePunct', 'autoSpaceDE', 'autoSpaceDN', 'bidi', 'adjustRightInd', 'snapToGrid',
        'spacing', 'ind', 'contextualSpacing', 'mirrorIndents', 'suppressOverlap', 'jc', 'textDirection',
        'textAlignment', 'textboxTightWrap', 'outlineLvl', 'divId', 'cnfStyle', 'rPr', 'sectPr', 'pPrChange']
TCPR = ['cnfStyle', 'tcW', 'gridSpan', 'hMerge', 'vMerge', 'tcBorders', 'shd', 'noWrap', 'tcMar',
        'textDirection', 'tcFitText', 'vAlign', 'hideMark']
BORDER_ORDER = ['top', 'start', 'left', 'bottom', 'end', 'right', 'insideH', 'insideV', 'tl2br', 'tr2bl']


def el(xml):
    return etree.fromstring(f'<x {NSDECL}>{xml}</x>')[0]


def els(xml):
    return list(etree.fromstring(f'<x {NSDECL}>{xml}</x>'))


def put(parent, name, order, node):
    """按 schema 顺序插入/替换子元素（顺序错了 Word 会报文件损坏）"""
    old = parent.find(f'w:{name}', ns)
    if old is not None: parent.remove(old)
    idx = order.index(name); pos = len(parent)
    for i, ch in enumerate(parent):
        ln = etree.QName(ch).localname
        if ln in order and order.index(ln) > idx: pos = i; break
    parent.insert(pos, node)
    return node


# ---------------- 文字与段落 ----------------
def run(text, sz=21, color=C.TEXT, b=False, font=C.FONT, shd=None, spacing=None, pos=None):
    rpr = f'<w:rFonts w:ascii="{font}" w:hAnsi="{font}" w:eastAsia="{font}" w:cs="{font}"/>'
    if b: rpr += '<w:b/><w:bCs/>'
    rpr += f'<w:color w:val="{color}"/>'
    if spacing: rpr += f'<w:spacing w:val="{spacing}"/>'
    if pos: rpr += f'<w:position w:val="{pos}"/>'
    rpr += f'<w:sz w:val="{sz}"/><w:szCs w:val="{sz}"/>'
    if shd:   # 底色标签：同色边框撑出内边距，避免首尾空格底色错位
        rpr += (f'<w:bdr w:val="single" w:sz="18" w:space="0" w:color="{shd}"/>'
                f'<w:shd w:val="clear" w:color="auto" w:fill="{shd}"/>')
    text = text.replace('&', '&amp;').replace('<', '&lt;')
    return f'<w:r><w:rPr>{rpr}</w:rPr><w:t xml:space="preserve">{text}</w:t></w:r>'


def para(runs, jc='left', before=0, after=0, line=None, ind_left=0, keep=False, hanging=0):
    """文字段落一律“固定值”行距：与字体无关，Word/WPS/LibreOffice 行高一致，分页才可预测。
    含图片的段落用单倍行距（固定值会把图片裁掉）。line 按“240=单倍”的倍数理解。"""
    body = ''.join(runs)
    if '<w:drawing' in body:
        rule = f' w:line="{line or 240}" w:lineRule="auto"/>'
    else:
        sizes = [int(x) for x in re.findall(r'<w:sz w:val="(\d+)"', body)] or [21]
        exact = max(20, round(max(sizes) * 10 * 1.3 * (line or 240) / 240))
        rule = f' w:line="{exact}" w:lineRule="exact"/>'
    sp = f'<w:spacing w:before="{before}" w:after="{after}"' + rule
    ind = (f'<w:ind w:left="{ind_left}" w:hanging="{hanging}"/>' if hanging else
           f'<w:ind w:left="{ind_left}" w:firstLine="0"/>' if ind_left else '<w:ind w:firstLine="0"/>')
    kn = '<w:keepNext/>' if keep else ''
    return (f'<w:p><w:pPr>{kn}<w:snapToGrid w:val="0"/>{sp}{ind}<w:jc w:val="{jc}"/></w:pPr>{body}</w:p>')


def chip(text, sz=20, bg=C.CHIP_BG, fg=C.CHIP_FG):
    return run(text, sz, fg, b=True, shd=bg)


def item(num, text, sz=19, color=C.TEXT):
    """编号条目：金色序号 + 悬挂缩进（换行与文字对齐）"""
    w = 300 if len(num) <= 2 else 420
    return para([run(num + '\t', sz, C.GOLD_D, b=True), run(text, sz, color)], line=290, after=40,
                ind_left=w, hanging=w)


def h2(text, sz=21):
    return para([run('▍', 20, C.GOLD), run(text, sz, C.NAVY, b=True)], line=260, before=120, after=60)


def h3(num, text, sz=20):
    return para([run(num + ' ', sz, C.GOLD_D, b=True), run(text, sz, C.NAVY, b=True)], line=260, before=60, after=30)


# ---------------- 边框与单元格 ----------------
def borders(tag, spec):
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


ALL_THIN = borders('tcBorders', {s: THIN for s in SIDES})
LABEL_B = borders('tcBorders', {'top': THIN, 'left': GOLD_BAR, 'bottom': THIN, 'right': THIN})


def set_tcpr(cell, fill=None, bdr_xml=None, valign=None, mar=None):
    pr = cell.find('w:tcPr', ns)
    if bdr_xml: put(pr, 'tcBorders', TCPR, el(bdr_xml))
    if fill: put(pr, 'shd', TCPR, el(f'<w:shd w:val="clear" w:color="auto" w:fill="{fill}"/>'))
    if mar: put(pr, 'tcMar', TCPR, el('<w:tcMar>' + ''.join(
        f'<w:{k} w:w="{v}" w:type="dxa"/>' for k, v in mar.items()) + '</w:tcMar>'))
    if valign: put(pr, 'vAlign', TCPR, el(f'<w:vAlign w:val="{valign}"/>'))


def set_side(cell, side, spec):
    pr = cell.find('w:tcPr', ns)
    b = pr.find('w:tcBorders', ns)
    if b is None: b = put(pr, 'tcBorders', TCPR, el('<w:tcBorders/>'))
    old = b.find(f'w:{side}', ns)
    if old is not None: b.remove(old)
    node = el(f'<w:{side} w:val="{spec[0]}" w:sz="{spec[1]}" w:space="0" w:color="{spec[2]}"/>')
    pos = len(b)
    for i, ch in enumerate(b):
        if BORDER_ORDER.index(etree.QName(ch).localname) > BORDER_ORDER.index(side): pos = i; break
    b.insert(pos, node)


def set_paras(cell, xml):
    for p in cell.findall('w:p', ns): cell.remove(p)
    for n in els(xml): cell.append(n)


def texts(node):
    return [s for s in (''.join(x.text or '' for x in p.iter(q('t'))).strip() for p in node.findall('w:p', ns)) if s]


def cant_split(row):
    trpr = row.find('w:trPr', ns)
    if trpr is None: trpr = row.makeelement(q('trPr')); row.insert(0, trpr)
    if trpr.find('w:cantSplit', ns) is None: trpr.insert(0, etree.Element(q('cantSplit')))


def unify_left_edge(rows, spec=GOLD_BAR):
    """表格最左侧统一为同一条边（用户明确指出过“最左边框颜色不统一”）"""
    for r in rows:
        set_side(r.find('w:tc', ns), 'left', spec)


# ---------------- 文档（图片/关系/保存） ----------------
class Doc:
    def __init__(self, unpacked, icon_dir=None):
        self.D = unpacked
        self.icon_dir = icon_dir
        self.doc_path = f'{unpacked}/word/document.xml'
        self.rels_path = f'{unpacked}/word/_rels/document.xml.rels'
        self.rels = open(self.rels_path, encoding='utf8').read()
        self.tree = etree.fromstring(open(self.doc_path, 'rb').read())
        self.body = self.tree.find('w:body', ns)
        self.rid_n = max(int(x) for x in re.findall(r'Id="rId(\d+)"', self.rels))
        ids = [int(x) for x in re.findall(r'<wp:docPr[^>]*id="(\d+)"', open(self.doc_path, encoding='utf8').read())]
        self.pic_id = max(ids or [0])
        self._icons = {}

    def main_table(self):
        return max(self.body.findall('w:tbl', ns), key=lambda t: len(t.findall('w:tr', ns)))

    def main_rows(self):
        return self.main_table().findall('w:tr', ns)

    def add_media(self, src, name):
        self.rid_n += 1
        os.makedirs(f'{self.D}/word/media', exist_ok=True)
        shutil.copy(src, f'{self.D}/word/media/{name}')
        self.rels = self.rels.replace('</Relationships>',
            f'<Relationship Id="rId{self.rid_n}" Type="{R}/image" Target="media/{name}"/></Relationships>')
        return f'rId{self.rid_n}'

    def inline(self, rid, cx, cy, lower=0):
        self.pic_id += 1; i = self.pic_id
        pos = f'<w:rPr><w:position w:val="-{lower}"/></w:rPr>' if lower else ''
        return (f'<w:r>{pos}<w:drawing><wp:inline distT="0" distB="0" distL="0" distR="0">'
                f'<wp:extent cx="{cx}" cy="{cy}"/><wp:docPr id="{i}" name="图示 {i}"/>'
                '<a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">'
                f'<pic:pic><pic:nvPicPr><pic:cNvPr id="{i}" name="img{i}.png"/><pic:cNvPicPr/></pic:nvPicPr>'
                f'<pic:blipFill><a:blip r:embed="{rid}"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill>'
                f'<pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="{cx}" cy="{cy}"/></a:xfrm>'
                '<a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr></pic:pic>'
                '</a:graphicData></a:graphic></wp:inline></w:drawing></w:r>')

    def icon(self, name, pt, lower=0):
        """name 形如 'target-1F3864'，对应 assets/icons/<name>.png（Tabler 图标，见 references/style.md）"""
        if name not in self._icons:
            self._icons[name] = self.add_media(f'{self.icon_dir}/{name}.png', f'ic_{name}.png')
        emu = int(pt * 12700)
        return self.inline(self._icons[name], emu, emu, lower)

    def figure(self, src, name, width_emu, caption=None):
        """居中图示段（+ 可选图注）。width_emu 请用单元格可用宽度：(tcW - 左右边距 - 40) * 635"""
        with open(src, 'rb') as f: pw, ph = struct.unpack('>II', f.read(24)[16:24])
        rid = self.add_media(src, name)
        xml = para([self.inline(rid, width_emu, int(width_emu * ph / pw))], jc='center', before=60,
                   after=20 if caption else 60)
        if caption: xml += para([run(caption, 16, '666666')], jc='center', after=80)
        return els(xml)

    def save(self):
        open(self.doc_path, 'wb').write(etree.tostring(self.tree, xml_declaration=True, encoding='UTF-8', standalone=True))
        open(self.rels_path, 'w', encoding='utf8').write(self.rels)
