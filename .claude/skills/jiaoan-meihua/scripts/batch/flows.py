"""四份教案“教学流程”图按蓝金配色重画（文字照原图转录）。
用法：python3 -I flows.py <输出目录>   → flow_d1..d4.svg，再用 svg2png.js 转 PNG
配色与新增附图一致：藏蓝=课中/主线，青绿=课前，金=课后与重点，浅底 VF/WF/HF。
"""
import sys, re
OUT = sys.argv[1]
NAVY, TEAL, GOLD, GOLDD = '#1F3864', '#2E7D7B', '#A8841A', '#8A6A12'   # 金色比附图深一档：黑白打印时白字与箭头仍清楚
VF, WF, HF = '#E8EEF7', '#E3F1EF', '#FBF3DC'
RED, TXT, SUB = '#9B1C1C', '#333333', '#555555'
FONT = 'Noto Sans CJK SC, Microsoft YaHei, sans-serif'


def esc(s):
    return s.replace('&', '&amp;').replace('<', '&lt;').replace('>', '&gt;')

def cw(ch):
    if ch == ' ': return 0.3
    if ord(ch) < 128: return 0.56
    return 1.0

def tw(s, size):
    return sum(cw(c) for c in s) * size

PUNCT = '，。；、：）》”’！？,.;:)'
OPEN = '（“《(‘'
def tok(c):
    return (c.isascii() and c.isalnum()) or c in '=≈×.%Φ'
def wrap(s, size, width):
    """按宽度折行：标点不出现在行首、数字/字母串不拆开；\n 强制换行"""
    out = []
    for seg in s.split('\n'):
        lines, cur = [], ''
        for ch in seg:
            if tw(cur + ch, size) > width and cur:
                k = len(cur)
                if ch in PUNCT:                       # 带上前面的字一起换行
                    k = len(cur) - 1
                    while k > 0 and cur[k] in PUNCT: k -= 1
                elif tok(ch):                         # 不拆数字/字母/公式串（p=q=r≈0.82、40×25×20）
                    while k > 0 and tok(cur[k - 1]): k -= 1
                if k <= 0: k = len(cur)
                while k > 1 and cur[k - 1] in OPEN: k -= 1   # 开括号/开引号不留在行尾
                lines.append(cur[:k]); cur = cur[k:] + ch
            else:
                cur += ch
        if cur: lines.append(cur)
        out += lines
    return out

def svg(w, h, body):
    return (f'<svg xmlns="http://www.w3.org/2000/svg" width="{w}" height="{h}" viewBox="0 0 {w} {h}" font-family="{FONT}">'
            f'<rect width="{w}" height="{h}" fill="#fff"/>'
            '<defs><marker id="ar" markerUnits="userSpaceOnUse" markerWidth="14" markerHeight="14" refX="7" refY="7" orient="auto">'
            f'<path d="M0,0 L14,7 L0,14 z" fill="{GOLD}"/></marker></defs>{body}</svg>')

def rect(x, y, w, h, fill='#fff', stroke=NAVY, sw=2, rx=10, dash=None):
    d = f' stroke-dasharray="{dash}"' if dash else ''
    return f'<rect x="{x:.1f}" y="{y:.1f}" width="{w:.1f}" height="{h:.1f}" rx="{rx}" fill="{fill}" stroke="{stroke}" stroke-width="{sw}"{d}/>'

def text(x, y, s, size, fill=TXT, w=400, anchor='start'):
    return f'<text x="{x:.1f}" y="{y:.1f}" font-size="{size}" font-weight="{w}" fill="{fill}" text-anchor="{anchor}">{esc(s)}</text>'

def para(x, y, s, size, width, fill=TXT, w=400, lh=1.42, lead=None, lead_fill=None, anchor='start'):
    """折行段落；lead 为加粗前缀（如“师：”）。返回 (svg, 用掉的高度)"""
    full = (lead or '') + s
    ls = wrap(full, size, width)
    out, yy = '', y + size
    for i, ln in enumerate(ls):
        if i == 0 and lead and ln.startswith(lead):
            out += (f'<text x="{x:.1f}" y="{yy:.1f}" font-size="{size}" fill="{fill}" text-anchor="{anchor}">'
                    f'<tspan font-weight="700" fill="{lead_fill or fill}">{esc(lead)}</tspan>{esc(ln[len(lead):])}</text>')
        else:
            out += text(x, yy, ln, size, fill, w, anchor)
        yy += size * lh
    return out, len(ls) * size * lh

def chevron(x, y, w, h, fill, label, size=30, first=False):
    t = 22
    pts = [(x, y), (x + w - t, y), (x + w, y + h / 2), (x + w - t, y + h), (x, y + h)]
    if not first: pts.insert(5, (x + t, y + h / 2))
    p = ' '.join(f'{a:.1f},{b:.1f}' for a, b in pts)
    return f'<polygon points="{p}" fill="{fill}"/>' + text(x + w / 2 + (0 if first else t / 2) - t / 2, y + h / 2 + size * 0.36, label, size, '#fff', 700, 'middle')

def badge(xr, y, label):
    """右上角重点/难点标签（右对齐于 xr）"""
    fill = GOLD if label == '重点' else RED
    w = tw(label, 20) + 22
    return rect(xr - w, y, w, 30, fill, fill, 0, 15) + text(xr - w / 2, y + 22, label, 20, '#fff', 700, 'middle')

def num(cx, cy, n, fill=NAVY):
    return f'<circle cx="{cx:.1f}" cy="{cy:.1f}" r="15" fill="{fill}"/>' + text(cx, cy + 7, str(n), 20, '#fff', 700, 'middle')

def head(x, y, w, k, name, key, size=22):
    """卡片标题：编号 + 环节名；第二行：时长（灰）+ 右侧重点/难点标签。返回 (svg, 高度)"""
    m = re.match(r'^(.*)（(\d+分钟)）$', name)
    nm, t = (m.group(1), m.group(2)) if m else (name, '')
    out = num(x + 24, y + 26, k, GOLDD if key else NAVY) + text(x + 46, y + 34, nm, size, NAVY, 700)
    out += text(x + 46, y + 62, t, 18, SUB, 700)
    if key:
        fs = 18 if w > 260 else 16
        bw = tw(key, fs) + 18
        out += rect(x + w - 12 - bw, y + 44, bw, 26, GOLD if key == '重点' else RED, 'none', 0, 13)
        out += text(x + w - 12 - bw / 2, y + 62, key, fs, '#fff', 700, 'middle')
    return out, 76

def arrow(x1, y1, x2, y2, sw=4):
    return f'<line x1="{x1:.1f}" y1="{y1:.1f}" x2="{x2:.1f}" y2="{y2:.1f}" stroke="{GOLD}" stroke-width="{sw}" marker-end="url(#ar)"/>'

def save(name, s):
    open(f'{OUT}/{name}.svg', 'w', encoding='utf8').write(s)


# ======================= 第三、四章：三栏 + 课中 3×2 卡片 =======================
def flow_cards(pre, steps, post, foot):
    W = 1500
    SX, SW = 20, 214                 # 课前栏
    PX, PW = W - 20 - 214, 214       # 课后栏
    MX, MW = SX + SW + 34, PX - 34 - (SX + SW + 34)
    CG = 30                          # 卡片间距（放箭头）
    CW = (MW - 2 * CG) / 3
    b = ''
    # 顶部阶段条
    b += chevron(SX, 16, SW + 20, 52, TEAL, pre['title'], 28, True)
    b += chevron(MX - 8, 16, MW + 16, 52, NAVY, steps['title'], 28)
    b += chevron(PX - 12, 16, PW + 12, 52, GOLDD, post['title'], 28)
    top = 90
    # 课中卡片：先排版算高度
    cards = []
    for st in steps['items']:
        inner = CW - 32
        parts = [para(0, 0, st['body'], 20, inner, TXT)[1],
                 para(0, 0, st['t'], 19, inner, NAVY, lead='师：')[1]]
        if st.get('s'): parts.append(para(0, 0, st['s'], 19, inner, TEAL, lead='生：')[1])
        cards.append(76 + sum(parts) + 8 * (len(parts) - 1) + 16)
    rows = [cards[:3], cards[3:]]
    rh = [max(r) for r in rows]
    y = top
    for r in range(2):
        for c in range(3):
            k = r * 3 + c
            st = steps['items'][k]
            x = MX + c * (CW + CG)
            key = st.get('tag')
            fill, stroke = (HF, GOLDD) if key else ('#FFFFFF', '#6F84A6')
            b += rect(x, y, CW, rh[r], fill, stroke, 2.5 if key else 2, 12)
            s_, h_ = head(x, y, CW, k + 1, st['name'], key); b += s_
            yy = y + h_
            s_, h_ = para(x + 16, yy, st['body'], 20, CW - 32, TXT); b += s_; yy += h_ + 8
            s_, h_ = para(x + 16, yy, st['t'], 19, CW - 32, NAVY, lead='师：'); b += s_; yy += h_ + 8
            if st.get('s'):
                s_, h_ = para(x + 16, yy, st['s'], 19, CW - 32, TEAL, lead='生：'); b += s_
            if c < 2: b += arrow(x + CW + 3, y + rh[r] / 2, x + CW + CG - 3, y + rh[r] / 2)
        if r == 0:   # 第一行末 → 第二行首（折返箭头）
            ya = y + rh[0] + 4; yb = y + rh[0] + 26
            x3 = MX + 2 * (CW + CG) + CW / 2; x1 = MX + CW / 2
            b += (f'<polyline points="{x3:.1f},{ya:.1f} {x3:.1f},{yb - 10:.1f} {x1:.1f},{yb - 10:.1f} {x1:.1f},{yb + 2:.1f}" '
                  f'fill="none" stroke="{GOLD}" stroke-width="4" marker-end="url(#ar)"/>')
        y += rh[r] + 32
    mid_bottom = y - 32
    # 课前 / 课后栏
    def side(x, w, blocks, tag, tint, edge, tag_fill, tag_fg):
        yy = top
        out = ''
        n = len(blocks)
        avail = mid_bottom - top - 58 - 14 * n
        hs = [para(0, 0, t, 20, w - 28, TXT, lead=ld)[1] + 28 for ld, t in blocks]
        extra = max(0, (avail - sum(hs)) / n)
        for (ld, t), h in zip(blocks, hs):
            hh = h + extra
            out += rect(x, yy, w, hh, tint, edge, 2, 12)
            s_, _ = para(x + 14, yy + 12, t, 20, w - 28, TXT, lead=ld, lead_fill=edge); out += s_
            yy += hh + 14
        out += rect(x, yy, w, 44, tag_fill, tag_fill, 0, 10) + text(x + w / 2, yy + 29, tag.replace(' · ', '·'), 19, tag_fg, 700, 'middle')
        return out
    b += side(SX, SW, pre['blocks'], pre['tag'], WF, TEAL, TEAL, '#fff')
    b += side(PX, PW, post['blocks'], post['tag'], HF, GOLDD, GOLDD, '#fff')
    # 两侧箭头
    ym = top + (mid_bottom - top) / 2
    b += arrow(SX + SW + 4, ym, MX - 6, ym) + arrow(MX + MW + 6, ym, PX - 6, ym)
    # 页脚：每行可放多段（label, body），段间用竖线分隔
    fy = mid_bottom + 20
    fh = 20 + 34 * len(foot)
    b += rect(20, fy, W - 40, fh, NAVY, NAVY, 0, 12)
    yy = fy + 10
    for line in foot:
        yy += 34
        t_ = f'<text x="40" y="{yy - 8:.1f}" font-size="20" fill="#fff">'
        for i, (lab, body) in enumerate(line):
            if i: t_ += f'<tspan fill="{GOLD}">　｜　</tspan>'
            t_ += f'<tspan font-weight="700" fill="{GOLD}">{esc(lab)}</tspan>{esc(body)}'
        b += t_ + '</text>'
    return svg(W, int(fy + fh + 16), b)


FOOT_EVAL = '课前检测10% → 课中考勤·参与互评20% → 课中实操作品30% → 课后成果40%'
d3 = flow_cards(
    {'title': '课前准备', 'tag': '翻转课堂 · 以学定教', 'blocks': [
        ('教师：', '发布微课（截交线性质/两圆柱正交相贯画法）、课前检测与任务单；查看数据，以学定教'),
        ('学生：', '观看微课、完成课前检测；搜集带孔轴套、切槽滑块等零件图')]},
    {'title': '课中研学（80分钟）', 'items': [
        {'name': '情境导入（5分钟）', 'body': '带孔轴套、切槽滑块、三通管等真实零件三维动画导入，发布任务单与评价标准',
         't': '播放零件三维动画，创设岗位情境，发布任务单', 's': '观察讨论表面交线构成，明确任务（抢答积分）'},
        {'name': '理论探究一（15分钟）', 'tag': '重点', 'body': '截交线的概念与性质：截平面与立体表面的交线，封闭平面图形；形状分析（矩形、圆、椭圆等）',
         't': '模型＋三维切割演示＋板书，平台即时测练', 's': '跟随分析，完成即时练习，说出截交线性质'},
        {'name': '示范讲解·实操练习一（15分钟）', 'tag': '重点', 'body': '任务1：绘制切槽圆柱（截切体）三视图（Φ40、槽宽20、深15）',
         't': '示范找特殊点、补一般点、连线判断可见性，巡回指导', 's': '独立绘制，按评价表自检（特殊点/线型），提交'},
        {'name': '理论探究二（15分钟）', 'tag': '难点', 'body': '相贯线的概念与性质：两回转体表面的共有线，封闭空间曲线；两圆柱正交相贯的投影特征',
         't': '三维相贯演示＋认知冲突设问（“相贯线向哪侧弯曲？”）', 's': '观察思考，小组讨论辨析相贯线特征'},
        {'name': '示范讲解·实操练习二（20分钟）', 'tag': '重点·难点', 'body': '任务2：绘制两圆柱正交相贯三视图（Φ30与Φ20垂直相交），用简化画法画相贯线',
         't': '示范作图与简化画法，分层指导，组织互评', 's': '实操绘制，自检互评，提交作品'},
        {'name': '总结评价（10分钟）', 'body': '截交线/相贯线对比梳理＋作品点评＋随堂检测＋思政升华',
         't': '点评发布检测，总结升华工匠精神'},
    ]},
    {'title': '课后拓学', 'tag': '迁移应用 · 衔接组合体', 'blocks': [
        ('教师：', '发布课后作业与拓展任务，提供模型库；线上答疑、批阅反馈'),
        ('学生：', '完成作业与拓展任务，上传成果，参与线上讨论与互评')]},
    [[('评价贯穿：', FOOT_EVAL), ('重点·难点环节：', '共65分钟（占81%）')], [('思政融入：', '标准意识·工匠精神·质量至上')]])
save('flow_d3', d3)

d4 = flow_cards(
    {'title': '课前准备', 'tag': '翻转课堂 · 以学定教', 'blocks': [
        ('教师：', '发布微课（轴测图基本知识/正等轴测图画法）、课前检测与任务单；查看数据，以学定教'),
        ('学生：', '观看微课、完成课前检测；搜集轴测图在产品说明书、装配图中的应用实例')]},
    {'title': '课中研学（80分钟）', 'items': [
        {'name': '情境导入（5分钟）', 'body': '轴承座、阀体三视图与轴测图对比动画导入，发布任务单与评价标准',
         't': '播放三维旋转与三视图—轴测图对比动画，创设岗位情境，发布任务单', 's': '观察讨论两图特点，明确任务（抢答积分）'},
        {'name': '理论探究一（15分钟）', 'tag': '重点', 'body': '轴测图基本知识：轴测投影形成、轴间角与轴向伸缩系数、分类；正等测轴间角120°、p=q=r≈0.82（简化取1）',
         't': '模型＋三维投影演示＋板书轴测轴，平台即时测练', 's': '跟随分析，完成即时练习，说出轴间角与伸缩系数含义'},
        {'name': '示范讲解·实操练习一（15分钟）', 'tag': '重点', 'body': '任务1：坐标法绘制长方体切割体（40×25×20）正等轴测图',
         't': '示范“先定轴、再定点、后连线”，巡回指导', 's': '独立绘制，按评价表自检（轴测轴/比例/图面），提交'},
        {'name': '理论探究二（15分钟）', 'tag': '难点', 'body': '圆与回转体的正等轴测图：圆的正等测为椭圆（长轴≈1.22d、短轴≈0.7d）；菱形法（四心近似法）',
         't': '三维演示圆投影变形＋认知冲突设问（“圆为什么变成了椭圆？”）', 's': '观察思考，小组讨论辨析椭圆长轴方向'},
        {'name': '示范讲解·实操练习二（20分钟）', 'tag': '重点·难点', 'body': '任务2：Φ30、高40圆柱正等轴测图（菱形法）；任务3（选做）：圆柱与长方体叠加组合体',
         't': '示范菱形法画椭圆与圆柱，分层指导，组织互评', 's': '实操绘制，自检互评，提交作品'},
        {'name': '总结评价（10分钟）', 'body': '轴测图与三视图对比梳理＋作品点评＋随堂检测＋思政升华',
         't': '点评发布检测，总结升华工匠精神'},
    ]},
    {'title': '课后拓学', 'tag': '迁移应用 · 衔接斜二测', 'blocks': [
        ('教师：', '发布课后作业与拓展任务，提供模型库；线上答疑、批阅反馈'),
        ('学生：', '完成作业与拓展任务（补画轴测图），上传成果，参与线上讨论与互评')]},
    [[('评价贯穿：', FOOT_EVAL), ('重点·难点环节：', '共65分钟（占81%）')], [('思政融入：', '标准意识·工匠精神·图样直观表达')]])
save('flow_d4', d4)


# ======================= 第五章：横向七环节 + 特色设计框 =======================
def flow_d2():
    W = 1500
    b = ''
    b += chevron(20, 16, 190, 52, TEAL, '课前', 28, True)
    b += chevron(222, 16, 1060, 52, NAVY, '课中（80分钟）', 28)
    b += chevron(1294, 16, 186, 52, GOLDD, '课后', 28)
    y0, bh = 88, 196
    # 课前 / 课后
    def endbox(x, w, title, lines, hl, tint, edge):
        out = rect(x, y0, w, bh, tint, edge, 2, 12) + text(x + w / 2, y0 + 44, title, 26, edge, 700, 'middle')
        yy = y0 + 84
        for ln in lines:
            out += text(x + w / 2, yy, ln, 19, SUB, 400, 'middle'); yy += 28
        out += text(x + w / 2, yy + 4, hl, 20, RED if hl.startswith('前测') else GOLDD, 700, 'middle')
        return out
    b += endbox(20, 176, '任务准备', ['线上推送微课', '《组合体与形体', '分析法》'], '前测3题', WF, TEAL)
    b += endbox(1304, 176, '任务拓展', ['切割式组合体', '三视图绘制'], 'CAD建模对照', HF, GOLDD)
    steps = [('情境导入', '6min', None, '展示零件模型', '代入绘图员'),
             ('前测反馈', '4min', None, '回顾投影规律', '订正互评'),
             ('新知探究', '24min', '重点', '讲组合形式\n形体分析', '观察归纳'),
             ('方法学习', '12min', None, '示范画图步骤', '模仿练习'),
             ('实操训练', '24min', '难点', '巡视指导', '分组绘三视图'),
             ('成果展示', '4min', None, '组织互评', '展示PK'),
             ('总结提升', '6min', None, '点评+思政', '归纳口诀')]
    x0, x1 = 222, 1282
    gap = 8
    w = (x1 - x0 - gap * 6) / 7
    for k, (nm, t, tag, shi, sheng) in enumerate(steps):
        x = x0 + k * (w + gap)
        fill, edge = (HF, GOLDD) if tag else ('#FFFFFF', '#6F84A6')
        b += rect(x, y0, w, bh, fill, edge, 2.5 if tag else 2, 12)
        b += num(x + 22, y0 + 24, k + 1, GOLDD if tag else NAVY)
        b += text(x + w / 2, y0 + 72, nm, 25, NAVY, 700, 'middle')
        b += text(x + w / 2, y0 + 102, t, 20, SUB, 700, 'middle')
        if tag:
            b += rect(x + w / 2 - 34, y0 + 112, 68, 28, GOLD if tag == '重点' else RED, 'none', 0, 14)
            b += text(x + w / 2, y0 + 133, tag, 18, '#fff', 700, 'middle')
        s_, _ = para(x + w / 2, y0 + 146, shi, 17, w - 10, NAVY, lead='师：', anchor='middle')
        b += s_
    # 学生活动线
    ly = y0 + bh + 18
    b += f'<line x1="{x0}" y1="{ly}" x2="{x1}" y2="{ly}" stroke="{TEAL}" stroke-width="2.5"/>'
    for k, st in enumerate(steps):
        x = x0 + k * (w + gap) + w / 2
        b += f'<circle cx="{x:.1f}" cy="{ly}" r="5" fill="{TEAL}"/>'
        b += (f'<text x="{x:.1f}" y="{ly + 32}" font-size="19" fill="{TEAL}" text-anchor="middle">'
              f'<tspan font-weight="700">生：</tspan>{esc(st[4])}</text>')
    b += arrow(196 + 4, y0 + bh / 2, x0 - 4, y0 + bh / 2) + arrow(x1 + 4, y0 + bh / 2, 1304 - 4, y0 + bh / 2)
    # 特色设计框
    by = ly + 58
    items = [('【岗课融合】', '以轴承座、支架等真实零件组合体为载体，对接机械绘图员“识图—分解—绘图—标注”岗位工作过程，学生扮绘图员完成三视图表达。'),
             ('【以赛促学】', '对接机械制图技能大赛与CAD等级考试中组合体视图表达要求，小组“三视图作品PK”即赛场模拟。'),
             ('【信息化】', '微课前置、组合体三维模型动画、Inventor建模验证视图，学习通实时采集前测/随堂/互评数据，以学定教。'),
             ('【教学创新】', '重点（形体分析法24min）+难点（三视图规范绘制24min）合计48min，占课中60%；教师退为“咨询师/教练”，学生“做中学、错中悟”。'),
             ('【课程思政】', '融入严谨细致、一丝不苟的工匠精神与按国标出图的“工程语言”规范意识；归纳口诀，增强职业自信。')]
    inner = W - 40 - 60
    body, yy = '', by + 76
    for lab, t in items:
        ls = wrap(lab + t, 20, inner)
        for i, ln in enumerate(ls):
            if i == 0:
                body += (f'<text x="50" y="{yy:.1f}" font-size="20" fill="{TXT}"><tspan font-weight="700" fill="{GOLDD}">{esc(lab)}</tspan>'
                         f'{esc(ln[len(lab):])}</text>')
            else:
                body += text(50, yy, ln, 20, TXT)
            yy += 30
        yy += 6
    tl = '时间分配：导入6+前测4+探究24+方法12+训练24+展示4+总结6 = 80min，重难点占60%。'
    body += text(W / 2, yy + 8, tl, 21, RED, 700, 'middle')
    bh2 = yy + 28 - by
    b += rect(20, by, W - 40, bh2, '#FFFFFF', GOLD, 3, 14, '10 7')
    b += text(W / 2, by + 40, '教学能力大赛特色设计（岗课赛证融合 · 信息化 · 思政）', 25, NAVY, 700, 'middle')
    b += body
    return svg(W, int(by + bh2 + 16), b)
save('flow_d2', flow_d2())


# ======================= 第二章：课前 | 课中 2×4 | 课后 + 重点拓展 =======================
def flow_d1():
    W = 1500
    b = ''
    b += chevron(20, 16, 210, 52, TEAL, '课前', 28, True)
    b += chevron(242, 16, 1018, 52, NAVY, '课中（80分钟）', 28)
    b += chevron(1272, 16, 208, 52, GOLDD, '课后', 28)
    steps = [('任务领受（5分钟）', None, '发布识读与绘制点线面投影图工作任务'),
             ('任务计划（5分钟）', None, '分析任务，小组制定学习计划与绘图步骤'),
             ('快速检测（5分钟）', None, '投影法与三视图基础知识预习抢答'),
             ('机理分析（18分钟）', '重点', '点的三面投影规律与坐标特性；直线投影特性（一般位置/平行线/垂直线）；平面投影特性（一般位置/平行面/垂直面）'),
             ('快修训练（27分钟）', '重点·难点', '教师示范点线面三面投影图绘制＋学生实操运用投影规律作图，巡回指导纠错'),
             ('分组对抗（8分钟）', None, '小组竞赛，判断直线平面空间位置，比拼速度准确性'),
             ('评价反思（7分钟）', None, '作品互评，总结点线面投影常见错误'),
             ('一招制胜（5分钟）', None, '技巧口诀——点投影看坐标，直线投影看倾角，平面投影看类似，特殊位置有积聚')]
    x0, x1, gap = 262, 1250, 26
    w = (x1 - x0 - 3 * gap) / 4
    y0 = 88
    hs = [76 + para(0, 0, s[2], 20, w - 28)[1] + 16 for s in steps]
    rh = [max(hs[:4]), max(hs[4:])]
    y = y0
    for r in range(2):
        for c in range(4):
            k = r * 4 + c
            nm, tag, body = steps[k]
            x = x0 + c * (w + gap)
            fill, edge = (HF, GOLDD) if tag else ('#FFFFFF', '#6F84A6')
            b += rect(x, y, w, rh[r], fill, edge, 2.5 if tag else 2, 12)
            s_, h_ = head(x, y, w, k + 1, nm, tag, 21); b += s_
            yy = y + h_
            s_, _ = para(x + 14, yy, body, 20, w - 28, TXT); b += s_
            if c < 3: b += arrow(x + w + 3, y + rh[r] / 2, x + w + gap - 3, y + rh[r] / 2)
        if r == 0:
            ya = y + rh[0] + 4; yb = y + rh[0] + 26
            xa = x0 + 3 * (w + gap) + w / 2; xb = x0 + w / 2
            b += (f'<polyline points="{xa:.1f},{ya:.1f} {xa:.1f},{yb - 10:.1f} {xb:.1f},{yb - 10:.1f} {xb:.1f},{yb + 2:.1f}" '
                  f'fill="none" stroke="{GOLD}" stroke-width="4" marker-end="url(#ar)"/>')
        y += rh[r] + 32
    bottom = y - 32
    # 课前：两块；课后：一块
    hh = (bottom - y0 - 14) / 2
    for i, (ld, t) in enumerate((('任务准备——', '预习点线面投影基本知识，观看三面投影动画'), ('本节课设计——', '上传预习笔记，记录疑问'))):
        yy = y0 + i * (hh + 14)
        b += rect(20, yy, 210, hh, WF, TEAL, 2, 12)
        s_, h_ = para(34, 0, t, 21, 182, TXT, lead=ld, lead_fill=TEAL)
        b += f'<g transform="translate(0,{yy + (hh - h_) / 2 - 6:.1f})">{s_}</g>'
        b += arrow(234, yy + hh / 2, x0 - 6, yy + hh / 2)
    b += rect(1272, y0, 208, bottom - y0, HF, GOLDD, 2, 12)
    s_, h_ = para(1286, 0, '完成习题集点线面投影练习；预习立体的投影', 21, 180, TXT, lead='任务拓展——', lead_fill=GOLDD)
    b += f'<g transform="translate(0,{y0 + (bottom - y0 - h_) / 2 - 6:.1f})">{s_}</g>'
    b += arrow(x1 + 4, y0 + (bottom - y0) / 2, 1272 - 6, y0 + (bottom - y0) / 2)
    # 重点拓展条
    fy = bottom + 20
    b += rect(20, fy, W - 40, 56, NAVY, NAVY, 0, 12)
    b += rect(20, fy, 190, 56, GOLDD, GOLDD, 0, 12) + text(115, fy + 37, '重点拓展', 25, '#fff', 700, 'middle')
    b += text(232, fy + 37, '★ 运用点线面投影规律独立完成组合体视图中线面分析与投影作图。', 23, '#fff', 700)
    return svg(W, int(fy + 72), b)
save('flow_d1', flow_d1())
print('ok')
