"""四份教案的新增图示（SVG → PNG）。几何均按正投影/正等测关系计算。
用法：python3 -I diagrams.py <输出目录>   之后用 node+sharp 渲染 PNG
"""
import sys, math, re
import os
src = open(os.path.join(os.path.dirname(os.path.abspath(__file__)), '..', 'reference_jiaoan3', 'draw.py'), encoding='utf8').read()
exec(src.split('# ================= A')[0])          # svg/poly/line/text/iso、配色
OUT = sys.argv[1]
GRAY = '#8A99B3'
FS = 1.0
_text = text
def text(p, s_, size=26, *a, **k):
    return _text(p, s_, round(size * FS), *a, **k)

def crop(sv, x0, y0, w, h):
    sv = re.sub(r'width="[\d.]+" height="[\d.]+" viewBox="[^"]+"', f'width="{w}" height="{h}" viewBox="{x0} {y0} {w} {h}"', sv, 1)
    sv = sv.replace('<rect width=', f'<rect x="{x0}" y="{y0}" width=', 1)
    return sv.replace('fill="#fff"/>', 'fill="#F7F9FC"/>', 1)        # 缩略图底色 = 卡片底色

def save(name, s):
    open(f'{OUT}/{name}.svg', 'w', encoding='utf8').write(s)


# ---------- 知识框架链（总结评价） ----------
def wrap(t, n=17):
    out, cur = [], ''
    for part in re.split(r'(?<=[，、；/])|(?= · )', t):
        if len(cur) + len(part) > n and cur:
            out.append(cur.strip()); cur = part.replace(' · ', '', 1) if part.startswith(' · ') else part
        else: cur += part
    if cur.strip(): out.append(cur.strip())
    return out

def chain(title, nodes, foot=None):
    W, top, gap = 480, 84, 36
    b = text((W / 2, 50), title, 34)
    y = top
    for k, (t, sub) in enumerate(nodes):
        last = k == len(nodes) - 1
        fill, stroke, tc_ = (('#FBF3DC', GOLDD, '#8A6A12') if last else
                             (VF, NAVY, NAVY) if k < 2 else (WF, TEAL, TEAL))
        sl = wrap(sub) if sub else []
        bh = 62 + 32 * len(sl)
        b += f'<rect x="20" y="{y}" width="{W - 40}" height="{bh}" rx="14" fill="{fill}" stroke="{stroke}" stroke-width="3"/>'
        b += text((W / 2, y + 44), t, 32, tc_)
        for m, ln in enumerate(sl):
            b += text((W / 2, y + 82 + 32 * m), ln, 24, '#444', 400)
        if not last:
            b += line((W / 2, y + bh + 2), (W / 2, y + bh + gap - 6), GOLD, 5, marker='ar')
        y += bh + gap
    if foot:
        y += 14 - gap
        hh = 28 + 38 * len(foot)
        b += f'<rect x="20" y="{y}" width="{W - 40}" height="{hh}" rx="14" fill="#FFFFFF" stroke="{GOLD}" stroke-width="3" stroke-dasharray="8 6"/>'
        for m, fl in enumerate(foot):
            b += text((W / 2, y + 42 + 38 * m), fl, 28 if m == 0 else 23, NAVY, 700 if m == 0 else 400)
        y += hh
    return svg(W, int(y + 20), b)

save('fw_d1', chain('本课知识框架', [
    ('点的投影规律', '长对正 · 高平齐 · 宽相等'),
    ('各类直线的投影特性', '一般位置线 · 投影面平行线 · 投影面垂直线'),
    ('各类平面的投影特性', '一般位置面 · 投影面平行面 · 投影面垂直面'),
    ('“三等”关系口诀', '主俯长对正 · 主左高平齐 · 俯左宽相等'),
]))
save('fw_d2', chain('组合体三视图画图步骤', [
    ('形体分析', '分解基本体，分析组合形式、表面连接'),
    ('选择主视图', '反映形状特征'),
    ('选比例布图', '画中心线、基线'),
    ('逐块画图', '先大后小，先主后次'),
    ('检查加深', '核对表面连接处的投影'),
], foot=['口诀', '分解形体记关系，', '长对正、高平齐、宽相等；', '先大后小画视图，', '定形定位总尺寸']))
save('fw_d3', chain('本课知识框架', [
    ('截交线 / 相贯线的概念', '平面截切立体、两立体相交时表面产生的交线'),
    ('性质', '共有线 · 封闭性'),
    ('求作方法', '先求特殊点、再求一般点，光滑连线'),
    ('可见性判断', '判断可见性并描深'),
]))
save('fw_d4', chain('本课知识框架', [
    ('轴测图概念', '平行投影 · 直观性强、度量性差'),
    ('轴间角与伸缩系数', '正等测：轴间角 120°，简化系数 p=q=r=1'),
    ('平面立体坐标法', '定轴 → 定点 → 连线'),
    ('圆与回转体菱形法', '四心近似法画椭圆，再作公切线'),
]))


# ---------- 轴测图：正等测轴测轴（d4） ----------
def iso_axes(mini=False):
    W, H = 640, 620
    O = (320, 370); L = 210
    b = '' if mini else text((W / 2, 56), '正等轴测图的轴测轴', 36)
    ang = {'Z': 90, 'X': 210, 'Y': 330}
    for k, a in ang.items():
        r = math.radians(a)
        e = (O[0] + L * math.cos(r), O[1] - L * math.sin(r))
        b += line(O, e, NAVY, 5, marker='an')
        lab = (O[0] + (L + 34) * math.cos(r), O[1] - (L + 34) * math.sin(r) + 12)
        b += text(lab, k, 34)
    # 120° 角标
    R = 72
    for a1, a2, mid in ((90, 210, 150), (210, 330, 270), (330, 450, 30)):
        a1, a2 = a1 + 7, a2 - 7
        x1, y1 = O[0] + R * math.cos(math.radians(a1)), O[1] - R * math.sin(math.radians(a1))
        x2, y2 = O[0] + R * math.cos(math.radians(a2)), O[1] - R * math.sin(math.radians(a2))
        b += f'<path d="M{x1:.1f},{y1:.1f} A{R},{R} 0 0 0 {x2:.1f},{y2:.1f}" fill="none" stroke="{GOLD}" stroke-width="4"/>'
        lx, ly = O[0] + (R + 46) * math.cos(math.radians(mid)), O[1] - (R + 46) * math.sin(math.radians(mid)) + 10
        b += text((lx, ly), '120°', 26, GOLDD)
    b += f'<circle cx="{O[0]}" cy="{O[1]}" r="7" fill="{NAVY}"/>' + text((O[0] - 26, O[1] - 14), 'O', 26)
    b += text((W / 2, 590), 'p = q = r = 1' if mini else '简化轴向伸缩系数 p = q = r = 1', 28, TEAL)
    return crop(svg(W, H, b), 40, 80, 560, 540) if mini else svg(W, H, b)
save('iso_axes', iso_axes())
FS = 1.45; save('m_iso_axes', iso_axes(True)); FS = 1.0


# ---------- 轴测图：平行于坐标面的圆的正等测（d4） ----------
def iso_circles(mini=False):
    s = 150; P = iso(330, 300, s)
    b = '' if mini else text((330, 50), '平行于坐标面的圆的正等测', 34)
    A = 1.6
    # 立方体可见三面：顶面 z=A，左前面 x=A，右前面 y=A
    top = [(0, 0, A), (A, 0, A), (A, A, A), (0, A, A)]
    lf = [(A, 0, 0), (A, A, 0), (A, A, A), (A, 0, A)]
    rf = [(0, A, 0), (A, A, 0), (A, A, A), (0, A, A)]
    b += poly([P(p) for p in lf], '#E3F1EF', TEAL, 3) + poly([P(p) for p in rf], '#E8EEF7', NAVY, 3) + poly([P(p) for p in top], '#FBF3DC', GOLDD, 3)
    c, r = A / 2, A / 2 * 0.92
    def ell(fn, color):
        pts = [P(fn(math.radians(t))) for t in range(0, 361, 6)]
        return '<polyline points="' + ' '.join(f'{x:.1f},{y:.1f}' for x, y in pts) + f'" fill="none" stroke="{color}" stroke-width="5"/>'
    b += ell(lambda t: (c + r * math.cos(t), c + r * math.sin(t), A), '#C0392B')
    b += ell(lambda t: (A, c + r * math.cos(t), c + r * math.sin(t)), '#C0392B')
    b += ell(lambda t: (c + r * math.cos(t), A, c + r * math.sin(t)), '#C0392B')
    if mini: return crop(svg(660, 630, b), 105, 45, 450, 510)
    b += text((330, 600), '三个方向的圆均变成形状相同的椭圆', 26, '#555', 400)
    return svg(660, 630, b)
save('iso_circles', iso_circles())
save('m_iso_circles', iso_circles(True))


# ---------- 点、线、面（d1） ----------
def point_proj(mini=False):
    s = 52; P = iso(300, 372, s)
    XM, YM, ZM = 6.0, 5.0, 4.8
    b = '' if mini else text((300, 44), '点的三面投影', 34)
    b += poly([P((0, 0, 0)), P((XM, 0, 0)), P((XM, YM, 0)), P((0, YM, 0))], HF, GOLDD, 3)
    b += poly([P((0, 0, 0)), P((XM, 0, 0)), P((XM, 0, ZM)), P((0, 0, ZM))], VF, NAVY, 3)
    b += poly([P((0, 0, 0)), P((0, YM, 0)), P((0, YM, ZM)), P((0, 0, ZM))], WF, TEAL, 3)
    A = (3.4, 2.8, 2.9)
    a, a1, a2 = (A[0], A[1], 0), (A[0], 0, A[2]), (0, A[1], A[2])
    for p_ in (a, a1, a2):
        b += line(P(A), P(p_), GRAY, 2.5, '8 6')
    # 投影连线到轴
    for p_, q_ in ((a, (A[0], 0, 0)), (a1, (A[0], 0, 0)), (a1, (0, 0, A[2])), (a2, (0, 0, A[2])), (a, (0, A[1], 0)), (a2, (0, A[1], 0))):
        b += line(P(p_), P(q_), '#B4C3DA', 2)
    for p_, lab, col in ((A, 'A', NAVY), (a, 'a', '#C0392B'), (a1, "a′", '#C0392B'), (a2, 'a″', '#C0392B')):
        x, y = P(p_)
        b += f'<circle cx="{x:.1f}" cy="{y:.1f}" r="8" fill="{col}"/>' + text((x + 22, y - 10), lab, 30, col)
    b += text(P((XM + 0.5, 0, 0)), 'X', 28) + text(P((0, 0, ZM + 0.35)), 'Z', 28) + text(P((0, YM + 0.5, 0)), 'Y', 28)
    b += text(P((XM - 0.6, 0, ZM - 0.6)), 'V', 34) + text(P((XM - 0.8, YM - 0.6, 0)), 'H', 34, GOLDD) + text(P((0, YM - 0.5, ZM - 0.7)), 'W', 34, TEAL)
    return crop(svg(600, 690, b), 0, 70, 600, 620) if mini else svg(600, 690, b)
save('pt_proj', point_proj())
FS = 1.35; save('m_pt_proj', point_proj(True)); FS = 1.0

def three_cases(title, kind, mini=False):
    """直线/平面相对 H 面的三种位置（等轴测示意）"""
    W, H = 900, 430
    b = '' if mini else text((W / 2, 48), title, 34)
    labels = {'line': [('平行', '反映实长'), ('垂直', '积聚成点'), ('倾斜', '长度缩短')],
              'plane': [('平行', '反映实形'), ('垂直', '积聚成直线'), ('倾斜', '类似形')]}[kind]
    objs = {'line': [[(2.3, 0.7, 1.9), (0.7, 2.3, 1.9)], [(1.5, 1.5, 2.3), (1.5, 1.5, 0.9)], [(2.3, 0.7, 0.7), (0.7, 2.3, 2.3)]],
            'plane': [[(2.4, 0.7, 1.9), (0.7, 1.1, 1.9), (1.7, 2.4, 1.9)], [(2.3, 0.7, 0.8), (0.7, 2.3, 0.8), (1.5, 1.5, 2.4)],
                      [(2.5, 0.6, 0.5), (0.6, 1.0, 2.2), (1.8, 2.5, 1.0)]]}[kind]
    for k, ((pos, res), ob) in enumerate(zip(labels, objs)):
        cx = 150 + k * 300
        P = iso(cx, 200, 40)
        b += poly([P((0, 0, 0)), P((3, 0, 0)), P((3, 3, 0)), P((0, 3, 0))], HF, GOLDD, 2.5)
        b += text(P((2.75, 0.2, 0.08)), 'H', 22, GOLDD)
        pr = [(x, y, 0) for x, y, z in ob]
        for a_, p_ in zip(ob, pr):
            b += line(P(a_), P(p_), GRAY, 2.2, '7 5')
        if kind == 'line':
            if pr[0] == pr[1]:
                x, y = P(pr[0]); b += f'<circle cx="{x:.1f}" cy="{y:.1f}" r="8" fill="#C0392B"/>'
            else:
                b += line(P(pr[0]), P(pr[1]), '#C0392B', 7)
            b += line(P(ob[0]), P(ob[1]), NAVY, 7)
        else:
            b += poly([P(p_) for p_ in pr], '#E8A39B', '#C0392B', 4)
            b += poly([P(p_) for p_ in ob], '#CAD8EE', NAVY, 3.5, op=0.92)
        if mini: b += text((cx, 372), pos + '·' + {'反映实长': '实长', '积聚成点': '积聚', '长度缩短': '缩短', '反映实形': '实形', '积聚成直线': '积聚', '类似形': '类似'}[res], 30, GOLDD)
        else: b += text((cx, 368), pos + ' H 面', 28) + text((cx, 408), res, 26, GOLDD)
    return crop(svg(W, H, b), 0, 80, W, 315) if mini else svg(W, H, b)
save('line_cases', three_cases('直线的投影特性', 'line'))
save('plane_cases', three_cases('平面的投影特性', 'plane'))
FS = 1.7; save('m_line_cases', three_cases('', 'line', True)); save('m_plane_cases', three_cases('', 'plane', True)); FS = 1.0

def convert(mini=False):
    W, H = 780, 440
    b = '' if mini else text((W / 2, 50), '空间位置 ⇄ 投影图', 34)
    s = 50; P = iso(165, 230, s)
    X, Y, Z1, Z2, Y1 = 2.4, 1.8, 0.8, 1.6, 0.9         # 下层 y 0..1.8 / 上层（靠后）y 0..0.9
    faces = [([(0, Y1, Z1), (X, Y1, Z1), (X, Y, Z1), (0, Y, Z1)], '#E8EEF7'),
             ([(0, 0, Z2), (X, 0, Z2), (X, Y1, Z2), (0, Y1, Z2)], '#E8EEF7'),
             ([(0, Y, 0), (X, Y, 0), (X, Y, Z1), (0, Y, Z1)], '#CAD8EE'),
             ([(0, Y1, Z1), (X, Y1, Z1), (X, Y1, Z2), (0, Y1, Z2)], '#CAD8EE'),
             ([(X, 0, 0), (X, Y, 0), (X, Y, Z1), (X, Y1, Z1), (X, Y1, Z2), (X, 0, Z2)], '#AFC3DF')]
    for f, c in faces: b += poly([P(p) for p in f], c, NAVY, 3)
    if not mini: b += text((165, 410), '空间形体', 26, NAVY)
    k = 52; ox, oy = 470, 110                          # 主视图左上角
    fw, fh = X * k, Z2 * k
    b += poly([(ox, oy), (ox + fw, oy), (ox + fw, oy + fh), (ox, oy + fh)], VF, NAVY, 3)
    b += line((ox, oy + (Z2 - Z1) * k), (ox + fw, oy + (Z2 - Z1) * k), NAVY, 3)
    ty = oy + fh + 26                                  # 俯视图（前方在下）
    b += poly([(ox, ty), (ox + fw, ty), (ox + fw, ty + Y * k), (ox, ty + Y * k)], '#FBF3DC', GOLDD, 3)
    b += line((ox, ty + Y1 * k), (ox + fw, ty + Y1 * k), GOLDD, 3)
    lx = ox + fw + 26                                  # 左视图（前方在右）
    prof = [(0, 0), (Y, 0), (Y, Z1), (Y1, Z1), (Y1, Z2), (0, Z2)]
    b += poly([(lx + y * k, oy + (Z2 - z) * k) for y, z in prof], WF, TEAL, 3)
    if mini:
        b += line((315, 200), (430, 200), GOLD, 9, marker='ar') + line((430, 262), (315, 262), GOLD, 9, marker='ar')
        return crop(svg(W, H, b), 40, 85, 700, 330)
    b += text((ox + fw / 2, oy - 12), '主视图', 18, '#777', 400) + text((lx + Y * k / 2, oy - 12), '左视图', 18, '#777', 400)
    b += text((ox + fw + 18, ty + Y * k / 2 + 8), '俯视图', 18, '#777', 400, 'start')
    b += line((310, 205), (430, 205), GOLD, 5, marker='ar') + text((370, 190), '由物画图', 22, GOLDD)
    b += line((430, 255), (310, 255), GOLD, 5, marker='ar') + text((370, 288), '由图想物', 22, GOLDD)
    return svg(W, H, b)
save('convert', convert())
save('m_convert', convert(True))


# ---------- 组合体：表面连接关系（d2） ----------
def joints():
    W, H = 1000, 660
    b = text((W / 2, 46), '组合体表面连接关系与画法', 34)
    k = 34
    def fv(ox, oy, pts, extra=()):
        out = poly([(ox + x * k, oy - z * k) for x, z in pts], VF, NAVY, 3)
        for (x1, z1), (x2, z2) in extra:
            out += line((ox + x1 * k, oy - z1 * k), (ox + x2 * k, oy - z2 * k), NAVY, 3)
        return out
    # 平齐 / 不平齐：等轴测小图 + 主视图
    for i, (name, rule, yb) in enumerate((('平齐', '不画分界线', 2.0), ('不平齐', '画分界线', 1.2))):
        cx = 125 + i * 250
        P = iso(cx + 40, 210, 26)
        base = [((4, 0, 0), (4, 2, 0), (4, 2, 1), (4, 0, 1)), ((0, 2, 0), (4, 2, 0), (4, 2, 1), (0, 2, 1)), ((0, 0, 1), (4, 0, 1), (4, 2, 1), (0, 2, 1))]
        for f, c in zip(base, ('#AFC3DF', '#CAD8EE', '#E8EEF7')): b += poly([P(p) for p in f], c, NAVY, 2.5)
        top = [((3, 0, 1), (3, yb, 1), (3, yb, 2.2), (3, 0, 2.2)), ((1, yb, 1), (3, yb, 1), (3, yb, 2.2), (1, yb, 2.2)), ((1, 0, 2.2), (3, 0, 2.2), (3, yb, 2.2), (1, yb, 2.2))]
        for f, c in zip(top, ('#AFC3DF', '#CAD8EE', '#E8EEF7')): b += poly([P(p) for p in f], c, NAVY, 2.5)
        outline = [(0, 0), (4, 0), (4, 1), (3, 1), (3, 2.2), (1, 2.2), (1, 1), (0, 1)]
        b += fv(cx - 68, 470, outline, [] if name == '平齐' else [((1, 1), (3, 1))])
        b += text((cx, 520), '主视图', 20, '#777', 400)
        b += text((cx, 570), name, 30) + text((cx, 610), rule, 24, GOLDD)
    # 相切 / 相交：俯视图 + 主视图（圆柱 R=1，耳板厚 0.5）
    for i, (name, rule, hw) in enumerate((('相切', '不画切线', 1.0), ('相交', '画交线', 0.6))):
        cx = 625 + i * 250
        ox, oy = cx - 10, 230                    # 俯视图：圆心
        xi = -math.sqrt(1 - hw * hw)             # 板侧面与圆柱面的交点 x
        b += f'<circle cx="{ox}" cy="{oy}" r="{k}" fill="{VF}" stroke="{NAVY}" stroke-width="3"/>'
        b += line((ox - 3 * k, oy - hw * k), (ox + xi * k, oy - hw * k), NAVY, 3) + line((ox - 3 * k, oy + hw * k), (ox + xi * k, oy + hw * k), NAVY, 3)
        b += line((ox - 3 * k, oy - hw * k), (ox - 3 * k, oy + hw * k), NAVY, 3)
        b += text((cx, oy + 78), '俯视图', 20, '#777', 400)
        fy = 470                                 # 主视图：底面
        b += poly([(ox - k, fy), (ox + k, fy), (ox + k, fy - 1.6 * k), (ox - k, fy - 1.6 * k)], VF, NAVY, 3)
        b += poly([(ox - 3 * k, fy), (ox + xi * k, fy), (ox + xi * k, fy - 0.5 * k), (ox - 3 * k, fy - 0.5 * k)], VF, 'none', 0)
        b += line((ox - 3 * k, fy), (ox + xi * k, fy), NAVY, 3) + line((ox - 3 * k, fy - 0.5 * k), (ox + xi * k, fy - 0.5 * k), NAVY, 3)
        b += line((ox - 3 * k, fy), (ox - 3 * k, fy - 0.5 * k), NAVY, 3)
        b += line((ox - k, fy - 1.6 * k), (ox - k, fy - 0.5 * k), NAVY, 3)        # 圆柱左轮廓：板以上可见
        if name == '相交':
            b += line((ox + xi * k, fy), (ox + xi * k, fy - 0.5 * k), '#C0392B', 5)   # 交线
        else:
            b += f'<circle cx="{ox}" cy="{fy - 0.5 * k}" r="5" fill="#C0392B"/>'
        b += text((cx, 520), '主视图', 20, '#777', 400)
        b += text((cx, 570), name, 30) + text((cx, 610), rule, 24, GOLDD)
    return svg(W, H, b)
save('joints', joints())
print('ok')


# ---------- 截交线 / 相贯线（d3） ----------
def pl(pts, stroke=NAVY, sw=3, dash=None, fill='none'):
    d = f' stroke-dasharray="{dash}"' if dash else ''
    return (f'<polyline points="{" ".join(f"{x:.1f},{y:.1f}" for x, y in pts)}" fill="{fill}" stroke="{stroke}" '
            f'stroke-width="{sw}" stroke-linejoin="round"{d}/>')

def rng(a, b, n=60):
    return [math.radians(a + (b - a) * i / n) for i in range(n + 1)]

def cut_cylinder(mini=False):
    W, H = 640, 540
    s = 110; P = iso(310, 410, s)
    b = '' if mini else text((W / 2, 46), '截交线（平面截切圆柱）', 32)
    z0 = lambda t: 1.15 + 0.45 * math.cos(t)            # 截平面 z = 1.15 + 0.45x
    # 截平面（半透明）
    pc = [(x, y, 1.15 + 0.45 * x) for x, y in ((1.7, -1.5), (1.7, 1.5), (-1.7, 1.5), (-1.7, -1.5))]
    b += poly([P(p) for p in pc], '#D9CCE8', '#8E7CC3', 2.5, op=0.45)
    # 下部圆柱体：左轮廓 t=315° → 截交线前半 → 右轮廓 t=135° → 底面前半
    out = [P((math.cos(t), math.sin(t), 0)) for t in rng(135, -45)]
    out += [P((math.cos(t), math.sin(t), z0(t))) for t in rng(-45, 135)]
    b += poly(out, '#CAD8EE', NAVY, 3)
    # 截断面（实形为椭圆）
    sec = [P((math.cos(t), math.sin(t), z0(t))) for t in rng(0, 360, 90)]
    b += poly(sec, '#F4C7C0', '#C0392B', 5)
    # 被切去部分（虚线示意）
    top = [P((math.cos(t), math.sin(t), 2.3)) for t in rng(0, 360, 90)]
    b += pl(top, GRAY, 2.5, '8 6')
    for t in (math.radians(135), math.radians(-45)):
        b += line(P((math.cos(t), math.sin(t), z0(t))), P((math.cos(t), math.sin(t), 2.3)), GRAY, 2.5, '8 6')
    # 标注
    ex = P((math.cos(math.radians(60)), math.sin(math.radians(60)), z0(math.radians(60))))
    b += line(ex, (ex[0] + 90, ex[1] + 70), '#C0392B', 2.5) + text((ex[0] + 96, ex[1] + 96), '截交线', 30, '#C0392B', 700, 'start')
    fx = P((-0.1, 0.0, 1.15 - 0.045))
    b += line(fx, (fx[0] - 150, fx[1] - 120), NAVY, 2.5) + text((fx[0] - 150, fx[1] - 130), '截断面', 30, NAVY)
    px = P(pc[2])
    b += text((px[0] - 10, px[1] - 10), '截平面P', 28, '#6A55A8', 700, 'end')
    return crop(svg(W, H, b), 0, 60 if mini else 0, W, H - (60 if mini else 0))

def tee(mini=False):
    W, H = 640, 580
    s = 105; P = iso(320, 330, s)
    b = '' if mini else text((W / 2, 46), '相贯线（两圆柱正交）', 32)
    R, r, Lx, Hz = 1.0, 0.62, 1.9, 2.0
    # 横圆柱（轴线沿 x）
    body = [P((x, math.cos(math.radians(315)), math.sin(math.radians(315)))) for x in (-Lx, Lx)]
    body += [P((Lx, math.cos(u), math.sin(u))) for u in rng(315, 135)]
    body += [P((x, math.cos(math.radians(135)), math.sin(math.radians(135)))) for x in (Lx, -Lx)]
    body += [P((-Lx, math.cos(u), math.sin(u))) for u in rng(135, -45)]
    b += poly(body, '#CAD8EE', NAVY, 3)
    b += poly([P((Lx, math.cos(u), math.sin(u))) for u in rng(0, 360, 90)], '#AFC3DF', NAVY, 3)
    # 竖圆柱：相贯线前半 → 两条轮廓 → 顶面
    zc = lambda t: math.sqrt(R * R - (r * math.sin(t)) ** 2)
    up = [P((r * math.cos(t), r * math.sin(t), zc(t))) for t in rng(-45, 135)]
    up += [P((r * math.cos(t), r * math.sin(t), Hz)) for t in rng(135, -45)]
    b += poly(up, '#E8EEF7', NAVY, 3)
    b += poly([P((r * math.cos(t), r * math.sin(t), Hz)) for t in rng(0, 360, 90)], '#F3F6FB', NAVY, 3)
    curve = [P((r * math.cos(t), r * math.sin(t), zc(t))) for t in rng(-45, 135)]
    b += pl(curve, '#C0392B', 6)
    ex = P((r * math.cos(math.radians(80)), r * math.sin(math.radians(80)), zc(math.radians(80))))
    b += line(ex, (ex[0] + 110, ex[1] + 40), '#C0392B', 2.5) + text((ex[0] + 116, ex[1] + 66), '相贯线', 30, '#C0392B', 700, 'start')
    return crop(svg(W, H, b), 0, 60 if mini else 0, W, H - (60 if mini else 0))

save('cut_cyl', cut_cylinder()); save('tee', tee())
FS = 1.3; save('m_cut_cyl', cut_cylinder(True)); save('m_tee', tee(True)); FS = 1.0
