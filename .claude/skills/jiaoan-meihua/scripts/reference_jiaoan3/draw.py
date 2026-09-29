# 重绘粗糙原图：统一扁平风格（藏蓝 / 青绿 / 金），几何按正投影关系计算
import math
NAVY, TEAL, GOLD, GOLDD = '#1F3864', '#2E7D7B', '#C9A227', '#A8841A'
VF, HF, WF = '#E8EEF7', '#FBF3DC', '#E3F1EF'
FONT = 'Noto Sans CJK SC, Microsoft YaHei, sans-serif'
C30, S30 = math.cos(math.radians(30)), 0.5

def svg(w, h, body):
    return (f'<svg xmlns="http://www.w3.org/2000/svg" width="{w}" height="{h}" viewBox="0 0 {w} {h}" '
            f'font-family="{FONT}"><rect width="{w}" height="{h}" fill="#fff"/>'
            '<defs><marker id="ar" markerUnits="userSpaceOnUse" markerWidth="16" markerHeight="16" refX="8" refY="8" orient="auto">'
            f'<path d="M0,0 L16,8 L0,16 z" fill="{GOLD}"/></marker>'
            '<marker id="an" markerUnits="userSpaceOnUse" markerWidth="14" markerHeight="14" refX="7" refY="7" orient="auto">'
            f'<path d="M0,0 L14,7 L0,14 z" fill="{NAVY}"/></marker></defs>{body}</svg>')

def poly(pts, fill='none', stroke=NAVY, sw=3, dash=None, op=1):
    d = f' stroke-dasharray="{dash}"' if dash else ''
    return (f'<polygon points="{" ".join(f"{x:.1f},{y:.1f}" for x, y in pts)}" fill="{fill}" fill-opacity="{op}" '
            f'stroke="{stroke}" stroke-width="{sw}" stroke-linejoin="round"{d}/>')

def line(a, b, stroke=NAVY, sw=3, dash=None, marker=None):
    d = f' stroke-dasharray="{dash}"' if dash else ''
    m = f' marker-end="url(#{marker})"' if marker else ''
    return f'<line x1="{a[0]:.1f}" y1="{a[1]:.1f}" x2="{b[0]:.1f}" y2="{b[1]:.1f}" stroke="{stroke}" stroke-width="{sw}"{d}{m}/>'

def text(p, s, size=26, fill=NAVY, w=700, anchor='middle'):
    return f'<text x="{p[0]:.1f}" y="{p[1]:.1f}" font-size="{size}" font-weight="{w}" fill="{fill}" text-anchor="{anchor}">{s}</text>'

# —— 形体：长方体左上方切去一个斜角（与原图 2-2 一致）——
L, WD, H = 4.0, 2.4, 3.0            # 长(x) 宽(y) 高(z)
def solid_faces(ox, oy, oz):
    """返回各可见面（等轴测下 +x/+y/+z 朝向观察者）；x 越大越靠“左”"""
    X = lambda x: ox + x; Y = lambda y: oy + y; Z = lambda z: oz + z
    front = [(X(0), Y(WD), Z(0)), (X(L), Y(WD), Z(0)), (X(L), Y(WD), Z(H / 2)), (X(L / 2), Y(WD), Z(H)), (X(0), Y(WD), Z(H))]
    top = [(X(0), Y(0), Z(H)), (X(L / 2), Y(0), Z(H)), (X(L / 2), Y(WD), Z(H)), (X(0), Y(WD), Z(H))]
    cham = [(X(L / 2), Y(0), Z(H)), (X(L), Y(0), Z(H / 2)), (X(L), Y(WD), Z(H / 2)), (X(L / 2), Y(WD), Z(H))]
    left = [(X(L), Y(0), Z(0)), (X(L), Y(WD), Z(0)), (X(L), Y(WD), Z(H / 2)), (X(L), Y(0), Z(H / 2))]
    return [(left, '#AFC3DF'), (top, '#E8EEF7'), (cham, '#F1DC9A'), (front, '#CAD8EE')]

def iso(cx, cy, s):
    return lambda p: (cx + (-p[0] + p[1]) * s * C30, cy + (p[0] + p[1]) * s * S30 - p[2] * s)

# ================= A：物体在三投影面体系中投射（替换原图 7、8） =================
def fig_A(with_object=True, arrows=False, title='三视图的形成'):
    s = 62; P = iso(470, 470, s)
    XM, YM, ZM = 7.2, 5.6, 5.4
    OX, OY, OZ = 1.6, 1.6, 1.2          # 物体位置
    b = ''
    Vp = [P((0, 0, 0)), P((XM, 0, 0)), P((XM, 0, ZM)), P((0, 0, ZM))]
    Hp = [P((0, 0, 0)), P((XM, 0, 0)), P((XM, YM, 0)), P((0, YM, 0))]
    Wp = [P((0, 0, 0)), P((0, YM, 0)), P((0, YM, ZM)), P((0, 0, ZM))]
    b += poly(Hp, HF, GOLDD, 3) + poly(Vp, VF, NAVY, 3) + poly(Wp, WF, TEAL, 3)
    # 三面投影
    fv = [(OX + 0, 0, OZ + 0), (OX + L, 0, OZ + 0), (OX + L, 0, OZ + H / 2), (OX + L / 2, 0, OZ + H), (OX + 0, 0, OZ + H)]
    tv = [(OX, OY, 0), (OX + L, OY, 0), (OX + L, OY + WD, 0), (OX, OY + WD, 0)]
    lv = [(0, OY, OZ), (0, OY + WD, OZ), (0, OY + WD, OZ + H), (0, OY, OZ + H)]
    b += poly([P(p) for p in fv], '#CAD8EE', NAVY, 4)
    b += poly([P(p) for p in tv], '#F6E6B4', GOLDD, 4) + line(P((OX + L / 2, OY, 0)), P((OX + L / 2, OY + WD, 0)), GOLDD, 3)
    b += poly([P(p) for p in lv], '#C9E4E0', TEAL, 4) + line(P((0, OY, OZ + H / 2)), P((0, OY + WD, OZ + H / 2)), TEAL, 3)
    if with_object:
        # 投射线（虚线）：选取特征点
        for a, c in (((OX + L / 2, OY, OZ + H), (OX + L / 2, 0, OZ + H)), ((OX + L, OY, OZ + H / 2), (OX + L, 0, OZ + H / 2)),
                     ((OX + L, OY + WD, OZ), (OX + L, OY + WD, 0)), ((OX, OY + WD, OZ), (OX, OY + WD, 0)),
                     ((OX, OY + WD, OZ + H), (0, OY + WD, OZ + H)), ((OX, OY + WD, OZ), (0, OY + WD, OZ))):
            b += line(P(a), P(c), '#8A99B3', 2, '7 6')
        for face, fill in solid_faces(OX, OY, OZ):
            b += poly([P(p) for p in face], fill, NAVY, 3)
    # 轴与标注
    b += text(P((XM + 0.35, 0, -0.1)), 'X', 30) + text(P((0, 0, ZM + 0.25)), 'Z', 30) + text(P((0, YM + 0.35, -0.1)), 'Y', 30)
    if not with_object: b += text((P((0, 0, 0))[0] + 22, P((0, 0, 0))[1] + 10), 'O', 28)
    b += text(P((XM - 0.5, 0, ZM - 0.6)), 'V', 38) + text(P((XM - 0.6, YM - 0.6, 0)), 'H', 38, GOLDD) + text(P((0, YM - 0.5, ZM - 0.6)), 'W', 38, TEAL)
    b += text(P((OX + L / 2, 0, OZ + H + 0.7)), '主视图', 26) + text(P((OX + L / 2 + 0.6, OY + WD + 1.1, 0)), '俯视图', 26, GOLDD)
    b += text(P((0, OY + WD / 2, OZ + H + 0.7)), '左视图', 26, TEAL)
    if with_object:
        # 投射方向
        f0 = P((OX + L / 2, OY + WD + 2.2, OZ + H / 2)); f1 = P((OX + L / 2, OY + WD + 0.5, OZ + H / 2))
        b += line(f0, f1, GOLD, 5, marker='ar') + text((f0[0] + 10, f0[1] + 30), '由前向后', 24, GOLDD)
        t0 = P((OX + L * 0.25, OY + WD / 2, OZ + H + 2.0)); t1 = P((OX + L * 0.25, OY + WD / 2, OZ + H + 0.4))
        b += line(t0, t1, GOLD, 5, marker='ar') + text((t0[0], t0[1] - 12), '由上向下', 24, GOLDD)
        l0 = P((OX + L + 2.3, OY + WD / 2, OZ + H * 0.3)); l1 = P((OX + L + 0.5, OY + WD / 2, OZ + H * 0.3))
        b += line(l0, l1, GOLD, 5, marker='ar') + text((l0[0] - 8, l0[1] + 32), '由左向右', 24, GOLDD)
    if arrows:
        # H 面绕 OX 向下转 90°，W 面绕 OZ 向右转 90°
        h0 = P((XM * 0.55, YM, 0)); b += f'<path d="M{h0[0]:.1f},{h0[1]:.1f} q30,70 -10,120" fill="none" stroke="{GOLD}" stroke-width="5" marker-end="url(#ar)"/>'
        b += text((h0[0] + 70, h0[1] + 90), 'H 面向下转 90°', 24, GOLDD)
        w0 = P((0, YM, ZM * 0.55)); b += f'<path d="M{w0[0]:.1f},{w0[1]:.1f} q70,-10 110,-70" fill="none" stroke="{GOLD}" stroke-width="5" marker-end="url(#ar)"/>'
        b += text((w0[0] + 40, w0[1] + 50), 'W 面向右转 90°', 24, GOLDD)
    b += text((470, 60), title, 38)
    return svg(940, 900, b)

# ================= C：三视图（替换原图 9，图2-2c） =================
def views(ox, oy, s, labels=True, frame=False):
    b = ''
    fx, fy = ox, oy                         # 主视图左上角
    fw, fh = L * s, H * s
    b += poly([(fx, fy + fh), (fx + fw, fy + fh), (fx + fw, fy), (fx + fw / 2, fy), (fx, fy + fh / 2)], '#CAD8EE', NAVY, 4)
    ty = fy + fh + 0.9 * s                  # 俯视图：长对正
    th = WD * s
    b += poly([(fx, ty), (fx + fw, ty), (fx + fw, ty + th), (fx, ty + th)], '#F6E6B4', GOLDD, 4) + line((fx + fw / 2, ty), (fx + fw / 2, ty + th), GOLDD, 3)
    lx = fx + fw + 0.9 * s                  # 左视图：高平齐
    b += poly([(lx, fy), (lx + th, fy), (lx + th, fy + fh), (lx, fy + fh)], '#C9E4E0', TEAL, 4) + line((lx, fy + fh / 2), (lx + th, fy + fh / 2), TEAL, 3)
    if labels:
        b += text((fx + fw / 2, fy + fh + 40), '主视图', 26) + text((fx + fw / 2, ty + th + 40), '俯视图', 26, GOLDD)
        b += text((lx + th / 2, fy + fh + 40), '左视图', 26, TEAL)
    return b, (fx, fy, fw, fh, ty, th, lx)

def fig_C():
    b, _ = views(120, 120, 95)
    b += text((430, 60), '展开后的三视图', 38)
    return svg(860, 880, b)

# ================= D：投影面展开（替换原图 16） =================
def fig_D():
    s = 70; ox, oy = 470, 430            # O 点
    b = ''
    b += poly([(80, 110), (ox, 110), (ox, oy), (80, oy)], VF, NAVY, 3)
    b += poly([(ox, 110), (820, 110), (820, oy), (ox, oy)], WF, TEAL, 3)
    b += poly([(80, oy), (ox, oy), (ox, 800), (80, 800)], HF, GOLDD, 3)
    b += line((50, oy), (850, oy), NAVY, 3) + line((ox, 80), (ox, 830), NAVY, 3)
    b += text((60, oy - 12), 'X', 30) + text((ox + 20, 96), 'Z', 30) + text((ox - 26, oy + 34), 'O', 28)
    b += text((840, oy + 36), 'Yw', 28, TEAL) + text((ox + 30, 830), 'Yh', 28, GOLDD)
    b += text((115, 150), 'V', 38) + text((785, 150), 'W', 38, TEAL) + text((115, 780), 'H', 38, GOLDD)
    # 三视图：主视图贴近 OX、OZ
    fw, fh, th = L * s, H * s, WD * s
    fx, fy = ox - 60 - fw, oy - 50 - fh
    b += poly([(fx, fy + fh), (fx + fw, fy + fh), (fx + fw, fy), (fx + fw / 2, fy), (fx, fy + fh / 2)], '#CAD8EE', NAVY, 4)
    ty = oy + 50
    b += poly([(fx, ty), (fx + fw, ty), (fx + fw, ty + th), (fx, ty + th)], '#F6E6B4', GOLDD, 4) + line((fx + fw / 2, ty), (fx + fw / 2, ty + th), GOLDD, 3)
    lx = ox + 50
    b += poly([(lx, fy), (lx + th, fy), (lx + th, fy + fh), (lx, fy + fh)], '#C9E4E0', TEAL, 4) + line((lx, fy + fh / 2), (lx + th, fy + fh / 2), TEAL, 3)
    # OY 轴一分为二的说明
    b += f'<path d="M{ox + 120},{oy + 40} A110,110 0 0 1 {ox + 40},{oy + 120}" fill="none" stroke="{GOLD}" stroke-width="4" stroke-dasharray="8 6"/>'
    b += text((ox + 190, oy + 150), 'OY 一分为二', 24, GOLDD)
    b += text((450, 60), '投影面展开：V 不动，H、W 旋转', 34)
    return svg(900, 880, b)

# ================= E：正投影与斜投影（替换原图 12） =================
def fig_E():
    b = ''
    def panel(cx, title, slant):
        out = ''
        base = [(cx - 190, 560), (cx + 110, 560), (cx + 190, 470), (cx - 110, 470)]
        out += poly(base, '#E3F1EF', TEAL, 3)
        tri_proj = [(cx - 90, 540), (cx + 70, 530), (cx + 10, 480)]
        dx = 90 if slant else 0
        tri_obj = [(x + dx, y - 280) for x, y in tri_proj]
        for a, c in zip(tri_obj, tri_proj):
            out += line((a[0] + (a[0] - c[0]) * 0.45, a[1] - 120), c, '#8A99B3', 2, '7 6')
        out += poly(tri_proj, '#F6E6B4', GOLDD, 4)
        out += poly(tri_obj, '#CAD8EE', NAVY, 4)
        out += text((cx, 110), title, 32)
        return out
    b += panel(250, '正投影法', False) + panel(730, '斜投影法', True)
    b += text((250, 640), '投射线 ⊥ 投影面', 26, TEAL, 400) + text((730, 640), '投射线倾斜于投影面', 26, TEAL, 400)
    b += text((250, 690), '度量性好，机械图样采用', 26, GOLDD) + text((730, 690), '常用于绘制斜轴测图', 26, GOLDD)
    return svg(980, 730, b)

# ================= F：单面投影不能确定形状（替换原图 14） =================
def fig_F():
    s = 46; b = ''
    b += text((560, 56), '一个视图不能唯一确定物体形状', 34)
    arc = [(2 * math.cos(math.radians(t)), 2.2 * math.sin(math.radians(t))) for t in range(0, 91, 6)]
    profiles = [
        ('长方体', [(0, 0), (2, 0), (2, 2.2), (0, 2.2)]),
        ('斜面体', [(0, 0), (2, 0), (0, 2.2)]),
        ('四分之一圆柱', [(0, 0)] + arc),
        ('阶梯体', [(0, 0), (2, 0), (2, 1.1), (1, 1.1), (1, 2.2), (0, 2.2)]),
    ]
    view = (1, 1, 1)
    Lx = 2.4
    for k, (name, prof) in enumerate(profiles):
        cx = 250 + k * 240
        P = iso(cx, 400, s)
        n = len(prof)
        cy_ = sum(p[0] for p in prof) / n; cz_ = sum(p[1] for p in prof) / n
        center = (Lx / 2, cy_, cz_)
        faces = [([(0, y, z) for y, z in prof], False), ([(Lx, y, z) for y, z in prof], False)]
        for i in range(n):
            j = (i + 1) % n
            (y1, z1), (y2, z2) = prof[i], prof[j]
            curved = name == '四分之一圆柱' and 1 <= i < n - 1
            faces.append(([(0, y1, z1), (0, y2, z2), (Lx, y2, z2), (Lx, y1, z1)], curved))
        drawn = []
        for pts, curved in faces:
            # Newell 法向量
            nx = ny = nz = 0.0
            for i in range(len(pts)):
                x1, y1, z1 = pts[i]; x2, y2, z2 = pts[(i + 1) % len(pts)]
                nx += (y1 - y2) * (z1 + z2); ny += (z1 - z2) * (x1 + x2); nz += (x1 - x2) * (y1 + y2)
            c = [sum(p[t] for p in pts) / len(pts) for t in range(3)]
            out = [c[t] - center[t] for t in range(3)]
            if nx * out[0] + ny * out[1] + nz * out[2] < 0: nx, ny, nz = -nx, -ny, -nz
            if nx * view[0] + ny * view[1] + nz * view[2] <= 1e-9: continue
            shade = '#AFC3DF' if abs(nx) > max(abs(ny), abs(nz)) else ('#E8EEF7' if nz > abs(ny) else '#CAD8EE')
            drawn.append((sum(c), pts, curved, shade))
        drawn.sort(key=lambda d: d[0])
        for _, pts, curved, shade in drawn:
            b += poly([P(p) for p in pts], shade, shade if curved else NAVY, 1.2 if curved else 2.5)
        if name == '四分之一圆柱':
            b += line(P((0, 0, 2.2)), P((Lx, 0, 2.2)), NAVY, 2.5) + line(P((0, 2, 0)), P((Lx, 2, 0)), NAVY, 2.5)
            for xx in (0, Lx):
                pl = [P((xx, y, z)) for y, z in arc]
                b += '<polyline points="' + ' '.join(f"{x:.1f},{y:.1f}" for x, y in pl) + f'" fill="none" stroke="{NAVY}" stroke-width="2.5"/>'
        b += text((cx, 610), name, 24, NAVY, 400)
        rx, ry = cx - 55, 150
        b += poly([(rx, ry), (rx + 110, ry), (rx + 110, ry + 100), (rx, ry + 100)], '#F6E6B4', GOLDD, 3)
    b += text((40, 208), '主视图', 26, GOLDD, 700, 'start') + text((40, 440), '形体', 26, NAVY, 700, 'start')
    b += text((560, 665), '四个形体的主视图完全相同，需要多面视图配合表达', 26, '#555', 400)
    return svg(1120, 700, b)

files = {
    'fig_A.svg': fig_A(True, False, '物体在三投影面体系中的投影'),
    'fig_B.svg': fig_A(False, True, '三面投影与投影面展开'),
    'fig_C.svg': fig_C(),
    'fig_D.svg': fig_D(),
    'fig_E.svg': fig_E(),
    'fig_F.svg': fig_F(),
}
for n, c in files.items():
    open(f'/tmp/lo2/svg/{n}', 'w').write(c)
print('ok')
