# 重画 3 张图：直角弯板 / 中心投影与正投影 / 课本搭建三投影面体系
src = open('/tmp/lo2/svg/draw.py', encoding='utf8').read().split('# ================= A')[0]
exec(src)

def pts(P, arr):
    return [P(p) for p in arr]

# ================= 1. 直角弯板 =================
def fig_wanban():
    s = 40; P = iso(560, 330, s)
    LB, W, T = 10.0, 4.0, 1.0          # 底板：长(x) 宽(y，+y 为前) 厚(z)
    TV, HV = 1.2, 5.0                  # 竖板：厚(x 0..TV) 高
    CY, CZ = 1.6, 2.0                  # 切角：顶面前缘切去 CY × CZ
    S0, S1, Y0, Y1 = LB - 2.0, LB, 1.3, 2.7   # 方槽：左端开口，贯通底板
    FILL_TOP, FILL_FRONT, FILL_SIDE, FILL_IN = '#E3ECF6', '#C6D8E8', '#AFC3DF', '#9DB3D3'
    b = ''
    # 竖板：顶面、朝左侧面、切角面、前面
    b += poly(pts(P, [(0, 0, HV), (TV, 0, HV), (TV, W - CY, HV), (0, W - CY, HV)]), FILL_TOP, NAVY, 3)
    b += poly(pts(P, [(TV, 0, T), (TV, W, T), (TV, W, HV - CZ), (TV, W - CY, HV), (TV, 0, HV)]), FILL_SIDE, NAVY, 3)
    b += poly(pts(P, [(0, W - CY, HV), (TV, W - CY, HV), (TV, W, HV - CZ), (0, W, HV - CZ)]), '#F4DE9C', GOLDD, 3.5)
    b += poly(pts(P, [(0, W, T), (TV, W, T), (TV, W, HV - CZ), (0, W, HV - CZ)]), FILL_FRONT, NAVY, 3)
    # 方槽内壁：x = S0 的端壁、y = Y0 的侧壁（朝前，可见）
    b += poly(pts(P, [(S0, Y0, 0), (S0, Y1, 0), (S0, Y1, T), (S0, Y0, T)]), FILL_IN, NAVY, 2.5)
    b += poly(pts(P, [(S0, Y0, 0), (S1, Y0, 0), (S1, Y0, T), (S0, Y0, T)]), FILL_IN, NAVY, 2.5)
    # 底板顶面（扣除方槽）
    b += poly(pts(P, [(TV, 0, T), (S1, 0, T), (S1, Y0, T), (S0, Y0, T), (S0, Y1, T), (S1, Y1, T), (S1, W, T), (TV, W, T)]),
              FILL_TOP, NAVY, 3)
    # 底板前面（与竖板前面同一平面）
    b += poly(pts(P, [(0, W, 0), (S1, W, 0), (S1, W, T), (0, W, T)]), FILL_FRONT, NAVY, 3)
    # 底板左端面（被方槽分成两块）
    b += poly(pts(P, [(S1, 0, 0), (S1, Y0, 0), (S1, Y0, T), (S1, 0, T)]), FILL_SIDE, NAVY, 3)
    b += poly(pts(P, [(S1, Y1, 0), (S1, W, 0), (S1, W, T), (S1, Y1, T)]), FILL_SIDE, NAVY, 3)
    # 标注
    c = P((TV / 2, W - CY / 2, HV - CZ / 2))
    b += line((c[0] + 150, c[1] - 110), (c[0] + 12, c[1] - 8), GOLD, 3, marker='ar') + text((c[0] + 190, c[1] - 120), '切角', 32, GOLDD)
    sl = P(((S0 + S1) / 2, (Y0 + Y1) / 2, T))
    b += line((sl[0] - 120, sl[1] - 120), (sl[0] - 8, sl[1] - 10), GOLD, 3, marker='ar') + text((sl[0] - 150, sl[1] - 130), '方槽', 32, GOLDD)
    fr = P((LB * 0.45, W, T / 2))
    b += line((fr[0] + 130, fr[1] + 110), (fr[0] + 14, fr[1] + 14), GOLD, 5, marker='ar') + text((fr[0] + 200, fr[1] + 150), '主视方向', 32)
    b += text((450, 60), '直角弯板', 42)
    return svg(900, 700, b)

# ================= 2. 中心投影与正投影 =================
def fig_touying():
    b = text((450, 60), '中心投影与正投影', 42)
    def panel(y0, title, central):
        out = text((60, y0 + 10), title, 32, NAVY, 700, 'start')
        px = 740                                  # 投影面位置
        out += line((px, y0 + 40), (px, y0 + 330), '#6B7A99', 6)
        ox = 470                                  # 物体位置
        oy1, oy2 = y0 + 130, y0 + 240
        out += line((ox, oy1), (ox, oy2), NAVY, 9)
        if central:
            S = (120, y0 + 185)
            k = (px - S[0]) / (ox - S[0])
            q1 = (px, S[1] + (oy1 - S[1]) * k); q2 = (px, S[1] + (oy2 - S[1]) * k)
            for yy in (oy1, (oy1 + oy2) / 2, oy2):
                qy = S[1] + (yy - S[1]) * k
                out += line(S, (px - 4, qy), GOLD, 3, marker='ar')
            out += f'<circle cx="{S[0]}" cy="{S[1]}" r="11" fill="{GOLD}"/>'
            out += text((S[0], S[1] + 50), '投射中心 S', 24, GOLDD)
            out += line(q1, q2, '#C0392B', 9)
            out += text((px + 60, (q1[1] + q2[1]) / 2 + 8), '投影', 26, '#C0392B', 700, 'start')
            out += text((600, y0 + 360), '投影比物体大：不反映真实大小', 24, '#555', 400)
        else:
            for yy in (y0 + 90, oy1, (oy1 + oy2) / 2, oy2):
                out += line((140, yy), (px, yy), TEAL, 3)
                out += f'<path d="M{300},{yy - 8} L{320},{yy} L{300},{yy + 8} z" fill="{TEAL}"/>'
            # 直角符号：画在最上面一条投射线与投影面的交点处（投射线下方、投影面左侧）
            ry = y0 + 90
            out += f'<path d="M{px - 22},{ry} L{px - 22},{ry + 22} L{px},{ry + 22}" fill="none" stroke="{NAVY}" stroke-width="3"/>'
            out += line((px, oy1), (px, oy2), '#C0392B', 9)
            out += text((px + 60, (oy1 + oy2) / 2 + 8), '投影', 26, '#C0392B', 700, 'start')
            out += text((600, y0 + 360), '投射线互相平行且垂直于投影面：反映真实大小', 24, '#555', 400)
            out += text((220, y0 + 75), '投射线', 24, TEAL)
        out += text((ox, oy2 + 50), '物体', 26)
        out += text((px, y0 + 30), '投影面', 26, '#6B7A99')
        return out
    b += panel(110, '中心投影法', True)
    b += line((40, 520), (860, 520), '#D5DEEA', 2)
    b += panel(560, '正投影法', False)
    return svg(900, 960, b)

# ================= 3. 课本搭建三投影面体系（按原照片重绘） =================
def fig_keben():
    s = 58; P = iso(470, 470, s)
    XM, YM, ZM = 6.0, 6.0, 5.4
    BOARD, EDGE, RED = '#F4F1EA', '#6B6456', '#C0392B'
    b = ''
    # 三块“课本”板：H 地面、V 左墙（y=0）、W 后墙（x=0）
    b += poly(pts(P, [(0, 0, 0), (XM, 0, 0), (XM, YM, 0), (0, YM, 0)]), BOARD, EDGE, 4)
    b += poly(pts(P, [(0, 0, 0), (XM, 0, 0), (XM, 0, ZM), (0, 0, ZM)]), '#ECE8DE', EDGE, 4)
    b += poly(pts(P, [(0, 0, 0), (0, YM, 0), (0, YM, ZM), (0, 0, ZM)]), '#F7F4EC', EDGE, 4)
    # 板厚（前缘加一条平行线，示意课本厚度）
    # 交线（装订处）加粗
    b += line(P((0, 0, 0)), P((0, 0, ZM)), '#3A352C', 5) + line(P((0, 0, 0)), P((XM, 0, 0)), '#3A352C', 4) + line(P((0, 0, 0)), P((0, YM, 0)), '#3A352C', 4)
    # 红色投影：V 面三角形、W 面矩形、H 面矩形（与原照片一致）
    b += poly(pts(P, [(4.6, 0, 1.0), (3.4, 0, 1.0), (4.3, 0, 4.2)]), 'none', RED, 6)
    b += poly(pts(P, [(0, 1.6, 2.2), (0, 3.6, 2.2), (0, 3.6, 4.6), (0, 1.6, 4.6)]), 'none', RED, 6)
    b += poly(pts(P, [(1.2, 2.2, 0), (3.0, 2.2, 0), (3.0, 4.6, 0), (1.2, 4.6, 0)]), 'none', RED, 6)
    # 字母
    b += text(P((XM - 0.7, 0, ZM - 0.9)), 'V', 46, RED)
    b += text(P((0, YM - 0.7, ZM - 0.9)), 'W', 46, RED)
    b += text(P((XM - 1.0, YM - 0.8, 0)), 'H', 46, RED)
    b += text((470, 60), '用课本搭建三投影面体系', 40)
    b += text((470, 870), '课本竖放搭出互相垂直的 V、H、W 三个面，在各面上画出物体的投影', 24, '#555', 400)
    return svg(940, 900, b)

for name, fn in (('wanban', fig_wanban), ('touying', fig_touying), ('keben', fig_keben)):
    open(f'/tmp/lo2/svg/re_{name}.svg', 'w', encoding='utf8').write(fn())
print('ok')
