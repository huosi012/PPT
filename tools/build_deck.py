# -*- coding: utf-8 -*-
"""根据《项目汇报PPT文本.md》生成课题申报汇报 PPT（版式风格复刻参考页）。

用法：python3 tools/build_deck.py [输出路径]
依赖：python-pptx、Pillow；图标首次生成需 node + tools/package.json 中的依赖（已生成的 PNG 在 assets/icons）。
"""
import os
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from pptkit import *  # noqa: E402,F401,F403
import pptkit  # noqa: E402

OUT = os.path.join(ROOT, '项目汇报PPT.pptx')
LOCO = os.path.join(ROOT, 'assets', 'illustrations', LOCO_NAME + '.png')
LOCO_FADED = os.path.join(ROOT, 'assets', 'illustrations', LOCO_NAME + '_faded.png')
TOTAL = 25
FOOTER = '7.3 多智能体协同运行与调度引擎'
BAR = STYLE['top_bar']  # 卡片顶线 / 左侧色条粗细（随配色方案变化，现版 6px）


# 经费明细（第 3、21 页共用）：序号、科目、金额、测算依据、色块、条内标签
BUDGET = [
    ('1', '直接投入费用', 50, 'AI算力服务器（GPU+NPU异构）1台 40万；仿真与调度一体化工作站 2台 10万', M1.main, '直接投入 50'),
    ('2', '人员人工费用', 30, '项目参研人员薪酬费用，按投入人月测算：<br><bc=%s>【待填】人·月 × 【待填】万元/人·月</bc>' % PH.text,
     BUDGET2, '人员 30'),
    ('3', '固定资产相关费用（折旧）', 5, '现有仿真与测试设备折旧分摊', GRAYS[0], None),
    ('4', '试验检验及试制外协费用', 5, '建模数据采集处理 2万；仿真模型校验与第三方验证 3万', GRAYS[1], None),
    ('5', '研发成果相关费用', 5, '发明专利申请 2万；论文版面及标准草案编制 3万', GRAYS[2], None),
    ('6', '与研发活动直接相关的<br>其他费用', 5, '差旅费 3万；会议费 2万', GRAYS[3], None),
]

# 研究阶段（第 17、22 页共用）：名称、简称、起止月（0—24）、色、起止时间、阶段目标、阶段交付、里程碑、日期、里程碑名
STAGES = [
    ('阶段一', '架构与规范', 0, 6, STAGE[0], '2027.01—2027.06', '需求分析与总体架构设计', '总体架构、验证判据知识化方法与标准化接口规范',
     'M1', '2027.06', '架构与接口规范'),
    ('阶段二', '智能体与引擎', 6, 15, STAGE[1], '2027.07—2028.03', '智能体与引擎研制', '关键系统仿真智能体／智能孪生模型、调度引擎原理样机',
     'M2', '2028.03', '智能体与引擎'),
    ('阶段三', '样机与验证', 15, 21, STAGE[2], '2028.04—2028.09', '样机集成与验证', '列车级数字样机、典型工况协同验证结果', 'M3',
     '2028.09', '数字样机与验证'),
    ('阶段四', '示范与验收', 21, 24, STAGE[3], '2028.10—2028.12', '示范应用与验收', '示范应用与效能评估、方法体系与技术规范、标准草案',
     'M4', '2028.12', '示范应用与验收'),
]


def target_chip(s, x, y, label, value, size=17, h=34, padx=12):
    """量化目标占位：虚线框 + 指标名 + 目标值（【待填】）。返回宽度。"""
    w = text_width(label + '　', size) + text_width(value, size, True) + padx * 2
    shp = rect(s, x, y, w, h, fill=PH.bg, line=PH.main, lw=1.2, dash='dash')
    text(s, x, y, w, h, '%s　<bc=%s>%s</bc>' % (label, PH.text, value), size=size, color=BODY, align='c', anchor='m',
         lh=round(size * 1.3), wrap=False, shp=shp)
    return w


def loco_train(s, x, y_rail, w_each, n=3, gap=10, faded=True):
    """在标尺线（轨道）上摆放 n 节机车插画。y_rail 为轨面位置。"""
    h = w_each * 0.25
    for i in range(n):
        pic(s, LOCO_FADED if faded else LOCO, x + i * (w_each + gap), y_rail - h * 0.992, w_each, h)


# =====================================================================  01 封面
def s01_cover(d):
    s = d.new_slide()
    rect(s, 64, 34, 1792, STYLE['rule_w'], fill=RULE)
    rect(s, 64, 132, 8, 30, fill=RED)
    text(s, 88, 128, 1400, 38, '****“***”科技重大专项课题申报', size=25, bold=True, color=RED, cs=3, anchor='m',
         lh=38, wrap=False)
    text(s, 64, 196, 1792, 230, ['<r>7.3</r> 轨道交通装备性能验证的', '多智能体协同运行与调度引擎技术研究'],
         size=78, bold=True, color=INK, lh=110)
    rect(s, 64, 446, STYLE['title_bar'][0] * 1.25, STYLE['title_bar'][1] + 2, fill=RED)
    text(s, 64, 480, 1792, 46, '所属项目：7. 基于智能体的轨道交通装备性能数字样机关键技术研究', size=29, color=SUB,
         anchor='m', lh=46, wrap=False)
    cells = [('牵头单位', '**********'), ('项目负责人', '***'), ('申报层级', '课题级'),
             ('研究周期', '2027年1月—2028年12月')]
    ws = [330, 270, 250, 520]
    x = 64
    for i, ((lab, val), w) in enumerate(zip(cells, ws)):
        xx = x + (30 if i else 0)
        if i:
            rect(s, x, 598, 2, 88, fill=LINE)
        text(s, xx, 594, w - 30, 32, lab, size=20, bold=True, color=MUTED, lh=32, cs=1, wrap=False)
        text(s, xx, 632, w - 30, 54, val, size=32, bold=True, color=INK, lh=54, wrap=False)
        x += w
    # 机车 + 标尺（轨道）
    ruler(s, 64, 944, 1792)
    loco_train(s, 560, 944, 428, n=3, gap=8, faded=False)
    rect(s, 64, 1010, 1792, STYLE['rule_w'], fill=RULE)
    sep = '<c=%s>   |   </c>' % RULE
    text(s, 64, 1024, 1300, 36, FOOTER + sep + '课题申报汇报' + sep + '研究周期 2027.01—2028.12', size=17,
         color=MUTED, bold=True, cs=2.5, anchor='m', lh=36, wrap=False)


# =====================================================================  02 提纲
def s02_agenda(d):
    s = d.new_slide()
    frame(d, s, '汇报提纲：七个部分，核心是研究内容、技术路线与方法', '共 25 页 · 第三部分“主要研究内容、技术路线与方法”为核心章节（第 7—17 页）',
          '汇报提纲')
    rows = [
        ('01', '项目简介', '项目基本信息 · 项目子课题与本课题边界 · 预计成果物 · 预计经费', '第 3 页', 'file-description'),
        ('02', '研究意义与价值', '研究背景与重要性 · 研究目的 · 应用价值与推动作用 · 社会经济效益 ｜ 立项必要性与紧迫性 · 未来价值',
         '第 4—6 页', 'bulb'),
        ('03', '主要研究内容、技术路线与方法',
         '核心命题 · 研究内容①—④ · 研究方法 · 技术路线 · 典型验证场景 · 创新点 · 技术对标 · 阶段目标与进度',
         '第 7—17 页', 'route'),
        ('04', '预期成果与效益', '子任务成果与考核方式 · 预期目标与成果对应 · 量化指标与效益评估', '第 18—20 页', 'target-arrow'),
        ('05', '经费预算合理性', '预算明细与测算依据 · 分年度 · 合理性说明', '第 21 页', 'coins'),
        ('06', '项目管理计划', '任务分配 · 项目流程 · 时间安排 · 执行保障 ｜ 团队研究能力', '第 22—23 页', 'sitemap'),
        ('07', '成果介绍', '五项成果 ＋ 知识产权与标准', '第 24 页', 'trophy'),
    ]
    y0, rh = 222, 110
    for i, (no, title, sub, pages, ic) in enumerate(rows):
        y = y0 + i * rh
        core = (no == '03')
        if core:
            rect(s, 64, y + 4, 1792, rh - 8, fill=RED_BG)
            rect(s, 64, y + 4, STYLE['concl_bar'], rh - 8, fill=RED)
        text(s, 92, y, 110, rh, no, size=46, bold=True, color=RED if core else AGENDA_NUM, font=MONO, anchor='m',
             lh=58, wrap=False)
        oval(s, 208, y + (rh - 64) / 2.0, 64, 64, fill=WHITE if core else GRAYBG, line=None)
        icon(s, ic, RED if core else SUB, 208 + 15, y + (rh - 64) / 2.0 + 15, 34)
        text(s, 300, y + 16, 1200, 44, title, size=30, bold=True, color=INK, anchor='m', lh=44, wrap=False)
        text(s, 300, y + 60, 1330, 34, sub, size=20, color=SUB, anchor='m', lh=34, wrap=False)
        if core:
            tag(s, 300 + text_width(title, 30, True) + 18, y + 22, '核心章节', fill=RED, size=18, h=32)
        text(s, 1600, y + 58, 256, 36, pages, size=21, bold=True, color=RED if core else MUTED, align='r',
             anchor='m', lh=36, wrap=False)
        if i < len(rows) - 1 and not core and rows[i + 1][0] != '03':
            line(s, 300, y + rh, 1856, y + rh, LINE, 1)


# =====================================================================  03 项目简介
def s03_intro(d):
    s = d.new_slide()
    frame(d, s, '项目简介：研究重心是引擎，数字样机是引擎的集成产物', '项目基本信息 · 项目子课题与本课题边界 · 预计成果物 · 预计经费', '项目简介')
    # ---- 左：基本信息
    L, LW = 64, 840
    header(s, L, 212, '项目基本信息', w=LW)
    lab = dict(fill=GRAYBG, bold=True, color=SUB, size=17)
    val = dict(color=INK, size=19)
    yy = table(s, L, 256, [96, 278, 112, 354], [
        ['项目名称', '7. 基于智能体的轨道交通装备<br>性能数字样机关键技术研究', '课题名称',
         '<b>7.3 轨道交通装备性能验证的</b><br><b>多智能体协同运行与调度引擎技术研究</b>'],
        ['牵头单位', '**********', '项目负责人', '***'],
        ['主管部门', '********部（***）', '产业领域', '轨道交通装备'],
        ['项目目的', '基础前瞻共性技术研究', '研究周期', '2027.01—2028.12（24个月）'],
        ['总预算', '<r>100万元</r>', '当年预算', '50万元'],
    ], size=19, lh=27, pad_x=10, pad_y=8, col_styles={0: lab, 1: val, 2: lab, 3: val})
    # ---- 左：主要成果物（3×2）
    y1 = yy + 22
    header(s, L, y1, '预计成果物', note='软件 · 模型 · 系统 · 规范 · 报告 · 知识产权', w=LW)
    items = [
        ('多智能体协同运行与调度引擎', '软件1套', M1, 'settings-automation'),
        ('关键系统仿真智能体／<br>智能孪生模型', '模型1套', M2, 'robot'),
        ('列车级数字样机', '系统1套', M3, 'train'),
        ('列车级数字样机构建方法体系与技术规范', '规范1套', M3, 'file-certificate'),
        ('示范应用与验证报告', '报告1份', M4, 'report-analytics'),
        ('发明专利／论文／<br>企业标准草案', '4项／2篇／2项', M4, 'certificate'),
    ]
    cw, ch, gx, gy = (LW - 24) / 3.0, 116, 12, 12
    for i, (name, form, th, ic) in enumerate(items):
        cx = L + (i % 3) * (cw + gx)
        cy = y1 + 46 + (i // 3) * (ch + gy)
        rect(s, cx, cy, cw, ch, fill=WHITE, line=LINE)
        rect(s, cx, cy, cw, min(4, BAR), fill=th.main)
        icon(s, ic, th.main, cx + 16, cy + 18, 30)
        tag(s, cx + cw - 16 - (text_width(form, 16, True) + 20), cy + 18, form, fill=th.bg, color=th.text, size=16,
            h=28, padx=10)
        text(s, cx + 16, cy + 56, cw - 28, 52, name, size=18, bold=True, color=INK, lh=25)
    # ---- 左：经费概览（堆叠条）
    y2 = y1 + 46 + 2 * ch + gy + 22
    header(s, L, y2, '预计经费', note='总预算 100 万元', w=LW)
    by = y2 + 48
    stacked_bar(s, L, by, LW, 40, [(v, c, lab, WHITE) for _, _, v, _, c, lab in BUDGET])
    unit = (LW - 2 * 5) / 100.0
    text(s, L, by + 46, 50 * unit, 28, 'GPU+NPU 异构算力服务器 40 · 一体化工作站 10', size=16, color=SUB, lh=28,
         wrap=False)
    text(s, L + 440, by + 46, LW - 440, 28, '人员 30 · 折旧 · 外协 · 成果 · 其他 各 5', size=16, color=SUB,
         align='r', lh=28, wrap=False)

    # ---- 右：在项目7中的位置与边界
    R, RW = 944, 912
    header(s, R, 212, '项目子课题与本课题边界', note='项目7共设5个子课题 · 本课题研究重心是“引擎”', w=RW)
    rows = [
        ('7.1', '高质量数据集', '边界：仅采用其数据规范与接口，不承担构建（<b>7.1 建库、7.3 按需生成</b>）', 'n'),
        ('7.2', '多物理域建模与仿真智能体', '边界：为验证需要自建控制系统、车辆动力学与线路等必要智能体，<b>不覆盖弓网/牵引/制动等全部关键系统</b>',
         'n'),
        ('7.3', '协同运行与调度引擎（本课题）',
         '<b>核心攻关多智能体协同运行与调度引擎</b>，<br>解决多专业智能体“接得上、跑得动、调得灵”的问题；<br>'
         '<b>列车级数字样机是引擎承载智能体后形成的集成产物</b>；<br>'
         '本课题分解为<b> 4 个子任务</b>：知识 → 调度 → 执行 → 集成（详见第 18 页）', 'me'),
        ('7.4', '数字样机构建', '边界：承接智能体集成与协同运行', 'n'),
        ('7.5', '方案优化与智能生成', '边界：仅为优化提供验证支撑', 'n'),
        ('6.2', '通用智能体调度平台', '边界：本课题为<b>面向性能验证任务的领域化编排</b>，强调实时协同与多物理域时序', 'ext'),
    ]
    nb, tgap = 76, 20
    tw_ = RW - nb - tgap - 22
    ext_gap = 20
    groups = [34 + 8 + wrap_lines(desc, tw_, 19) * 29 for _, _, desc, _ in rows]
    pad = (996 - 256 - 10 * (len(rows) - 1) - ext_gap - sum(groups)) / float(len(rows))
    y = 256
    for (no, name, desc, kind), g in zip(rows, groups):
        h = g + pad
        if kind == 'ext':
            line(s, R, y + 5, R + RW, y + 5, RULE, 1.5, dash='dash')
            y += ext_gap
        me = (kind == 'me')
        rect(s, R, y, RW, h, fill=RED_BG if me else WHITE, line=RED if me else LINE, lw=2 if me else 1,
             dash='dash' if kind == 'ext' else None)
        shp = rect(s, R, y, nb, h, fill=RED if me else (PANEL if kind == 'ext' else GRAYBG))
        text(s, R, y, nb, h, no, size=26, bold=True, color=WHITE if me else INK, font=MONO, align='c', anchor='m',
             lh=34, shp=shp, wrap=False)
        tx = R + nb + tgap
        top = y + (h - g) / 2.0
        text(s, tx, top, 640, 34, name, size=22, bold=True, color=RED if me else INK, lh=34, wrap=False)
        if me:
            tag(s, tx + text_width(name, 22, True) + 14, top + 3, '本课题', fill=RED, size=16, h=28, padx=10)
        if kind == 'ext':
            tag(s, tx + text_width(name, 22, True) + 14, top + 3, '相邻项目', fill=GRAYBG, color=SUB, size=16, h=28,
                padx=10)
        text(s, tx, top + 42, tw_, g - 40, desc, size=19, color=BODY if me else SUB, lh=29)
        y += h + 10


# =====================================================================  04 研究意义与价值（闭环页版式）
def s04_value(d):
    s = d.new_slide()
    frame(d, s, '研究意义与价值：实现列车关键性能的快速、可复现、全覆盖验证', '研究背景与重要性 → 研究目的 → 应用价值与推动作用 · 潜在社会经济效益', '研究意义与价值')
    X, XW = 300, 1556
    side_label(s, 64, 222, 206, '研究背景与重要性', '实物试验为主、模型各自为战，难以支撑列车级性能验证')
    cw = (XW - 2 * 24) / 3.0
    cards = [
        ('实物试验＋离线仿真为主', '轨道交通装备研发验证仍以实物试验＋单系统离线仿真为主，周期长、成本高，<b>极端工况与故障场景难以全覆盖</b>'),
        ('求解慢 · 精度低 · 靠经验', '现有数字样机与孪生模型存在<b>运算求解慢、极端工况精度低、经验依赖性强</b>等突出问题'),
        ('各专业模型各自为战', '各专业仿真模型各自为战，<b>缺乏统一的协同运行与调度能力</b>，难以支撑列车级性能验证'),
    ]
    for i, (t, b) in enumerate(cards):
        problem_card(s, X + i * (cw + 24), 214, cw, 176, t, b, body_size=20)
    # 研究目的
    side_label(s, 64, 424, 206, '研究目的')
    rect(s, X, 414, XW, 88, fill=RED_BG)
    rect(s, X, 414, STYLE['concl_bar'], 88, fill=RED)
    text(s, X + 30, 414, XW - 60, 88,
         '突破多智能体协同运行与调度引擎技术，构建包含关键系统仿真智能体的列车级数字样机，实现列车关键性能的<r>快速、可复现、全覆盖</r>验证。',
         size=23, bold=True, color=INK, anchor='m', lh=36)
    # 价值与效益
    side_label(s, 64, 540, 206, '价值与效益', '应用价值与推动作用<br>潜在社会经济效益')
    y0, ph = 528, 468
    pw = (XW - 32) / 2.0
    yy = panel(s, X, y0, pw, ph, M1, '应用价值与对领域发展的推动作用', '01', 'trending-up', title_size=28, badge=76)
    bullets(s, X + 40, yy + 20, pw - 80, 210, [
        '应用于<b>机车研发验证与运维测试</b>，覆盖速度跟踪控制品质、牵引能耗与再生能量利用、运行平稳性等关键性能；支撑复杂工况的<b>低成本、高频次</b>验证',
        '推动研发验证从<b>单系统、离线</b>向<b>列车级、实时协同、智能调度</b>演进；将多智能体协同从通用平台层引入<b>装备性能验证</b>领域',
    ], M1.main, size=20, lh=31, sa=12)
    concl(s, X + 40, y0 + ph - 32 - 72, pw - 80, 72, '输出列车级数字样机构建方法体系与技术规范，<br>为行业提供可复用的方法支撑', M1,
          size=20)
    x2 = X + pw + 32
    yy = panel(s, x2, y0, pw, ph, M3, '潜在社会经济效益', '02', 'coins', title_size=30, badge=76)
    effects = [
        ('经济效益', 'report-money', '减少实物试验次数与样车试制投入，缩短研制迭代周期，降低运用考核与线路试验成本'),
        ('社会效益', 'shield-check', '提升关键性能验证的覆盖度与可信度，支撑轨道交通安全可靠运营'),
        ('产业效益', 'building-factory-2', '方法体系具备向<b>制动、辅助系统</b>扩展的能力，可在**内主机企业推广'),
    ]
    hs_ = [max(56, 30 + wrap_lines(t, pw - 156, 19) * 29) for _, _, t in effects]
    top, bottom = yy + 22, y0 + ph - 30
    gap_ = (bottom - top - sum(hs_)) / float(len(effects) - 1)
    ey = top
    for (lab_, ic, txt), hh in zip(effects, hs_):
        rect(s, x2 + 40, ey, 56, 56, fill=M3.bg)
        icon(s, ic, M3.main, x2 + 40 + 12, ey + 12, 32)
        text(s, x2 + 116, ey - 2, 200, 30, lab_, size=21, bold=True, color=M3.text, lh=30, wrap=False)
        text(s, x2 + 116, ey + 28, pw - 156, hh - 26, txt, size=19, color=BODY, lh=29)
        ey += hh + gap_


# =====================================================================  05 立项必要性与紧迫性
def s05_urgency(d):
    s = d.new_slide()
    frame(d, s, '立项必要性与紧迫性：晚一年布局，就多一年被动', '政策与规划已有明确要求 · 受制于人的风险已经量化 · 验证环节存在差距 · 窗口期有限', '研究意义与价值')
    L, LW = 64, 880
    R, RW = 984, 872
    # 一、政策
    num_header(s, L, 214, '一', '政策与规划已有明确要求', w=LW)
    bullets(s, L, 262, LW, 250, [
        '《“***”铁路科技创新规划》（****〔****〕**号）将<b>数字孪生</b>列为智能铁路关键技术，明确要求“开展智能建造<b>数字孪生平台</b>研发应用”，研发具备“<b>自感知、自决策、自适应</b>”能力的*****，并提出“到****年<b>智能铁路技术全面突破</b>”',
        '该规划同时坦承短板：“部分关键核心技术亟待突破，<b>更高速度、更加智能、更高效率</b>技术有待补强”',
        '进入“***”，智能铁路已由“技术突破”转入“<b>体系化落地</b>”，数字样机与智能验证成为核心抓手',
    ], RED, size=20, lh=30, sa=8)
    # 二、风险量化
    num_header(s, L, 530, '二', '底层能力受制于人的风险已经量化', w=LW)
    kw_ = (LW - 2 * 16) / 3.0
    tiles = [
        ('约 5%', '研发设计类工业软件<br>国产化率', '有研究报告指出；<br>CAD/CAE/EDA 长期被欧美少数巨头企业垄断'),
        ('超 3 亿美元', '全球最大 CAE 厂商<br>年研发投入', '相当于国内所有工业软件<br>公司投入总和的数倍'),
        ('约 5% / 年', '用户端软件<br>年均涨价', '用户端每年软件支出达<br>数百万元，且不提供个性化定制'),
    ]
    for i, (v, lab_, sub) in enumerate(tiles):
        kpi(s, L + i * (kw_ + 16), 576, kw_, 226, v, lab_, sub, color=RED, value_size=40, label_size=19,
            sub_size=17)
    icon(s, 'alert-triangle', RED, L, 824, 28)
    text(s, L + 40, 818, LW - 40, 40, '**指南亦已将“研发设计软件<b>受制于人</b>”“设计仿真数据割裂”列为待解决事项', size=20, color=BODY,
         lh=30)
    # 三、差距
    num_header(s, R, 214, '三', '验证环节的差距：现状 → 目标形态', w=RW)
    gaps = [('验证方式', '实物试验为主', '虚拟验证前移，数字孪生驱动'), ('建模范围', '单系统、离线', '多系统实时协同'),
            ('工况覆盖', '受试验条件限制', '边界与故障工况可计算覆盖'), ('经验依赖', '强，靠专家判断', '判据知识化、流程自动化'),
            ('结论可信', '难复现、难追溯', '可回放、可追溯、可回归')]
    text(s, R + 150, 258, 260, 30, '现状', size=17, bold=True, color=MUTED, lh=30, wrap=False)
    text(s, R + 470, 258, 300, 30, '目标形态', size=17, bold=True, color=OK.text, lh=30, wrap=False)
    gy = 292
    for dim, now, goal in gaps:
        text(s, R, gy, 140, 50, dim, size=20, bold=True, color=INK, anchor='m', lh=30, wrap=False)
        chip(s, R + 140, gy, now, color=SUB, fill=GRAYBG, line=None, size=19, h=50, w=280, align='l', radius=0)
        tri(s, R + 440, gy + 25, 16, 22, TRI)
        chip(s, R + 462, gy, goal, color=OK.text, fill=OK.bg, line=None, size=19, h=50, w=RW - 462, bold=True,
             align='l', radius=0)
        gy += 60
    # 四、窗口期
    num_header(s, R, 610, '四', '窗口期有限', w=RW)
    rect(s, R, 656, RW, 206, fill=PANEL)
    quote = ('《规划》面向世界科技前沿提出，要“<b>抢占新一轮科技革命和产业变革制高点</b>，保持我国铁路科技创新领先优势，<b>增强国际话语权和影响力</b>”。数字孪生与智能体技术正处于从概念验证走向工程落地的窗口期——<r>若“***”期间不能拿下协同运行与调度引擎的自主权，验证环节将再次依赖国外工具链，重演工业软件的路径。</r>')
    qn_ = wrap_lines(quote, RW - 90, 20)
    qt = 656 + (206 - qn_ * 31) / 2.0
    icon(s, 'quote', MUTED, R + 20, qt + 2, 32)
    text(s, R + 66, qt, RW - 90, qn_ * 31 + 4, quote, size=20, color=BODY, lh=31)
    # 结论
    summary(s, L, 894, 1792, 96, '结论：本课题要做的不是“锦上添花的工具”，而是<r>验证环节的自主引擎</r>。晚一年布局，就多一年被动。',
            title_size=28)


# =====================================================================  06 引擎的未来价值（对领域发展的推动作用）
def s06_future(d):
    s = d.new_slide()
    frame(d, s, '未来价值：从“一个课题的交付物”到“可生长的验证能力底座”', '对领域发展的推动作用：能力扩展 · 场景延伸 · 范式转变 · 资产沉淀', '研究意义与价值')
    items = [
        ('能力扩展', 'puzzle', M1, '新增一个关键系统＝新一轮建模与集成', '按标准化接口接入智能体即可，<b>引擎不需重构</b>'),
        ('场景延伸', 'arrows-split-2', M2, '数字样机只服务研发验证', '延伸到<b>智能运维</b>（性能复现、状态评估）；方法体系可迁移至其他装备领域'),
        ('范式转变', 'repeat', M3, '验证是项目式、一次性的', '<b>常态化、可回归</b>——设计改一次，全工况验证自动重跑'),
        ('资产沉淀', 'database', M4, '经验在专家个人手里', '判据、工况、结论全程留痕，<b>个人经验沉淀为组织资产</b>'),
    ]
    cw, ch, y0 = (1792 - 3 * 24) / 4.0, 560, 214
    for i, (t, ic, th, now, after) in enumerate(items):
        x = 64 + i * (cw + 24)
        rect(s, x, y0, cw, ch, fill=WHITE, line=LINE)
        icon(s, ic, th.main, x + (cw - 60) / 2.0, y0 + 30, 60)
        card_title(s, x, y0 + 102, cw, t, th, size=27)
        # 现在
        tag(s, x + 28, y0 + 184, '现在', fill=GRAYBG, color=SUB, size=17, h=30, padx=12)
        text(s, x + 28, y0 + 224, cw - 56, 64, now, size=20, color=SUB, lh=30)
        tri(s, x + cw / 2.0, y0 + 312, 18, 26, TRI, direction='d')
        # 有了引擎之后
        ay = y0 + 338
        rect(s, x + 20, ay, cw - 40, ch - 338 - 20, fill=th.bg)
        tag(s, x + 40, ay + 18, '有了引擎之后', fill=th.main, size=17, h=30, padx=12)
        text(s, x + 40, ay + 60, cw - 80, ch - 338 - 20 - 70, after, size=21, color=INK, lh=32)
    # 底部：红条总结 + 机车标尺
    summary(s, 64, 832, 880, 120, '引擎的价值不在“跑通一次验证”',
            '而在于把列车性能验证变成一项<r>可积累、可复用、可生长</r>的能力。', title_size=30, desc_size=22)
    ruler(s, 980, 952, 876, step=44, tick=10)
    loco_train(s, 980 + 4, 952, 284, n=3, gap=8, faded=True)


# =====================================================================  07 核心命题
def s07_core(d):
    s = d.new_slide()
    frame(d, s, '核心命题：让列车性能验证像软件回归测试一样', '四层架构构成“协同运行与调度引擎” · 对接项目7总体目标 · 全流程跑在可复现沙箱中', '核心命题')
    L, LW = 64, 952
    header(s, L, 212, '本课题的四层架构', note='自下而上：基础层 → 决策层 → 执行层 → 集成层', w=LW)
    layers = {
        4: ['列车级数字样机', '多智能体协同运行', '可复现沙箱'],
        3: ['仿真智能体', '用例生成', '测试识别', '测试报告', '数据回归'],
        2: ['目标解析与任务分解', '环境自装配', '保真度与算力调度'],
        1: ['验证判据知识库', '任务工作流定义'],
    }
    y = 262
    bh, gap = 126, 12
    for i in (4, 3, 2, 1):
        th = LAYER[i]
        rect(s, L, y, LW, bh, fill=WHITE, line=LINE)
        rect(s, L, y, 250, bh, fill=th.bg)
        rect(s, L, y, BAR, bh, fill=th.main)
        text(s, L + 26, y + 24, 214, 34, LAYER_NO[i] + ' ' + LAYER_KIND[i], size=24, bold=True, color=th.text, lh=34,
             wrap=False)
        text(s, L + 26, y + 66, 218, 30, LAYER_NAME[i], size=18, bold=True, color=INK, lh=26, wrap=False)
        items = layers[i]
        arrow = (i == 3)
        g = 26 if arrow else 14
        cx = L + 250 + 26
        chips_row(s, cx, y + (bh - 44) / 2.0, items, gap=g, arrow=arrow, arrow_color=th.main, color=INK, line=th.main,
                  size=19, h=44, padx=14, lw=1.5)
        y += bh + gap
    # 右：课题定位
    R, RW = 1056, 800
    header(s, R, 212, '课题定位：四层如何构成“协同运行与调度引擎”', w=RW)
    maps = [
        ('多智能体模型的<b>高效实时协同运行</b>', [(3, '③ 执行层'), (4, '④ 集成层')]),
        ('多智能体模型的<b>调度引擎</b>', [(2, '② 决策层')]),
        ('形成<b>包含关键系统仿真智能体/智能孪生模型的列车级数字样机及构建方法体系</b>', [(4, '④ 集成层产出')]),
        ('支持<b>列车关键性能验证与智能运维</b>', [(4, '④ 集成层数字样机的用途')]),
    ]
    text(s, R, 262, 480, 30, '课题要求', size=17, bold=True, color=MUTED, lh=30, wrap=False)
    text(s, R + 500, 262, 300, 30, '四层中的落点', size=17, bold=True, color=MUTED, lh=30, wrap=False)
    my = 296
    for req, tags in maps:
        n = wrap_lines(strip_tags(req), 480, 20)
        rh = max(74, n * 30 + 30)
        line(s, R, my, R + RW, my, LINE, 1)
        text(s, R, my, 480, rh, req, size=20, color=BODY, anchor='m', lh=30)
        tx = R + 500
        for k, t in tags:
            tx += tag(s, tx, my + (rh - 32) / 2.0, t, fill=LAYER[k].main, size=17, h=32, padx=12) + 8
        my += rh
    line(s, R, my, R + RW, my, LINE, 1)
    # 对接项目7总体目标
    header(s, R, my + 22, '对接项目7总体目标', w=RW)
    goals = [('模型自动标定', '③ 执行层', 3, 'adjustments'), ('指标实时推演', '④ 协同运行', 4, 'activity'),
             ('性能自动评估', '③ 测试识别', 3, 'clipboard-check')]
    gw = (RW - 2 * 14) / 3.0
    gy = my + 66
    for i, (g, t, k, ic) in enumerate(goals):
        gx = R + i * (gw + 14)
        rect(s, gx, gy, gw, 112, fill=PANEL, line=LINE)
        rect(s, gx, gy, gw, min(4, BAR), fill=LAYER[k].main)
        icon(s, ic, LAYER[k].main, gx + 18, gy + 26, 30)
        text(s, gx + 58, gy + 22, gw - 66, 38, g, size=21, bold=True, color=INK, anchor='m', lh=30, wrap=False)
        text(s, gx + 18, gy + 68, gw - 24, 30, '→ ' + t, size=18, bold=True, color=LAYER[k].text, lh=26, wrap=False)
    # 工程现状 → 价值锚点
    by, bw = 846, 872
    rect(s, L, by, bw, 144, fill=GRAYBG)
    tag(s, L + 24, by + 16, '工程现状', fill=RED, size=16, h=28, padx=10)
    text(s, L + 124, by + 16, bw - 148, 28, '一次控制软件迭代的投入；复杂场景 HIL 模拟不了，只能靠现场验证', size=16, color=SUB,
         anchor='m', lh=22, wrap=False)
    tw_ = (bw - 48 - 3 * 12) / 4.0
    for i, (v, u, lab_) in enumerate([('20', '人', '研发'), ('10', '人', '测试'), ('10', '人', '现场'), ('<s=18><n>约</n></s>2', '个月', '周期')]):
        x = L + 24 + i * (tw_ + 12)
        rect(s, x, by + 58, tw_, 68, fill=WHITE, line=LINE)
        text(s, x + 16, by + 58, tw_ - 32, 68, '%s<s=16><n><c=%s> %s</c></n></s>' % (v, SUB, u), size=32, bold=True,
             color=INK, font=MONO, anchor='m', lh=40, wrap=False)
        text(s, x + tw_ - 16 - 70, by + 58, 70, 68, lab_, size=16, bold=True, color=MUTED, align='r', anchor='m', lh=22,
             wrap=False)
    tri(s, L + bw + 24, by + 72, 20, 26, TRI)
    summary(s, L + bw + 48, by, 1792 - bw - 48, 144, '价值锚点：设计改一次，全工况自动重跑一遍',
            '全流程跑在<b>可复现沙箱</b>中，验证可回放、可追溯，研发验证转向“数字样机为主”。', title_size=26,
            desc_size=19)
    notes(s, '说明：“工程现状”数据为一次控制软件迭代的实际投入（研发 20 人＋测试 10 人＋现场 10 人，约 2 个月）。')



# =====================================================================  研究内容 ①—④ 共用：左栏、场景条、要点卡、附图
def content_left(s, layer, label=None, body=None, outs=()):
    """左栏：四层导航 +（可选）补充说明 + 产出。"""
    th = LAYER[layer]
    y = layer_nav(s, 64, 214, 260, layer) + 22
    if label:
        text(s, 64, y, 260, 34, label, size=22, bold=True, color=INK, lh=32, wrap=False)
        paras = body if isinstance(body, list) else [{'t': body}]
        bh = sum(wrap_lines(p['t'], 260, 19) * 30 + p.get('sa', 0) for p in paras)
        text(s, 64, y + 46, 260, bh + 4, body, size=19, color=SUB, lh=30)
        y += 46 + bh + 28
    if outs:
        text(s, 64, y, 260, 34, '产出', size=22, bold=True, color=INK, lh=32, wrap=False)
        y += 44
        for o in outs:
            h = wrap_lines('✓ ' + o, 260 - STYLE['concl_bar'] - 24, 18, True) * 26 + 22
            concl(s, 64, y, 260, h, o, th, size=18, check=True, pad=12)
            y += h + 8
    return y


def scene_band(s, x, y, w, h, scene, role, role_label='本层作用', split=0.6):
    """工程场景条：左为实际工程场景（遇到的问题），右为本层作用（怎么解决）。"""
    rect(s, x, y, w, h, fill=GRAYBG)
    rect(s, x, y, STYLE['concl_bar'], h, fill=RED)
    lw = round(w * split)
    tag(s, x + 24, y + 14, '工程场景', fill=RED, size=16, h=28, padx=10)
    text(s, x + 24, y + 50, lw - 52, h - 54, scene, size=18, color=INK, lh=27)
    line(s, x + lw, y + 18, x + lw, y + h - 18, RULE, 1)
    text(s, x + lw + 28, y + 14, 240, 28, role_label, size=16, bold=True, color=MUTED, anchor='m', lh=22, wrap=False)
    text(s, x + lw + 28, y + 50, w - lw - 56, h - 54, role, size=18, color=BODY, lh=27)


def card_title(s, x, y, w, title, th, size=25, uw=None):
    """居中标题 + 同色下划线（参考页四栏卡片）。"""
    text(s, x, y, w, 40, title, size=size, bold=True, color=INK, align='c', anchor='m', lh=round(size * 1.3),
         wrap=False)
    uw = uw or min(w - 40, text_width(title, size, True) + 12)
    rect(s, x + (w - uw) / 2.0, y + 48, uw, 3, fill=th.main)


def point_card(s, x, y, w, h, no, title, desc, th, sub=None, title_size=20, desc_size=17):
    """研究要点卡：顶部色条 + 编号标题（+ 副标题）+ 说明。"""
    rect(s, x, y, w, h, fill=WHITE, line=LINE)
    rect(s, x, y, w, BAR, fill=th.main)
    text(s, x + 22, y + 16, w - 44, 32, '<c=%s><m>%s</m></c>  %s' % (th.text, no, title), size=title_size, bold=True,
         color=INK, anchor='m', lh=round(title_size * 1.3), wrap=False)
    ty = y + 58
    if sub:
        text(s, x + 22, ty - 4, w - 44, 24, sub, size=15, bold=True, color=th.text, lh=22, wrap=False)
        ty += 26
    text(s, x + 22, ty, w - 44, y + h - ty - 8, desc, size=desc_size, color=BODY, lh=round(desc_size * 1.52))


class Fig:
    """附图：所有形状收进一个组合，整组删除即可换成实际图片。"""

    def __init__(self, s, name):
        self.grp = s.shapes.add_group_shape()
        self.grp.name = name
        self.shapes = self.grp.shapes
        self._no = s._no


def fig_area(s, x, y, w, h, title, note='示意图'):
    """附图区：小节标题 + 浅底面板（面板与图内形状同属一个组合）。返回 (组合, 面板顶 y)。"""
    header(s, x, y, title, note=note, w=w)
    f = Fig(s, '附图 · ' + strip_tags(title))
    rect(f, x, y + 46, w, h - 46, fill=PANEL)
    return f, y + 46


def fig_placeholder(s, x, y, w, h, title, hint):
    """预留附图位置：虚线框 + 图片图标 + 名称 + 建议内容（整组删除后插入实际图片）。"""
    f = Fig(s, '附图位置 · ' + title)
    rect(f, x, y, w, h, fill=PH.bg, line=PH.main, lw=1.2, dash='dash')
    icon(f, 'photo', PH.main, x + (w - 40) / 2.0, y + h / 2.0 - 54, 40)
    text(f, x + 16, y + h / 2.0 - 6, w - 32, 28, '附图位置 · ' + title, size=18, bold=True, color=PH.text, align='c',
         lh=26, wrap=False)
    text(f, x + 16, y + h / 2.0 + 24, w - 32, 24, hint, size=15, color=PH.text, align='c', lh=22, wrap=False)
    return f


def fnode(f, x, y, w, h, txt, th, kind='n', size=17):
    """附图节点：n 普通 / k 强调 / on 实心 / w 警示。"""
    st = {'n': dict(color=BODY, fill=WHITE, line=LINE),
          'k': dict(color=th.text, fill=th.bg, line=th.soft, bold=True),
          'on': dict(color=on_color(th.main), fill=th.main, line=None, bold=True),
          'w': dict(color=RED, fill=RED_BG, line=RED, bold=True)}[kind]
    return chip(f, x, y, txt, w=w, h=h, size=size, **st)


def arrow(f, x1, y1, x2, y2, color=MUTED, lw=2, dash=None):
    return line(f, x1, y1, x2, y2, color, lw, dash=dash, tail='triangle')


def smooth(pts, n=8):
    """Catmull-Rom 插值：把折线点加密成平滑曲线。"""
    P = [pts[0]] + list(pts) + [pts[-1]]
    out = []
    for i in range(1, len(P) - 2):
        p0, p1, p2, p3 = P[i - 1], P[i], P[i + 1], P[i + 2]
        for k in range(n):
            t = k / float(n)
            out.append(tuple(0.5 * (2 * p1[j] + (p2[j] - p0[j]) * t + (2 * p0[j] - 5 * p1[j] + 4 * p2[j] - p3[j]) * t * t
                                    + (3 * p1[j] - p0[j] - 3 * p2[j] + p3[j]) * t ** 3) for j in (0, 1)))
    out.append(tuple(pts[-1]))
    return out


FIG_NOTE = ('说明：“工程场景”可替换为本单位的实际案例；“示意图”为临时附图（形状组合，可直接编辑），'
            '如有实际图片，选中该组合删除后插入即可。')


# =====================================================================  08 研究内容① 基础层
def s08_layer1(d):
    s = d.new_slide()
    th = LAYER[1]
    frame(d, s, '研究内容① 测试知识库与任务工作流：为上层提供知识与规则', '研究什么：验证判据知识化 · 任务工作流建模 · 知识与流程的迭代机制', '研究内容① 基础层')
    content_left(s, 1, '与上层的关系', '知识库与工作流是“任务理解与编排调度”的<b>知识与规则来源</b>。',
                 outs=['验证判据知识库', '任务工作流定义<br>与配置规范'])
    X, XW = 356, 1500
    scene_band(s, X, 214, XW, 112,
               '判定一个工况是否合格，依据分散在国标、行标、企业标准、试验大纲和专家经验中，查找与比对主要靠人工，判定口径不易统一。',
               '为上层提供<b>知识与规则</b>——没有判据知识库，任务理解只能靠通用大模型“猜”；没有任务工作流，编排调度无章可循。')
    cw = (XW - 48) / 3.0
    cards = [
        ('验证判据知识化', '从设计规范、运用要求与历史案例中提取判据，形式化为<b>可自动比对</b>的判据条目，构建验证判据知识库'),
        ('任务工作流建模', '定义验证任务的分解规则、执行顺序、流转条件与异常分支，形成<b>可配置、可复用</b>的任务工作流'),
        ('知识与流程的迭代机制', '依据验证结果回流，持续更新判据条目与工作流定义'),
    ]
    for i, (t, desc) in enumerate(cards):
        point_card(s, X + i * (cw + 24), 342, cw, 136, '0%d' % (i + 1), t, desc, th, desc_size=18)
    # 附图：判据知识化（上）→ 任务工作流（下）→ 结果回流（右）
    f, py = fig_area(s, X, 494, XW, 496, '从规范到判据、从判据到工作流')
    fx = X + 24
    step = 204                                   # 工作流节点间距（节点宽 150）
    tag(f, fx, py + 26, '01 判据知识化', fill=th.bg, color=th.text, size=16, h=28, padx=10)
    sy = py + 66
    for k, (a, b) in enumerate([('设计规范', '条款与限值要求'), ('运用要求', '运用条件与指标'), ('历史案例', '故障与异常记录')]):
        yy = sy + k * 60
        rect(f, fx, yy, 210, 52, fill=WHITE, line=LINE)
        icon(f, 'file-text', MUTED, fx + 12, yy + 12, 28)
        text(f, fx + 50, yy + 5, 150, 22, a, size=17, bold=True, color=INK, lh=22, wrap=False)
        text(f, fx + 50, yy + 28, 150, 20, b, size=14, color=MUTED, lh=20, wrap=False)
    cy = sy + 86
    arrow(f, fx + 220, cy, fx + 314, cy)
    text(f, fx + 210, cy - 30, 114, 22, '提取 · 形式化', size=15, bold=True, color=SUB, align='c', lh=20, wrap=False)
    rx, rw = fx + 326, 392
    rect(f, rx, sy, rw, 172, fill=WHITE, line=th.soft, lw=1.2)
    rect(f, rx, sy, rw, 36, fill=th.bg)
    text(f, rx + 14, sy, rw - 28, 36, '判据条目（结构化，可自动比对）', size=16, bold=True, color=th.text, anchor='m', lh=22,
         wrap=False)
    rows = [('验证对象', '牵引控制系统'), ('关键性能', '速度跟踪偏差'), ('适用工况', '长大下坡 · 湿轨'),
            ('判定规则', '|实测 − 目标| ≤ 允许偏差'), ('来源', '设计规范条款')]
    for k, (a, b) in enumerate(rows):
        ry = sy + 40 + k * 25
        text(f, rx + 14, ry, 84, 25, a, size=15, color=MUTED, anchor='m', lh=20, wrap=False)
        text(f, rx + 104, ry, rw - 118, 25, b, size=16, bold=(k == 3), color=INK, anchor='m', lh=22, wrap=False)
    arrow(f, rx + rw + 10, cy, rx + rw + 80, cy)
    text(f, rx + rw + 10, cy - 30, 70, 22, '入库', size=15, bold=True, color=SUB, align='c', lh=20, wrap=False)
    kx = fx + 4 * step                           # 知识库与下方“结果比对”节点对齐
    shape(f, MSO_SHAPE.CAN, kx, sy + 4, 150, 164, fill=th.bg, line=th.main, lw=1.2)
    text(f, kx, sy + 44, 150, 112, '判据<br>知识库', size=20, bold=True, color=th.text, align='c', anchor='m', lh=28,
         wrap=False)
    ex = kx + 180
    for k, t in enumerate(['速度跟踪偏差', '牵引能耗', '运行平稳性']):
        yy = sy + 10 + k * 52
        rect(f, ex, yy, X + XW - 24 - ex, 42, fill=WHITE, line=LINE)
        rect(f, ex, yy, 4, 42, fill=th.main)
        text(f, ex + 18, yy, 320, 42, '<n><c=%s>判据</c></n>　%s' % (MUTED, t), size=17, bold=True, color=INK,
             anchor='m', lh=22, wrap=False)
    # 02 任务工作流
    y2 = py + 262
    tag(f, fx, y2, '02 任务工作流', fill=th.bg, color=th.text, size=16, h=28, padx=10)
    ny = y2 + 42
    names = ['验证目标', '任务分解', '环境装配', '仿真执行', '结果比对', '验证报告']
    for k, (t, kd) in enumerate(zip(names, ['n', 'n', 'n', 'n', 'k', 'on'])):
        fnode(f, fx + k * step, ny, 150, 44, t, th, kd, size=18)
        if k < 5:
            arrow(f, fx + k * step + 156, ny + 22, fx + (k + 1) * step - 6, ny + 22)
    text(f, fx + 4 * step + 150, ny - 6, step - 150, 20, '通过', size=14, bold=True, color=SUB, align='c', lh=18,
         wrap=False)
    bx0, bx1 = fx + 4 * step + 75, fx + 2 * step + 75
    poly(f, [(bx0, ny + 44), (bx0, ny + 78), (bx1, ny + 78), (bx1, ny + 46)], RED, 2, tail='triangle')
    text(f, bx1, ny + 84, bx0 - bx1, 24, '未通过 → 异常分支：调整后重跑', size=15, bold=True, color=RED, align='c', lh=22,
         wrap=False)
    arrow(f, kx + 75, sy + 170, kx + 75, ny - 4, th.main, 1.5, dash='dash')
    text(f, kx + 86, (sy + 170 + ny) / 2.0 - 11, 100, 22, '调用判据', size=15, bold=True, color=th.text, lh=20,
         wrap=False)
    # 03 迭代：结果回流到知识库
    lx = X + XW - 40
    poly(f, [(fx + 5 * step + 154, ny + 22), (lx, ny + 22), (lx, sy + 162)], OK.main, 2, tail='triangle')
    tag(f, fx + 5 * step + 170, ny - 60, '03 迭代机制', fill=th.bg, color=th.text, size=16, h=28, padx=10)
    text(f, fx + 5 * step + 170, ny - 28, 230, 22, '结果回流，更新判据与工作流', size=15, color=SUB, lh=20, wrap=False)
    notes(s, FIG_NOTE)


# =====================================================================  09 研究内容② 决策层
def s09_layer2(d):
    s = d.new_slide()
    th = LAYER[2]
    frame(d, s, '研究内容② 任务理解与智能编排调度：引擎在此落地', '研究什么：验证目标解析与任务分解 · 测试环境自动装配 · 模型保真度与算力自适应调度 · 执行监控与失败重规划',
          '研究内容② 决策层')
    content_left(s, 2, '上下游关系', '向下承接①的<b>知识与规则</b>，向③执行层下发<b>执行方案</b>（智能体组合、工况、算力）。',
                 outs=['多智能体协同运行与调度引擎（软件1套）'])
    X, XW = 356, 1500
    scene_band(s, X, 214, XW, 112,
               '验证一条线路，要覆盖区段 × 环境温度 × 轨面状态 × 载重 × 手柄序列的组合，工况数量成倍增长：全用高精度模型跑不完，全用简化模型，关键工况又不可信。',
               '本课题的<b>核心决策层</b>，“多智能体协同运行与调度引擎”在此落地：按工况特征分配模型精度与算力，把高精度用在关键工况上。')
    cw = (XW - 60) / 4.0
    cards = [
        ('验证目标解析与任务分解', '把“验证该机车某项性能是否满足要求”拆解为可执行验证任务，并匹配知识库中的判据与工作流'),
        ('测试环境自动装配', '按验证目标自动完成智能体组合、接口连接与算力资源配置'),
        ('模型保真度与算力自适应调度', '按工况特征匹配模型精度与算力配置：常规工况降阶提速，极端工况全保真'),
        ('执行监控与失败重规划', '运行过程监控、异常中断后重规划，支撑指标实时推演'),
    ]
    for i, (t, desc) in enumerate(cards):
        point_card(s, X + i * (cw + 20), 342, cw, 150, '0%d' % (i + 1), t, desc, th)
    # 附图：一次验证任务中的调度过程
    f, py = fig_area(s, X, 508, XW, 482, '一次验证中的调度过程：按工况切换保真度、分配算力，异常时重规划')
    fx = X + 24
    chips_row(f, fx, py + 18, ['<k>验证任务</k>　速度跟踪 ＋ 牵引能耗', '<k>01 任务分解</k>　对象 · 性能 · 判据',
                               '<k>02 环境装配</k>　控制系统 · 车辆动力学 · 线路智能体', '<k>03—04 调度执行与监控</k>'],
              gap=44, arrow=True, arrow_color=TRI, size=16, h=40, padx=16, color=BODY, line=LINE)
    tx0, tx1 = fx + 160, X + XW - 24
    segs = [(0.19, '常规直线', 0), (0.16, '小半径曲线', 1), (0.19, '常规直线', 0), (0.27, '长大下坡 · 湿轨', 1),
            (0.19, '常规直线', 0)]
    xs = [tx0]
    for fr, _, _ in segs:
        xs.append(xs[-1] + fr * (tx1 - tx0))
    n3, n4 = '<c=%s>03</c> ' % th.text, '<c=%s>04</c> ' % th.text
    for lab, ry, rh in [('工况', 76, 40), (n3 + '模型保真度', 124, 40), (n3 + '算力 · GPU', 176, 56),
                        (n3 + '算力 · NPU', 244, 30), (n4 + '执行监控', 298, 40)]:
        text(f, fx, py + ry, 156, rh, lab, size=16, bold=True, color=INK, anchor='m', lh=22, wrap=False)
    for xb in xs[1:-1]:
        line(f, xb, py + 70, xb, py + 342, RULE, 1, dash='dash')
    gb, nb, my = py + 232, py + 274, py + 318       # GPU 基线 / NPU 基线 / 监控线
    gpu, npu = [(tx0, gb)], [(tx0, nb), (tx0, nb - 7)]
    for k, (fr, nm, hi) in enumerate(segs):
        x0, x1 = xs[k], xs[k + 1]
        rect(f, x0 + 2, py + 76, x1 - x0 - 4, 40, fill=WHITE, line=LINE)
        text(f, x0 + 2, py + 76, x1 - x0 - 4, 40, nm, size=16, color=SUB, align='c', anchor='m', lh=22, wrap=False)
        fl, fc = (th.main, on_color(th.main)) if hi else (th.bg, th.text)
        rect(f, x0 + 2, py + 124, x1 - x0 - 4, 40, fill=fl)
        text(f, x0 + 2, py + 124, x1 - x0 - 4, 40, '全保真 · 准' if hi else '降阶模型 · 快', size=16, bold=True, color=fc,
             align='c', anchor='m', lh=22, wrap=False)
        lv = 46 if hi else 16
        gpu += [(x0, gb - lv), (x1, gb - lv)]
        if k:
            npu += [(x0 - 15, nb - 7), (x0, nb - 23), (x0 + 15, nb - 7)]
    gpu.append((tx1, gb))
    npu += [(tx1, nb - 7), (tx1, nb)]
    poly(f, gpu, th.main, 1.5, fill=th.soft, closed=True)
    poly(f, npu, th.main, 1.5, fill=th.soft, closed=True)
    text(f, tx1 - 300, py + 180, 292, 22, '多物理域模型并行仿真', size=14, color=MUTED, align='r', lh=20, wrap=False)
    text(f, tx1 - 300, py + 246, 292, 20, '切换时的推理与编排决策', size=14, color=MUTED, align='r', lh=20, wrap=False)
    # 执行监控：异常中断 → 重规划
    line(f, tx0, my, tx1, my, RULE, 2)
    for k, (fr, nm, hi) in enumerate(segs):
        x0, x1 = xs[k], xs[k + 1]
        if k == 3:
            xa = x0 + 0.62 * (x1 - x0)
            icon(f, 'circle-check', OK.main, x0 + 0.25 * (x1 - x0) - 11, my - 11, 22)
            icon(f, 'alert-circle', RED, xa - 13, my - 13, 26)
            text(f, xa + 18, my - 30, 120, 22, '异常中断', size=14, bold=True, color=RED, lh=20, wrap=False)
            poly(f, [(xa, my + 13), (xa, my + 28), (x0 + 16, my + 28), (x0 + 16, my + 12)], RED, 1.8, tail='triangle')
            text(f, x0 + 16, my + 30, xa - x0 - 16, 20, '重规划', size=14, bold=True, color=RED, align='c', lh=20,
                 wrap=False)
        else:
            icon(f, 'circle-check', OK.main, (x0 + x1) / 2.0 - 11, my - 11, 22)
    arrow(f, tx0, py + 404, tx1, py + 404, RULER, 2)
    text(f, tx1 - 240, py + 378, 236, 22, '仿真推进（里程）', size=14, color=MUTED, align='r', lh=20, wrap=False)
    notes(s, FIG_NOTE)


# =====================================================================  10 研究内容③ 执行层
def s10_layer3(d):
    s = d.new_slide()
    th = LAYER[3]
    frame(d, s, '研究内容③ 智能体执行：一条完整的验证执行链', '仿真智能体 → 测试用例生成 → 测试识别 → 测试报告 → 数据回归，结果回流、支持回归重跑', '研究内容③ 执行层')
    content_left(s, 3, '对接项目7总体目标', [{'t': '<b>模型自动标定</b> → 仿真智能体', 'sa': 8}, {'t': '<b>性能自动评估</b> → 测试识别'}],
                 outs=['关键系统仿真智能体／<br>智能孪生模型', '用例与线路数据集', '验证报告与追溯记录'])
    X, XW = 356, 1500
    scene_band(s, X, 214, XW, 112,
               '2024 年 12 月，FXN5C 在临哈线长大下坡（约 −10℃）由牵引回手柄后，冷却风扇较长时间保持 40Hz，低温水从 25℃ 被冷却至 0℃，高温水仍在 78℃，散热器冻结。',
               '这类多系统、长时序的问题，现有 HIL＋经验数据难以复现；执行层要在研发阶段把它<b>跑出来、判出来、找到原因</b>。',
               role_label='本层怎么做')
    gap = 24
    cw, ch, y0 = (XW - 4 * gap) / 5.0, 200, 342
    steps = [
        ('仿真智能体', '研制与接入',
         '研制控制系统、车辆动力学与线路等<b>仿真智能体与智能孪生模型</b>；标准化接入、时序同步、<b>模型自动标定</b>'),
        ('测试用例生成', '用例与线路数据按需自生成', '按覆盖完备性要求，自动生成测试用例与线路模型数据（平纵断面、曲线半径、超高、坡度、不平顺谱）'),
        ('测试识别', '结果评判 · 异常定位', '逐工况对标判据，<b>识别异常工况并定位原因</b>，输出结果可信性判断'),
        ('测试报告', '报告生成', '验证结论自动汇总，生成<b>可追溯</b>的验证报告'),
        ('数据回归', '回流迭代', '验证结果回流，迭代模型标定、判据与工况集；支持设计变更后的<b>回归重跑</b>'),
    ]
    cxs = []
    for i, (t, sub, desc) in enumerate(steps):
        x = X + i * (cw + gap)
        cxs.append(x + cw / 2.0)
        point_card(s, x, y0, cw, ch, '0%d' % (i + 1), t, desc, th, sub=sub)
        if i < 4:
            tri(s, x + cw + gap / 2.0, y0 + ch / 2.0, 14, 20, TRI)
    yb = y0 + ch
    poly(s, [(cxs[4], yb), (cxs[4], yb + 22), (cxs[0], yb + 22), (cxs[0], yb + 2)], OK.main, 2, tail='triangle')
    lab = '数据回归：结果回流，迭代模型标定、判据与工况集 · 设计变更后回归重跑'
    lw_ = text_width(lab, 16, True) + 28
    lx_ = (cxs[0] + cxs[4]) / 2.0 - lw_ / 2.0
    rect(s, lx_, yb + 10, lw_, 24, fill=WHITE)
    text(s, lx_, yb + 10, lw_, 24, lab, size=16, bold=True, color=OK.text, align='c', anchor='m', lh=22, wrap=False)
    # 附图：用现场案例说明测试识别
    f, py = fig_area(s, X, 588, XW, 402, '以临哈线散热器冻结为例：逐工况对标判据，识别异常并定位原因', note='示意图 · 数据据现场记录')
    fx = X + 24
    px0, px1, pt, pb = fx + 70, fx + 900, py + 44, py + 226

    def P(t, v):                                 # t：时间 0—1；v：水温 ℃（0—90）
        return (px0 + t * (px1 - px0), pb - v / 90.0 * (pb - pt))

    lim = 6                                      # 防冻限值（示意位置，不标数值）
    hx0 = P(0.63, 0)[0]
    rect(f, hx0, pt - 6, px1 - hx0, pb - pt + 6, fill=RED_BG, line=RED, lw=1, dash='dash')
    ly_ = P(0, 42)[1]
    icon(f, 'alert-circle', RED, hx0 + 10, ly_ - 9, 18)
    text(f, hx0 + 32, ly_ - 11, px1 - hx0 - 38, 22, '低温水越限：过冷 → 散热器冻结', size=14, bold=True, color=RED,
         lh=20, wrap=False)
    line(f, px0, P(0, lim)[1], px1, P(0, lim)[1], RED, 1.5, dash='dash')
    text(f, px0 + 8, P(0, lim)[1] - 22, 200, 20, '防冻限值（判据）', size=14, bold=True, color=RED, lh=20, wrap=False)
    xh = P(0.2, 0)[0]
    line(f, xh, pt - 8, xh, pb, SUB, 1.2, dash='dash')
    text(f, xh + 6, pt - 8, 90, 20, '回手柄', size=14, bold=True, color=SUB, lh=20, wrap=False)
    ht = [(0, 80), (0.1, 79.5), (0.2, 79), (0.35, 78.2), (0.5, 78), (0.7, 78.2), (0.85, 78), (1, 78)]
    lt = [(0, 25), (0.1, 25.4), (0.2, 25), (0.3, 21), (0.42, 15), (0.55, 9), (0.67, 4.5), (0.78, 1.6), (0.86, 0.2),
          (1, 0)]
    poly(f, smooth([P(*p) for p in ht]), SUB, 2.5)
    poly(f, smooth([P(*p) for p in lt]), th.main, 2.5)
    text(f, P(0.3, 78)[0], P(0, 78)[1] - 26, 200, 22, '高温水 维持 78℃', size=14, bold=True, color=SUB, lh=20,
         wrap=False)
    text(f, P(0.3, 25)[0] + 10, P(0, 25)[1] - 20, 200, 22, '低温水 25℃ → 0℃', size=14, bold=True, color=th.text, lh=20,
         wrap=False)
    for v in (78, 25, 0):
        text(f, px0 - 56, P(0, v)[1] - 10, 48, 20, '%d℃' % v, size=13, color=MUTED, font=MONO, align='r', lh=20,
             wrap=False)
    arrow(f, px0, pb, px1 + 22, pb, MUTED, 1.5)
    arrow(f, px0, pb, px0, pt - 22, MUTED, 1.5)
    text(f, px0 - 60, pt - 30, 52, 20, '水温', size=14, color=MUTED, align='r', lh=20, wrap=False)
    text(f, px1 - 30, pb - 24, 52, 20, '时间', size=14, color=MUTED, align='r', lh=20, wrap=False)
    # 时间轴下的条带：手柄 / 风扇 / 线路与环境
    for k, (lab_, cells) in enumerate([('手柄', [(0, 0.2, '牵引', 'n'), (0.2, 1, '回手柄', 'k')]),
                                       ('风扇', [(0, 0.2, '', 'n'), (0.2, 1, '冷却风扇 40Hz 持续运行', 'w')]),
                                       ('线路', [(0, 1, '临哈线长大下坡 · 环境约 −10℃', 'n')])]):
        yy = pb + 10 + k * 30
        text(f, fx, yy, 60, 26, lab_, size=14, bold=True, color=MUTED, anchor='m', lh=20, wrap=False)
        for a, b, t, kd in cells:
            x0, x1 = P(a, 0)[0], P(b, 0)[0]
            fl, ln, fc = {'n': (WHITE, LINE, SUB), 'k': (th.bg, th.soft, th.text), 'w': (RED_BG, RED, RED)}[kd]
            rect(f, x0 + 1, yy, x1 - x0 - 2, 26, fill=fl, line=ln)
            if t:
                text(f, x0 + 1, yy, x1 - x0 - 2, 26, t, size=14, bold=(kd != 'n'), color=fc, align='c', anchor='m',
                     lh=20, wrap=False)
    # 识别结果
    rw = 470
    rx = X + XW - 24 - rw
    text(f, rx, py + 12, rw, 28, '识别结果', size=18, bold=True, color=INK, lh=26, wrap=False)
    hy = py + 46
    cols = [(rx, 170), (rx + 170, 150), (rx + 320, 150)]
    for (cx_, cw_), t in zip(cols, ['工况', '判据', '结论']):
        text(f, cx_ + 10, hy, cw_ - 10, 24, t, size=14, bold=True, color=MUTED, anchor='m', lh=20, wrap=False)
    line(f, rx, hy + 26, rx + rw, hy + 26, SUB, 1.2)
    res = [('牵引工况', '低温水防冻', '✓ 满足', False), ('回手柄 · 长大下坡', '低温水防冻', '✗ 过冷', True),
           ('回手柄 · 长大下坡', '高温水温度', '✓ 正常', False)]
    for k, (a, b, c, bad) in enumerate(res):
        ry = hy + 28 + k * 36
        if bad:
            rect(f, rx, ry, rw, 36, fill=RED_BG)
        for (cx_, cw_), t, st in zip(cols, (a, b, c), (dict(color=INK), dict(color=BODY),
                                                       dict(color=RED if bad else OK.text, bold=True))):
            text(f, cx_ + 10, ry, cw_ - 10, 36, t, size=15, anchor='m', lh=20, wrap=False, **st)
        line(f, rx, ry + 36, rx + rw, ry + 36, LINE, 1)
    cy2 = hy + 28 + 3 * 36 + 14
    rect(f, rx, cy2, rw, 84, fill=WHITE, line=RED, lw=1.2)
    text(f, rx + 16, cy2 + 8, rw - 32, 22, '原因定位', size=15, bold=True, color=RED, lh=20, wrap=False)
    text(f, rx + 16, cy2 + 32, rw - 32, 48, '回手柄后柴油机发热量下降，风扇仍以 40Hz 运行，低温水回路被过度冷却', size=15,
         color=BODY, lh=22)
    text(f, rx, cy2 + 96, rw, 22, '→ 写入验证报告；设计修改后，低温工况全量回归重跑', size=15, bold=True, color=th.text, lh=20,
         wrap=False)
    notes(s, '说明：本页工程场景为 2024 年 12 月临哈线现场案例（数据据现场记录）；水温曲线为示意，'
             '如有现场记录曲线，选中附图组合删除后替换即可。防冻限值线仅示意位置，请按设计规范取值。')


# =====================================================================  11 研究内容④ 集成层
def s11_layer4(d):
    s = d.new_slide()
    th = LAYER[4]
    frame(d, s, '研究内容④ 架构融合与协同运行：融合成一个能跑起来的整体', '研究什么：总体架构与接口规范 · 多智能体协同运行 · 可复现沙箱 · 列车级数字样机', '研究内容④ 集成层')
    content_left(s, 4, '对接项目7总体目标', '<b>指标实时推演</b> → 协同运行',
                 outs=['列车级数字样机', '列车级数字样机构建方法体系与集成技术规范'])
    X, XW = 356, 1500
    scene_band(s, X, 214, XW, 112,
               '现有手段是硬件在环（HIL）结合经验数据模拟运用工况，控制、冷却、动力、线路等系统难以在同一环境中长时序协同运行，复杂场景无法模拟。',
               '把前三层<b>融合成一个能跑起来的整体</b>，形成列车级数字样机；现有 HIL 台架可作为控制器在环节点接入。')
    cw = (XW - 60) / 4.0
    cards = [
        ('总体架构与接口规范', '定义四层之间的<b>接口、数据流与控制流</b>，形成<b>可扩展</b>的架构规范'),
        ('多智能体协同运行', '异构智能体的<b>实时协同求解、时序同步与状态同步</b>'),
        ('可复现沙箱', '<b>隔离</b>：单路模型发散不扩散；<b>记录</b>：参数、模型版本、随机种子；<b>回放</b>：任意一次验证可复现'),
        ('列车级数字样机', '集成关键系统仿真智能体，支持<b>列车关键性能验证与智能运维</b>'),
    ]
    for i, (t, desc) in enumerate(cards):
        point_card(s, X + i * (cw + 20), 342, cw, 150, '0%d' % (i + 1), t, desc, th)
    # 附图：数字样机构成
    f, py = fig_area(s, X, 508, XW, 482, '列车级数字样机构成：关键系统智能体经标准化接口接入引擎，在沙箱中协同运行')
    fx = X + 24
    zx, zy, zw, zh = fx, py + 34, 1072, 384
    dashed_zone(f, zx, zy, zw, zh, '03 可复现沙箱：隔离 · 记录 · 回放', label_fill=ARCH, label_x=zx + 24, label_size=16)
    ax0, agap = zx + 24, 12
    aw = (zw - 48 - 5 * agap) / 6.0
    ay = zy + 38
    by_ = ay + 56 + 38
    for k, (nm, on) in enumerate([('控制系统', True), ('车辆动力学', True), ('线路', True), ('牵引系统', False),
                                  ('制动系统', False), ('辅助系统', False)]):
        x = ax0 + k * (aw + agap)
        if on:
            rect(f, x, ay, aw, 56, fill=th.bg, line=th.soft, lw=1.2)
            text(f, x, ay, aw, 56, nm + '<br>智能体', size=17, bold=True, color=th.text, align='c', anchor='m', lh=23,
                 wrap=False)
        else:
            rect(f, x, ay, aw, 56, fill=WHITE, line=RULE, lw=1.2, dash='dash')
            text(f, x, ay, aw, 56, nm + '<br><s=14>预留扩展</s>', size=17, color=MUTED, align='c', anchor='m', lh=23,
                 wrap=False)
        line(f, x + aw / 2.0, ay + 56, x + aw / 2.0, by_ - 10, MUTED if on else RULE, 1.5, dash=None if on else 'dash')
        rect(f, x + aw / 2.0 - 5, by_ - 10, 10, 10, fill=th.main if on else RULE)
    bw_ = zw - 48
    rect(f, ax0, by_, bw_, 62, fill=th.main)
    text(f, ax0, by_, bw_, 62, ['02 多智能体协同运行与调度引擎',
                                {'t': '经 01 标准化接口接入 · 实时协同求解 · 时序同步 · 状态同步', 'size': 15, 'bold': False,
                                 'lh': 21}], size=20, bold=True, color=on_color(th.main), align='c', anchor='m', lh=27,
         wrap=False)
    lxc = ax0 + bw_ / 2.0
    rail = zy + zh - 16
    lw_, lh_ = 480, 120
    arrow(f, lxc, by_ + 64, lxc, rail - lh_ - 6, MUTED, 2)
    text(f, lxc + 10, by_ + 68, 80, 22, '集成', size=14, bold=True, color=SUB, lh=20, wrap=False)
    ruler(f, ax0, rail, bw_, step=36, tick=8, lw=2)
    pic(f, LOCO, lxc - lw_ / 2.0 + 60, rail - lh_ * 0.992, lw_, lh_)
    text(f, ax0, rail - lh_ + 4, 250, 30, '04 列车级数字样机', size=20, bold=True, color=INK, lh=28, wrap=False)
    text(f, ax0, rail - lh_ + 40, 240, 48, '关键系统仿真智能体 ＋ 协同运行与调度引擎', size=15, color=SUB, lh=22)
    # 右：接入现有条件 / 支撑的应用
    ox = zx + zw + 40
    ow = X + XW - 24 - ox
    text(f, ox, by_ - 30, ow, 22, '可接入现有条件', size=15, bold=True, color=MUTED, lh=22, wrap=False)
    hb = by_ - 4
    rect(f, ox, hb, ow, 70, fill=WHITE, line=LINE)
    rect(f, ox, hb, BAR, 70, fill=ARCH)
    text(f, ox + 20, hb + 8, ow - 36, 26, 'HIL 台架：控制器在环', size=17, bold=True, color=INK, lh=24, wrap=False)
    text(f, ox + 20, hb + 38, ow - 36, 24, '经验数据用于智能体标定', size=15, color=SUB, lh=22, wrap=False)
    arrow(f, ox - 4, by_ + 31, ax0 + bw_ + 6, by_ + 31, ARCH, 1.8, dash='dash')
    sy_ = zy + zh - 150
    text(f, ox, sy_ - 30, ow, 22, '支撑', size=15, bold=True, color=MUTED, lh=22, wrap=False)
    rect(f, ox, sy_, ow, 150, fill=WHITE, line=LINE)
    rect(f, ox, sy_, BAR, 150, fill=th.main)
    text(f, ox + 20, sy_ + 12, ow - 36, 26, '列车关键性能验证', size=17, bold=True, color=th.text, lh=24, wrap=False)
    text(f, ox + 20, sy_ + 42, ow - 36, 24, '速度跟踪 · 牵引能耗 · 运行平稳性', size=15, color=BODY, lh=22, wrap=False)
    line(f, ox + 20, sy_ + 76, ox + ow - 20, sy_ + 76, LINE, 1)
    text(f, ox + 20, sy_ + 86, ow - 36, 26, '智能运维', size=17, bold=True, color=th.text, lh=24, wrap=False)
    text(f, ox + 20, sy_ + 116, ow - 36, 24, '性能复现 · 状态评估', size=15, color=BODY, lh=22, wrap=False)
    arrow(f, zx + zw + 4, sy_ + 75, ox - 6, sy_ + 75, MUTED, 2)
    notes(s, FIG_NOTE)


# =====================================================================  12 研究方法
def s12_methods(d):
    s = d.new_slide()
    frame(d, s, '研究方法：四层各有方法，攻关四个技术难点', '知识工程方法 · 智能调度方法 · 实时协同与可信评判方法 · 系统集成方法', '研究方法')
    X, XW = 300, 1556
    side_label(s, 64, 222, 206, '四层方法', '四层各自的研究方法与关键技术手段')
    rows = [
        (1, '知识工程方法', ['领域知识建模', '标准规范形式化', '工作流建模']),
        (2, '智能调度方法', ['任务分解与匹配', '多目标寻优', '保真度—算力协同配置']),
        (3, '实时协同与可信评判方法', ['分布式协同仿真', '时序同步', '自动标定与不确定度量化']),
        (4, '系统集成方法', ['标准化接口', '服务化封装', '沙箱化运行与全程留痕']),
    ]
    c1, c2 = 400, 350
    for t, xx in (('层', X), ('研究方法', X + c1 + 24), ('关键技术手段', X + c1 + c2 + 24)):
        text(s, xx, 214, 300, 32, t, size=17, bold=True, color=MUTED, lh=24, anchor='m', wrap=False)
    y = 254
    rh = 118
    for k, meth, techs in rows:
        th = LAYER[k]
        rect(s, X, y, XW, rh, fill=WHITE, line=LINE)
        rect(s, X, y, c1, rh, fill=th.bg)
        rect(s, X, y, BAR, rh, fill=th.main)
        text(s, X + 26, y + 24, c1 - 40, 34, LAYER_NO[k] + ' ' + LAYER_NAME[k], size=22, bold=True, color=INK, lh=32,
             wrap=False)
        text(s, X + 26, y + 64, c1 - 40, 30, LAYER_KIND[k], size=18, bold=True, color=th.text, lh=26, wrap=False)
        text(s, X + c1 + 24, y, c2 - 24, rh, meth, size=24, bold=True, color=th.text, anchor='m', lh=34)
        chips_row(s, X + c1 + c2 + 24, y + (rh - 46) / 2.0, techs, gap=12, size=19, h=46, padx=16, color=INK,
                  line=th.main, lw=1.5)
        y += rh + 10
    # 技术难点
    side_label(s, 64, 790, 206, '技术难点', '方法层面的攻关点')
    probs = [
        ('判据知识化', 'file-search', '工程判据多以经验与文档形式存在，结构化为机器可用知识难度大'),
        ('时序同步', 'clock', '多物理域模型步长差异显著，联合求解易失真'),
        ('保真度调度', 'adjustments', '需在线权衡精度与效率，且极端工况不允许外推'),
        ('可信评判', 'shield-check', '极端工况实测数据稀缺，结论可信性难以自证'),
    ]
    pw = (XW - 3 * 20) / 4.0
    for i, (t, ic, b) in enumerate(probs):
        x = X + i * (pw + 20)
        problem_card(s, x, 786, pw, 204, t, b, body_size=21)
        icon(s, ic, MUTED, x + pw - 24 - 32, 786 + 26, 32)


# =====================================================================  13 技术路线（架构图版式）
def s13_route(d):
    s = d.new_slide()
    frame(d, s, '技术路线：四层递进 ＋ 沙箱底座，闭环迭代', '四层递进：① 知识 → ② 调度 → ③ 执行 → ④ 集成；沙箱底座：隔离 · 记录 · 回放', '技术路线')
    L, LW = 64, 1112
    header(s, L, 212, '总体技术路线', note='四层递进 ＋ 可复现沙箱底座', w=LW)
    zy, zh = 290, 556
    dashed_zone(s, L, zy, LW, zh, '可复现沙箱底座（隔离 · 记录 · 回放）', label_fill=ARCH, fill=ZONE_FILL, label_x=L + 28)
    cols = {
        1: ['判据知识化', '工作流建模', '知识迭代'],
        2: ['目标解析与任务分解', '环境自装配', '保真度与算力调度', '监控与重规划'],
        3: ['仿真智能体', '用例生成', '测试识别／报告', '数据回归'],
    }
    heads = {1: '① 知识库与工作流', 2: '② 任务理解与编排调度', 3: '③ 智能体执行'}
    cw, ag = 300, 64
    x0 = L + (LW - 3 * cw - 2 * ag) / 2.0
    cy0 = zy + 74
    chh = 52 + 16 + 4 * 42 + 3 * 10 + 16
    cxs = {}
    for k in (1, 2, 3):
        th = LAYER[k]
        x = x0 + (k - 1) * (cw + ag)
        cxs[k] = x + cw / 2.0
        rect(s, x, cy0, cw, chh, fill=WHITE, line=th.main, lw=1.5)
        shp = rect(s, x, cy0, cw, 52, fill=th.main)
        text(s, x, cy0, cw, 52, heads[k], size=21, bold=True, color=WHITE, align='c', anchor='m', lh=28, shp=shp,
             wrap=False)
        for j, n in enumerate(cols[k]):
            chip(s, x + 16, cy0 + 68 + j * 52, n, w=cw - 32, h=42, size=19, color=INK, fill=PANEL, line=LINE)
        if k < 3:
            line(s, x + cw + 8, cy0 + chh / 2.0, x + cw + ag - 8, cy0 + chh / 2.0, MUTED, 2.5, tail='triangle')
    # 数据回归 ③ → ①（红色虚线，走上方）
    ly = cy0 - 22
    poly(s, [(cxs[3], cy0), (cxs[3], ly), (cxs[1], ly), (cxs[1], cy0 - 2)], RED, 2, dash='dash', tail='triangle')
    tl = '数据回归：结果回流更新判据与工作流'
    tw = text_width(tl, 17, True) + 24
    text(s, cxs[2] - tw / 2.0, ly - 14, tw, 28, tl, size=17, bold=True, color=RED, align='c', anchor='m', lh=24,
         wrap=False, shp=rect(s, cxs[2] - tw / 2.0, ly - 14, tw, 28, fill=ZONE_FILL))
    # 汇入 ④
    by = cy0 + chh
    bus = by + 26
    for k in (1, 2, 3):
        line(s, cxs[k], by, cxs[k], bus, MUTED, 2)
    line(s, cxs[1], bus, cxs[3], bus, MUTED, 2)
    y4 = bus + 34
    line(s, cxs[2], bus, cxs[2], y4 - 2, MUTED, 2.5, tail='triangle')
    th = LAYER[4]
    bx, bw = x0, 3 * cw + 2 * ag
    rect(s, bx, y4, bw, 96, fill=th.bg, line=th.main, lw=1.5)
    shp = rect(s, bx, y4, 330, 96, fill=th.main)
    text(s, bx, y4, 330, 96, '④ 架构融合与协同运行', size=23, bold=True, color=WHITE, align='c', anchor='m', lh=30, shp=shp,
         wrap=False)
    line(s, bx + 350, y4 + 48, bx + 430, y4 + 48, th.main, 2.5, tail='triangle')
    icon(s, 'train', th.main, bx + 450, y4 + 24, 48)
    text(s, bx + 512, y4, 300, 96, '列车级数字样机', size=28, bold=True, color=INK, anchor='m', lh=36, wrap=False)
    text(s, bx + 760, y4, bw - 780, 96, '集成层产出', size=18, bold=True, color=th.text, anchor='m', align='r', lh=26,
         wrap=False)
    # 右：四层之间的衔接
    R, RW = 1216, 640
    header(s, R, 212, '四层之间的衔接', w=RW)
    links = [
        ('① → ②', LAYER[1].main, '知识与规则输入', '判据与工作流是任务理解与调度的<b>知识与规则输入</b>'),
        ('② → ③', LAYER[2].main, '执行方案下发', '<b>执行方案</b>（智能体组合、工况、算力）下发至执行层'),
        ('③ → ④', LAYER[3].main, '结果汇入', '执行结果汇入集成层，形成<b>数字样机与验证结论</b>'),
        ('③ → ①', RED, '数据回归', '<r>数据回归</r>：结果回流更新判据与工作流'),
    ]
    y = 262
    lh_ = (zy + zh - 262 - 3 * 13) / 4.0
    for t, col, st, body in links:
        rect(s, R, y, RW, lh_, fill=WHITE, line=LINE)
        rect(s, R, y, BAR, lh_, fill=col)
        tag(s, R + 26, y + 20, t, fill=col, size=19, h=34, padx=14)
        text(s, R + 150, y + 20, RW - 176, 34, st, size=19, bold=True, color=RED if col == RED else SUB, align='r',
             anchor='m', lh=26, wrap=False)
        text(s, R + 26, y + 66, RW - 52, 60, body, size=20, color=BODY, lh=30)
        y += lh_ + 13
    # 路线闭环
    y1 = 882
    rect(s, 64, y1, STYLE['concl_bar'], 104, fill=RED)
    text(s, 92, y1, 200, 104, '路线闭环', size=28, bold=True, color=INK, anchor='m', lh=36, wrap=False)
    loop = ['判据知识化', '任务生成与调度', '协同执行', '结果识别', '数据回归', '更新知识库']
    ex = chips_row(s, 300, y1 + 30, loop, gap=46, arrow=True, arrow_color=MUTED, size=20, h=46, padx=18, color=INK,
                   line=LINE, bold=True)
    icon(s, 'refresh', RED, ex + 20, y1 + 33, 40)
    text(s, ex + 70, y1 + 30, 1856 - ex - 70, 46, '闭环迭代', size=22, bold=True, color=RED, anchor='m', lh=30,
         wrap=False)


# =====================================================================  14 典型验证场景
def s14_scenario(d):
    s = d.new_slide()
    frame(d, s, '典型验证场景：控速＋能耗验证，全流程交给智能体', '输入一句话需求 → ① 任务理解 → ② 任务编排 → ③ 智能体执行 → 验证报告 ＋ 全要素留痕', '典型验证场景')
    # 输入
    y0 = 214
    shp = rect(s, 64, y0, 250, 84, fill=INK)
    text(s, 64, y0, 250, 84, '输入 · 一句话需求', size=22, bold=True, color=WHITE, align='c', anchor='m', lh=30, shp=shp,
         wrap=False)
    rect(s, 314, y0, 1542, 84, fill=PANEL)
    icon(s, 'quote', MUTED, 340, y0 + 24, 36)
    text(s, 394, y0, 1440, 84, '验证 FXN5C 机车在临哈线上按给定运行图运行时的<r>速度跟踪</r>与<r>牵引能耗</r>是否满足要求。', size=26, bold=True,
         color=INK, anchor='m', lh=36, wrap=False)
    # 三步面板
    py, ph = 326, 556
    ws = [450, 806, 456]
    gap = (1792 - sum(ws)) / 2.0
    xs = [64, 64 + ws[0] + gap, 64 + ws[0] + ws[1] + 2 * gap]
    specs = [('任务理解', '变成可执行的验证任务', M1, 'target-arrow'), ('任务编排', '生成数据、工况与环境', M2, 'route'),
             ('智能体执行', '协同求解与结果自评判', M3, 'robot')]
    tops = []
    for i, ((t, sub, th, ic), x, w) in enumerate(zip(specs, xs, ws)):
        rect(s, x, py, w, ph, fill=WHITE, line=LINE)
        rect(s, x, py, w, BAR, fill=th.main)
        icon_badge(s, ic, th, x + 28, py + 28, 64)
        text(s, x + 110, py + 26, w - 200, 38, '%s' % t, size=27, bold=True, color=INK, lh=36, wrap=False)
        text(s, x + 110, py + 64, w - 200, 30, '→ ' + sub, size=18, bold=True, color=th.text, lh=26, wrap=False)
        num_mark(s, x + w - 28 - 90, py + 18, 90, '0%d' % (i + 1), th.num, size=40)
        line(s, x + 28, py + 114, x + w - 28, py + 114, LINE, 1.5)
        tops.append(py + 132)
        if i < 2:
            tri(s, x + w + gap / 2.0, py + ph / 2.0, 20, 28, TRI)
    # ① 任务理解
    x, w, th = xs[0], ws[0], M1
    y = tops[0]
    for lab_, val in [('验证对象', '牵引控制系统<br>（含级位控制、防空转／防滑行）'), ('关键性能', '<b>速度跟踪偏差、牵引能耗</b>'),
                      ('验证判据', '从设计规范与运用要求中提取，形式化为<b>可自动比对</b>的判据条目')]:
        tag(s, x + 28, y, lab_, fill=th.bg, color=th.text, size=18, h=32, padx=12)
        n = wrap_lines(strip_tags(val), w - 56, 20)
        text(s, x + 28, y + 40, w - 56, n * 31, val, size=20, color=BODY, lh=31)
        y += 40 + n * 31 + 30
    # ② 任务编排：2×2 子块
    x, w, th = xs[1], ws[1], M2
    y = tops[1]
    sw = (w - 56 - 20) / 2.0
    shh = 196
    subs = [(x + 28, y), (x + 28 + sw + 20, y), (x + 28, y + shh + 16), (x + 28 + sw + 20, y + shh + 16)]

    def sub_head(bx, by, t):
        rect(s, bx, by, sw, shh, fill=PANEL)
        text(s, bx + 16, by + 12, sw - 32, 30, t, size=20, bold=True, color=INK, lh=28, wrap=False)

    bx, by = subs[0]
    sub_head(bx, by, '线路模型数据自生成')
    chips_row(s, bx + 16, by + 56, ['平纵断面', '曲线半径', '超高'], gap=10, size=17, h=38, padx=12, color=th.text,
              line=th.main)
    chips_row(s, bx + 16, by + 104, ['坡度', '轨道不平顺谱'], gap=10, size=17, h=38, padx=12, color=th.text, line=th.main)
    bx, by = subs[1]
    sub_head(bx, by, '工况矩阵自生成')
    for k, (g, v) in enumerate([('线路条件', '长大坡道／小半径曲线／隧道'), ('环境', '干轨／湿轨'), ('载荷', '满载／空载')]):
        yy = by + 50 + k * 34
        text(s, bx + 16, yy, 30, 30, '×' if k else '', size=18, bold=True, color=MUTED, lh=26, wrap=False)
        text(s, bx + 42, yy, 90, 30, g, size=17, bold=True, color=th.text, lh=26, wrap=False)
        text(s, bx + 130, yy, sw - 140, 30, v, size=17, color=INK, lh=26, wrap=False)
    text(s, bx + 16, by + 154, sw - 32, 30, '→ 自动覆盖边界与组合工况', size=17, bold=True, color=th.text, lh=26,
         wrap=False)
    bx, by = subs[2]
    sub_head(bx, by, '测试环境自装配')
    for k, t_ in enumerate(['牵引控制智能体', '车辆动力学智能体', '线路智能体']):
        ry = by + 44 + k * 36
        if k:
            text(s, bx + 16, ry, 24, 34, '＋', size=18, bold=True, color=MUTED, align='c', anchor='m', lh=24,
                 wrap=False)
        chip(s, bx + 44, ry, t_, h=30, size=17, padx=12, color=INK, line=LINE)
    text(s, bx + 16, by + 154, sw - 32, 26, '接口对接与算力配置', size=17, color=SUB, lh=24, wrap=False)
    bx, by = subs[3]
    sub_head(bx, by, '保真度分配')
    for k, (a, b, fc, bg) in enumerate([('常规直线工况', '降阶模型（快）', th.text, WHITE),
                                        ('长大下坡、小半径曲线、湿轨', '全保真模型（准）', RED, RED_BG)]):
        yy = by + 46 + k * 66
        text(s, bx + 16, yy, sw - 32, 28, a, size=17, color=BODY, lh=26, wrap=False)
        chip(s, bx + 16, yy + 28, '→ ' + b, h=36, size=17, color=fc, fill=bg, line=None if bg != WHITE else th.main,
             bold=True, padx=12)
    # ③ 智能体执行：纵向小流程
    x, w, th = xs[2], ws[2], M3
    y = tops[2]
    bullets(s, x + 28, y, w - 56, 70, ['三类智能体<b>实时协同求解</b>，模型参数自动标定'], th.main, size=20, lh=31, sa=0)
    fy = y + 76
    flow = [('逐工况对标判据', None, INK, LINE), ('定位异常工况', '如某小半径曲线段<br>速度偏差超限', RED, RED),
            ('追溯原因', '黏着利用不足？<br>控制参数不匹配？', RED, RED)]
    for k, (a, b, fc, lc) in enumerate(flow):
        chip(s, x + 28, fy, a, w=190, h=42, size=19, color=fc, line=lc, bold=True)
        if b:
            text(s, x + 232, fy - 2, w - 260, 46, b, size=17, color=SUB, anchor='m', lh=23)
        if k < 2:
            line(s, x + 123, fy + 42, x + 123, fy + 66, TRI, 2, tail='triangle')
        fy += 68
    concl(s, x + 28, py + ph - 28 - 96, w - 56, 96, '输出：验证报告 ＋ 全要素留痕<br>（可回放、可回归）', th, size=21)
    # 结论
    summary(s, 64, 904, 1792, 86, '这一次验证里：工况没人手写、线路数据没人准备、环境没人搭、异常没人看图判断。这就是“全流程交给智能体”。',
            title_size=27)


# =====================================================================  15 引擎创新点
def s15_innovation(d):
    s = d.new_slide()
    frame(d, s, '引擎创新点：先进性不在单项技术，而在调度的维度与粒度', '三个创新点：调度对象是“智能体” · 调度决策多一个维度：模型保真度 · 调度基础是可复现沙箱', '引擎创新点')
    specs = [
        (M1, 'robot', '调度对象是“智能体”，不是“模型”', '调度的是算例', '调度的是<b>可自主标定、自主判断的智能体</b>',
         '调度粒度从“跑模型”提升到“派任务”，人退出执行回路', None),
        (M2, 'adjustments', '调度决策多一个维度：模型保真度', '在“快”与“准”之间取舍',
         '按工况特征动态分配——<b>常规工况用降阶模型提速，关键与极端工况强制全保真保精度</b>', '批量提速与关键工况精度同时成立', ['降阶 · 快', '全保真 · 准']),
        (M3, 'history', '调度基础是可复现沙箱', '仿真结果难以复现', '对<b>参数、模型版本、随机种子</b>全程留痕', '验证可回放、可追溯、可回归',
         ['参数', '模型版本', '随机种子']),
    ]
    pw, ph, y0 = (1792 - 64) / 3.0, 676, 214
    for i, (th, ic, title, conv, ours, cap, mini) in enumerate(specs):
        x = 64 + i * (pw + 32)
        rect(s, x, y0, pw, ph, fill=WHITE, line=LINE)
        rect(s, x, y0, pw, BAR, fill=th.main)
        icon_badge(s, ic, th, x + 32, y0 + 30, 64)
        num_mark(s, x + pw - 32 - 100, y0 + 22, 100, '0%d' % (i + 1), th.num, size=46)
        text(s, x + 32, y0 + 110, pw - 64, 40, title, size=25, bold=True, color=INK, anchor='m', lh=34, wrap=False)
        line(s, x + 32, y0 + 166, x + pw - 32, y0 + 166, LINE, 1.5)
        # 常规做法
        cy = y0 + 184
        tag(s, x + 32, cy, '常规做法', fill=GRAYBG, color=SUB, size=17, h=30, padx=12)
        text(s, x + 32, cy + 40, pw - 64, 40, conv, size=21, color=SUB, lh=31)
        tri(s, x + pw / 2.0, cy + 102, 18, 26, TRI, direction='d')
        # 本课题
        by = cy + 124
        bh = 196
        rect(s, x + 32, by, pw - 64, bh, fill=th.bg)
        tag(s, x + 52, by + 18, '本课题', fill=th.main, size=17, h=30, padx=12)
        text(s, x + 52, by + 60, pw - 104, 96, ours, size=21, color=INK, lh=32)
        if mini:
            chips_row(s, x + 52, by + bh - 18 - 38, mini, gap=10, size=17, h=38, padx=14, color=th.text,
                      line=th.main, bold=True, fill=WHITE)
        else:
            w1 = chip(s, x + 52, by + bh - 18 - 38, '跑模型 · 算例', h=38, size=17, padx=14, color=SUB, fill=GRAYBG,
                      line=None, bold=True)
            tri(s, x + 52 + w1 + 20, by + bh - 18 - 19, 12, 16, th.main)
            chip(s, x + 52 + w1 + 40, by + bh - 18 - 38, '派任务 · 智能体', h=38, size=17, padx=14, color=th.text,
                 fill=WHITE, line=th.main, bold=True)
        # 带来的能力
        text(s, x + 32, by + bh + 22, 200, 28, '带来的能力', size=17, bold=True, color=MUTED, lh=26, wrap=False)
        concl(s, x + 32, y0 + ph - 28 - 84, pw - 64, 84, cap, th, size=21)
    summary(s, 64, 914, 1792, 76, '先进性不在单项技术，而在<r>调度决策多出的维度</r>与<r>调度对象粒度的提升</r>。', title_size=28)


# =====================================================================  16 技术对标与技术定位
def s16_benchmark(d):
    s = d.new_slide()
    frame(d, s, '技术对标：不重复造轮子，补上智能编排与实时协同这一层', '对标六类现有成熟技术：已解决什么 · 未解决什么 · 本课题定位在哪', '技术对标与定位')
    L = 64
    cols = [290, 372, 524]
    header(s, L, 212, '现有成熟技术对标', w=sum(cols))
    rows = [
        ('单系统商用仿真软件', '单专业高精度仿真', '单系统、离线、人工操作，难以形成列车级协同'),
        ('协同仿真标准与中间件', '模型封装与数据交换（接得上）', '只解决接口，不解决智能编排、实时性与算力调度'),
        ('硬件在环（HIL）试验台', '高保真实时验证', '成本高、构型专用、工况受限、不可批量'),
        ('实物与线路试验', '结果权威', '周期长、成本高、极端工况难覆盖、不可复现'),
        ('代理模型／降阶模型', '显著提速', '外推精度差，极端工况不可直接用'),
        ('通用智能体框架', '任务编排与工具调用', '缺多物理域实时协同与时序同步能力'),
    ]
    trs = [['<b>%s</b>' % a, '<c=%s>✓</c>  %s' % (OK.main, b), '<bc=%s>×</bc>  %s' % (RED, c)] for a, b, c in rows]
    table(s, L, 256, cols, trs, header=['现有成熟技术', '已解决', '未解决（本课题的定位）'], size=20, lh=30, pad_y=33,
          pad_x=18, col_styles={0: dict(size=20), 1: dict(color=OK.text), 2: dict(color=BODY)})
    # 右：技术定位
    R, RW = 1290, 566
    header(s, R, 212, '技术定位', w=RW)
    y = 262
    for k, t in enumerate(['不做仿真软件', '不做接口标准', '不做试验台']):
        cw_ = (RW - 2 * 10) / 3.0
        chip(s, R + k * (cw_ + 10), y, '× ' + t, w=cw_, h=46, size=19, color=SUB, fill=GRAYBG, line=None, bold=True)
    tri(s, R + RW / 2.0, y + 72, 16, 24, TRI, direction='d')
    by = y + 96
    rect(s, R, by, RW, 136, fill=RED_BG)
    rect(s, R, by, STYLE['concl_bar'], 136, fill=RED)
    text(s, R + 28, by + 14, RW - 48, 30, '在成熟技术之上，补上这一层', size=18, bold=True, color=RED, lh=26, wrap=False)
    text(s, R + 28, by + 48, RW - 48, 80, '面向列车级性能验证的<br><r>智能编排与实时协同</r>', size=27, bold=True, color=INK, lh=38)
    text(s, R, by + 150, RW, 28, '成熟技术底座', size=16, bold=True, color=MUTED, align='c', lh=24, wrap=False)
    base = ['商用仿真软件', '协同仿真标准', 'HIL 试验台', '实物与线路试验', '代理／降阶模型', '通用智能体框架']
    bw_ = (RW - 2 * 8) / 3.0
    for k, t in enumerate(base):
        chip(s, R + (k % 3) * (bw_ + 8), by + 182 + (k // 3) * 46, t, w=bw_, h=38, size=17, color=SUB, fill=PANEL,
             line=LINE)
    text(s, R, by + 318, RW, 150,
         '技术定位：不重复造轮子——不做仿真软件、不做接口标准、不做试验台；而是在成熟技术之上，补上“<b>面向列车级性能验证的智能编排与实时协同</b>”这一层。',
         size=21, color=BODY, lh=33)
    # 底部：与同类工作的差异
    summary(s, L, 890, 1792, 100, '与同类工作的差异',
            '现有工作多聚焦单一系统仿真或通用智能体平台；本课题面向<b>列车级性能验证</b>，把“协同运行与调度”做成<b>可交付的引擎</b>与<b>可复用的方法论体系</b>。',
            title_size=24, desc_size=20, gap=4)


# =====================================================================  17 阶段目标与进度安排
def s17_schedule(d):
    s = d.new_slide()
    frame(d, s, '阶段目标与进度：四阶段推进，前一阶段交付即后一阶段输入', '研究周期 2027.01—2028.12（24 个月） · 四个里程碑 M1—M4 与阶段交付物对齐', '阶段目标与进度安排')
    X0, X1 = 64, 1844
    mw = (X1 - X0) / 24.0
    stages = STAGES
    # 年份
    for yi, yr in enumerate(['2027 年', '2028 年']):
        text(s, X0 + yi * 12 * mw, 212, 12 * mw, 30, yr, size=19, bold=True, color=INK, align='c', lh=26, wrap=False)
    line(s, X0 + 12 * mw, 212, X0 + 12 * mw, 350, RULE, 1.5, dash='dash')
    # 甘特条
    by = 250
    for name, short, a, b, col, *_ in stages:
        x = X0 + a * mw + (2 if a else 0)
        w = (b - a) * mw - (2 if a else 0)
        shp = rect(s, x, by, w, 54, fill=col)
        fg = STAGE_ON[STAGE.index(col)] if STAGE_ON else on_color(col)
        full = '%s · %s' % (name, short)
        text(s, x, by, w, 54, full if text_width(full, 19, True) + 24 <= w else name, size=19, bold=True, color=fg,
             align='c', anchor='m', lh=26, shp=shp, wrap=False)
    # 月刻度
    ty = by + 62
    rect(s, X0, ty, X1 - X0, 3, fill=RULER)
    for m in range(25):
        xx = X0 + m * mw
        rect(s, xx - 1, ty + 3, 2, 12 if m % 12 == 0 else (8 if m % 3 == 0 else 5), fill=RULER)
    for m in range(24):
        text(s, X0 + (m + 0.5) * mw - 24, ty + 14, 48, 24, '%02d' % (m % 12 + 1), size=15, color=MUTED, font=MONO,
             align='c', lh=20, wrap=False)
    # 里程碑
    my = ty + 52
    for name, short, a, b, col, per, goal, deli, mk, mdate, mtext in stages:
        xx = X0 + b * mw
        shape(s, MSO_SHAPE.DIAMOND, xx - 12, my, 24, 24, fill=RED)
        al = 'r' if b == 24 else 'c'
        bx = xx - 230 if al == 'r' else xx - 115
        text(s, bx, my + 28, 230, 26, '%s · %s' % (mk, mdate), size=18, bold=True, color=RED, align=al, lh=24,
             wrap=False, font=FONT)
        text(s, bx, my + 54, 230, 26, mtext, size=17, color=SUB, align=al, lh=24, wrap=False)
    # 阶段卡片
    cy0, ch = 474, 400
    cw = (1792 - 3 * 36) / 4.0
    for i, (name, short, a, b, col, per, goal, deli, mk, mdate, mtext) in enumerate(stages):
        x = 64 + i * (cw + 36)
        rect(s, x, cy0, cw, ch, fill=WHITE, line=LINE)
        rect(s, x, cy0, BAR, ch, fill=col)
        tag(s, x + 26, cy0 + 22, name, fill=col, color=STAGE_ON[STAGE.index(col)] if STAGE_ON else on_color(col), size=18, h=32,
            padx=14)
        text(s, x + 140, cy0 + 22, cw - 160, 32, per, size=18, bold=True, color=MUTED, font=MONO, align='r',
             anchor='m', lh=24, wrap=False)
        text(s, x + 26, cy0 + 72, cw - 52, 40, goal, size=25, bold=True, color=INK, lh=34, wrap=False)
        line(s, x + 26, cy0 + 126, x + cw - 26, cy0 + 126, LINE, 1.5)
        text(s, x + 26, cy0 + 140, cw - 52, 28, '阶段交付', size=17, bold=True, color=MUTED, lh=24, wrap=False)
        parts = [p for p in deli.split('、')]
        bullets(s, x + 26, cy0 + 174, cw - 52, 150, parts,
                STAGE_BULLET[STAGE.index(col)] if STAGE_BULLET else ink_safe(col), size=19, lh=29, sa=6)
        rect(s, x + 20, cy0 + ch - 20 - 50, cw - 40, 50, fill=RED_BG)
        text(s, x + 36, cy0 + ch - 20 - 50, cw - 72, 50, '<r>%s</r>（%s）%s' % (mk, mdate, mtext), size=19, color=INK,
             anchor='m', lh=26, bold=True, wrap=False)
        if i < 3:
            tri(s, x + cw + 18, cy0 + ch / 2.0, 18, 24, TRI)
    summary(s, 64, 900, 1792, 92, '阶段间衔接：前一阶段交付物即后一阶段输入',
            '判据与接口规范 → 任务生成与智能体接入；智能体与引擎 → 样机集成；样机 → 性能验证', title_size=25, desc_size=20, gap=2)


# =====================================================================  18 子任务成果与考核方式
def s18_tasks(d):
    s = d.new_slide()
    frame(d, s, '子任务成果与考核方式：四个子任务，交付物逐级成为下一任务输入', '子任务1 知识 → 子任务2 调度 → 子任务3 执行 → 子任务4 集成；数据回归更新判据与工作流', '预期成果与效益')
    tasks = [
        (1, ['验证判据知识库', '任务工作流定义'], '知识库核验、报告评审'),
        (2, ['多智能体协同运行与<br>调度引擎（软件）'], '软件演示、功能核验'),
        (3, ['仿真智能体／孪生模型', '用例与线路数据', '验证报告'], '模型核验、比对试验'),
        (4, ['列车级数字样机', '构建方法体系与技术规范'], '示范应用验收、第三方验证'),
    ]
    links = ['判据与工作流', '执行方案', '验证结论']
    g = 112
    cw, ch, y0 = (1792 - 3 * g) / 4.0, 520, 214
    cxs = []
    for i, (k, outs, check) in enumerate(tasks):
        th = LAYER[k]
        x = 64 + i * (cw + g)
        cxs.append(x + cw / 2.0)
        rect(s, x, y0, cw, ch, fill=WHITE, line=LINE)
        rect(s, x, y0, cw, BAR, fill=th.main)
        tag(s, x + 26, y0 + 28, '子任务%d' % k, fill=th.main, size=18, h=34, padx=14)
        num_mark(s, x + cw - 26 - 90, y0 + 18, 90, '0%d' % k, th.num, size=40)
        text(s, x + 26, y0 + 80, cw - 52, 36, LAYER_NO[k] + ' ' + LAYER_NAME[k], size=23, bold=True, color=INK, lh=32,
             wrap=False)
        text(s, x + 26, y0 + 116, cw - 52, 28, '对应' + LAYER_KIND[k], size=17, bold=True, color=th.text, lh=24,
             wrap=False)
        line(s, x + 26, y0 + 160, x + cw - 26, y0 + 160, LINE, 1.5)
        text(s, x + 26, y0 + 176, cw - 52, 28, '主要成果', size=17, bold=True, color=MUTED, lh=24, wrap=False)
        bullets(s, x + 26, y0 + 210, cw - 52, 180, outs, th.main, size=21, lh=31, sa=8)
        text(s, x + 26, y0 + ch - 26 - 64 - 34, cw - 52, 28, '考核方式', size=17, bold=True, color=MUTED, lh=24,
             wrap=False)
        concl(s, x + 26, y0 + ch - 26 - 64, cw - 52, 64, check, th, size=20, check=True, pad=16)
        if i < 3:
            ax0, ax1 = x + cw + 10, x + cw + g - 10
            line(s, ax0, y0 + ch / 2.0, ax1, y0 + ch / 2.0, MUTED, 2.5, tail='triangle')
            text(s, x + cw, y0 + ch / 2.0 - 40, g, 32, links[i], size=16, bold=True, color=SUB, align='c', anchor='b',
                 lh=22, wrap=False)
    # 回流：子任务4 → 子任务2
    yb = y0 + ch
    poly(s, [(cxs[3], yb), (cxs[3], yb + 46), (cxs[0], yb + 46), (cxs[0], yb + 2)], OK.main, 2.5, tail='triangle')
    text(s, cxs[0], yb + 56, cxs[3] - cxs[0], 30, '数据回归：更新判据与工作流', size=20, bold=True, color=OK.text, align='c',
         lh=28, wrap=False)
    summary(s, 64, 872, 1792, 118, '每个子任务的交付物即下一子任务的输入', '形成“<r>知识—调度—执行—集成</r>”的闭环。', title_size=28,
            desc_size=22)


# =====================================================================  19 预期目标与成果对应
def s19_goals(d):
    s = d.new_slide()
    frame(d, s, '预期目标与成果对应：四项目标，每项都有对应成果物', '调度引擎 · 数字样机与性能验证 · 方法体系、智能运维与扩展能力 · 示范应用', '预期成果与效益')
    goals = [
        ('01', '调度引擎', M1, '形成<b>多智能体协同运行与调度引擎1套</b>，支持异构仿真智能体的实时协同运行与算力资源动态调度',
         [('多智能体协同运行与调度引擎', '软件1套')]),
        ('02', '数字样机与性能验证', M2,
         '建成包含车辆动力学与线路等关键系统仿真智能体的<b>列车级数字样机</b>，实现对列车速度跟踪控制品质、牵引能耗与再生能量利用、运行平稳性等关键性能的<b>多工况协同验证</b>',
         [('关键系统仿真智能体／智能孪生模型', '1套'), ('列车级数字样机', '1套'), ('示范应用与验证报告', '1份')]),
        ('03', '方法体系、智能运维\n与扩展能力', M3,
         '形成列车级数字样机构建方法体系与标准化集成技术规范，<b>支持列车关键性能验证与智能运维</b>，具备向制动、辅助系统等关键系统扩展的能力',
         [('列车级数字样机构建方法体系与技术规范', '1套')]),
        ('04', '示范应用', M4, '在 <b>FXN5C 机车</b>上完成示范应用（以临哈线典型区段为验证场景），关键性能验证的<b>效率与工况覆盖度显著提升</b>，形成可推广的智能验证能力',
         [('示范应用与验证报告', '1份'), ('专利 · 论文 · 标准草案', '4项 · 2篇 · 2项')]),
    ]
    text(s, 64, 212, 400, 32, '预期目标', size=18, bold=True, color=MUTED, lh=26, wrap=False)
    text(s, 1250, 212, 400, 32, '对应成果物', size=18, bold=True, color=MUTED, lh=26, wrap=False)
    y = 250
    rh = 176
    for no, name, th, txt, outs in goals:
        rect(s, 64, y, 1792, rh, fill=WHITE, line=LINE)
        rect(s, 64, y, BAR, rh, fill=th.main)
        text(s, 94, y + 20, 200, 60, no, size=46, bold=True, color=th.main, font=MONO, lh=58, wrap=False)
        text(s, 94, y + 86, 230, 76, name, size=21, bold=True, color=INK, lh=30)
        text(s, 340, y, 850, rh, txt, size=20, color=BODY, anchor='m', lh=31)
        tri(s, 1222, y + rh / 2.0, 18, 26, TRI)
        oy = y + (rh - (len(outs) * 44 + (len(outs) - 1) * 10)) / 2.0
        for nm, form in outs:
            rect(s, 1250, oy, 590, 44, fill=th.bg)
            text(s, 1266, oy, 420, 44, nm, size=19, bold=True, color=INK, anchor='m', lh=26, wrap=False)
            text(s, 1660, oy, 164, 44, form, size=19, bold=True, color=th.text, align='r', anchor='m', lh=26,
                 wrap=False)
            oy += 54
        y += rh + 10


# =====================================================================  20 量化指标与效益评估
def s20_metrics(d):
    s = d.new_slide()
    frame(d, s, '量化指标与效益评估：成果交付 ＋ 示范验证 ＋ 对比评估', '成果数量 · 能力覆盖与效率提升目标 · 效益评估方式与经济效益目标', '预期成果与效益')
    L, LW = 64, 872
    header(s, L, 212, '量化指标', note='成果数量 ＋ 能力与效率目标', w=LW)
    kp = [('软件', '1', '套'), ('模型', '1', '套'), ('系统', '1', '套'), ('规范', '1', '套'), ('报告', '1', '份'),
          ('专利', '4', '项'), ('论文', '2', '篇'), ('标准草案', '2', '项')]
    tw_, th_ = (LW - 3 * 12) / 4.0, 100
    for i, (lab_, v, u) in enumerate(kp):
        x = L + (i % 4) * (tw_ + 12)
        y = 256 + (i // 4) * (th_ + 12)
        rect(s, x, y, tw_, th_, fill=GRAYBG)
        text(s, x + 22, y + 10, 150, 56, '%s<s=19><n><c=%s> %s</c></n></s>' % (v, SUB, u), size=44, bold=True,
             color=KPI_COLOR if i >= 5 else INK, font=MONO, lh=56, wrap=False)
        text(s, x + 22, y + 66, tw_ - 30, 26, lab_, size=19, bold=True, color=SUB, lh=26, wrap=False)
    # 能力与效率目标：定性描述 + 目标值占位
    rows = [
        ('能力覆盖', '支持<b>异构</b>仿真智能体接入；支持速度跟踪控制品质、牵引能耗与再生能量利用、运行平稳性等关键性能的<b>多工况</b>协同验证',
         [('接入异构仿真智能体', '≥【待填】类'), ('协同验证典型工况', '≥【待填】个')]),
        ('效率提升', '关键性能验证的<b>效率与工况覆盖度显著提升</b>，减少现场验证投入（现状一次软件迭代：研发 20＋测试 10＋现场 10 人）',
         [('验证周期（现状约 2 个月）缩短', '≥【待填】%'), ('工况覆盖度提升', '≥【待填】%')]),
        ('扩展能力', '形成<b>可扩展</b>的智能体接入机制，支持<b>智能运维</b>应用，具备向制动、辅助系统等关键系统扩展的能力', []),
    ]
    y = 256 + 2 * th_ + 12 + 24
    y_end = 256 + 4 * 144 + 3 * 12
    tx, txw = L + 130, LW - 130
    ns = [wrap_lines(txt, txw, 20) for _, txt, _ in rows]
    base = [n * 31 + (46 if tg else 0) for n, (_, _, tg) in zip(ns, rows)]
    pad = (y_end - y - sum(base)) / float(len(rows))
    for (lab_, txt, tg), n, b in zip(rows, ns, base):
        h = b + pad
        line(s, L, y, L + LW, y, LINE, 1)
        top = y + (h - b) / 2.0
        tag(s, L, top - 1, lab_, fill=M1.bg, color=M1.text, size=18, h=32, padx=12)
        text(s, tx, top, txw, n * 31 + 4, txt, size=20, color=BODY, lh=31)
        xx = tx
        for lb, v in tg:
            xx += target_chip(s, xx, top + n * 31 + 12, lb, v) + 10
        y += h
    line(s, L, y, L + LW, y, LINE, 1)
    # 右：效益评估
    R, RW = 976, 880
    header(s, R, 212, '效益评估', note='效益 · 评估方式 · 目标', w=RW)
    bens = [
        ('技术效益', 'bulb', M1, '形成列车级数字样机构建方法体系与技术规范，填补多智能体协同调度的领域化应用空白', '方法体系与规范交付、示范应用验证', []),
        ('经济效益', 'coins', M3, '减少实物试验与样车试制投入，缩短研制迭代周期', '试验次数与周期对比评估',
         [('减少实物试验', '≥【待填】次/年'), ('节约试验与试制费用', '≥【待填】万元/年')]),
        ('管理效益', 'chart-bar', M2, '提升关键性能验证的效率与工况覆盖度，支撑研发流程标准化', '验证效率与覆盖度对比', []),
        ('推广效益', 'world', M4, '方法体系支持智能运维并具备向制动、辅助系统扩展能力，可在**内推广', '扩展性验证与推广方案', []),
    ]
    y = 256
    for t, ic, th, txt, how, tg in bens:
        bh = 180 if tg else 132
        rect(s, R, y, RW, bh, fill=WHITE, line=LINE)
        rect(s, R, y, BAR, bh, fill=th.main)
        icon(s, ic, th.main, R + 26, y + 20, 32)
        text(s, R + 70, y + 16, 200, 40, t, size=22, bold=True, color=INK, anchor='m', lh=30, wrap=False)
        text(s, R + 26, y + 58, RW - 52, 30, txt, size=19, color=BODY, lh=28, wrap=False)
        text(s, R + 26, y + 92, RW - 52, 28, '<k>评估方式</k>　' + how, size=18, color=th.text, lh=26, wrap=False)
        if tg:
            text(s, R + 26, y + 128, 100, 34, '<k>量化目标</k>', size=18, color=th.text, anchor='m', lh=26, wrap=False)
            xx = R + 26 + text_width('量化目标', 18, True) + 18
            for lb, v in tg:
                xx += target_chip(s, xx, y + 128, lb, v) + 10
        y += bh + 12
    summary(s, 64, 908, 1792, 84, '效益评估以“<r>成果交付＋示范验证＋对比评估</r>”三层支撑，每项成果均对应一项研究目标。', title_size=27)
    notes(s, '说明：虚线框内【待填】为量化目标占位，请按课题实际测算后填写（接入智能体类数、协同验证工况数、验证周期缩短比例、'
             '工况覆盖度提升比例、实物试验减少次数、节约费用）。')


# =====================================================================  21 经费预算


def s21_budget(d):
    s = d.new_slide()
    frame(d, s, '经费预算：总预算100万元，算力是核心投入且与任务强相关', '预算明细 · 分年度 · 合理性说明——无对外技术合作费支出，全部经费用于自主研发', '经费预算合理性')
    L, LW = 64, 1064
    # 总额 + 原生堆叠条
    text(s, L, 206, 240, 80, '100<s=24><n><c=%s> 万元</c></n></s>' % SUB, size=60, bold=True, color=INK, font=MONO,
         lh=76, wrap=False)
    text(s, L, 280, 240, 28, '总预算', size=18, bold=True, color=MUTED, lh=24, wrap=False)
    stacked_bar(s, L + 250, 226, LW - 250, 56, [(v, c, lab, WHITE) for _, _, v, _, c, lab in BUDGET], size=19)
    text(s, L + 250, 288, LW - 250, 26, '折旧 · 外协 · 成果 · 其他 各 5 万元（灰色段）', size=16, color=SUB, align='r', lh=22,
         wrap=False)
    # 明细表
    rows = []
    for no, name, v, basis, col, _ in BUDGET:
        rows.append([no, name, '%d' % v, '%d%%' % v, basis])
    rows.append(['', '<b>合计</b>', '<b>100</b>', '<b>100%</b>', '—'])
    cols = [64, 262, 124, 180, 434]
    table(s, L, 332, cols, rows, header=['序号', '科目', '金额（万元）', '占比', '测算依据'], size=19, lh=28, pad_x=14,
               pad_y=20, head_size=16,
               col_styles={0: dict(align='c', font=MONO, bold=True, color=SUB), 1: dict(bold=True, color=INK),
                           2: dict(align='c', font=MONO, bold=True, color=INK), 3: dict(align='r', font=MONO, color=SUB),
                           4: dict(color=BODY, size=18, lh=27)},
               fills={6: PANEL})
    # 占比列的迷你条 + 色标
    y = 332 + 44
    for i, (no, name, v, basis, col, _) in enumerate(BUDGET):
        n = max(wrap_lines(basis, cols[4] - 28, 18) * 27, wrap_lines(name, cols[1] - 28, 19, True) * 28)
        rh = n + 40
        bx = L + sum(cols[:3]) + 14
        rect(s, bx, y + rh / 2.0 - 7, 100 * v / 50.0, 14, fill=col)
        rect(s, L + 6, y + rh / 2.0 - 7, 4, 14, fill=col)
        y += rh
    # 右：分年度 + 合理性
    R, RW = 1168, 688
    header(s, R, 212, '分年度', note='投入均衡，与四阶段进度匹配', w=RW)
    tw_ = (RW - 16) / 2.0
    for i, (yr, amt) in enumerate([('2027 年', '50'), ('2028 年', '50')]):
        x = R + i * (tw_ + 16)
        rect(s, x, 256, tw_, 96, fill=GRAYBG)
        text(s, x + 22, 256, 140, 96, yr, size=22, bold=True, color=SUB, anchor='m', lh=30, wrap=False)
        text(s, x + 140, 256, tw_ - 160, 96, '%s<s=20><n><c=%s> 万元</c></n></s>' % (amt, SUB), size=44, bold=True,
             color=INK, font=MONO, align='r', anchor='m', lh=56, wrap=False)
    header(s, R, 380, '合理性说明', w=RW)
    bullets(s, R, 424, RW, 560, [
        '<b>算力是核心投入且与任务强相关</b>：NPU 承担智能体推理与任务编排决策，GPU 承担多物理域模型并行仿真，<br>二者协同满足<b>实时协同运行与保真度切换</b>要求',
        '<b>人员占比 30%</b>：以算法与软件开发为主，属智力密集型研究，占比合理',
        '<b>折旧 5 万</b>：充分利用现有仿真与测试设备，避免重复购置',
        '<b>外协 5 万</b>：用于建模数据采集与第三方验证，保障数据质量与结果客观性',
        '<b>成果 5 万</b>：支撑专利、论文与标准草案产出；<b>其他 5 万</b>：差旅与会议，用于示范应用与协同交流',
        '<r>无对外技术合作费支出</r>，全部经费用于自主研发；分年度投入均衡，与四阶段进度匹配',
    ], M1.main, size=20, lh=31, sa=20)
    notes(s, '说明：人员人工费用的测算依据中【待填】为占位，请按“投入人月 × 人月费用标准”填写，合计应为 30 万元。')


# =====================================================================  22 项目管理计划
def s22_management(d):
    s = d.new_slide()
    frame(d, s, '项目管理计划：按四层分组承担，按里程碑绑定考核', '任务分配 · 项目流程 · 时间安排 · 执行保障', '项目管理计划')
    # 任务分配：组织架构
    header(s, 64, 212, '组织架构与任务分配', w=900)
    bx, bw = 960 - 170, 340
    rect(s, bx, 214, bw, 58, fill=INK)
    icon(s, 'user-star', WHITE, bx + 70, 214 + 14, 30)
    text(s, bx + 110, 214, bw - 130, 58, '项目负责人', size=23, bold=True, color=WHITE, anchor='m', lh=30, wrap=False)
    groups = [
        (1, '判据知识化、工作流建模、知识与流程迭代', '判据知识库、工作流定义'),
        (2, '任务解析、环境装配、保真度与算力调度', '协同运行与调度引擎'),
        (3, '智能体研制、用例生成、识别与报告、数据回归', '仿真智能体／孪生模型、验证报告'),
        (4, '架构设计、协同运行、沙箱、样机集成', '列车级数字样机、方法体系与规范'),
    ]
    gw = (1792 - 3 * 24) / 4.0
    gy, gh = 322, 232
    line(s, 960, 272, 960, 296, MUTED, 1.5)
    line(s, 64 + gw / 2.0, 296, 64 + 3 * (gw + 24) + gw / 2.0, 296, MUTED, 1.5)
    for i, (k, task, deli) in enumerate(groups):
        th = LAYER[k]
        x = 64 + i * (gw + 24)
        line(s, x + gw / 2.0, 296, x + gw / 2.0, gy - 2, MUTED, 1.5, tail='triangle')
        rect(s, x, gy, gw, gh, fill=WHITE, line=LINE)
        rect(s, x, gy, gw, BAR, fill=th.main)
        tag(s, x + 22, gy + 20, '子任务%d组' % k, fill=th.main, size=17, h=30, padx=12)
        text(s, x + 22 + text_width('子任务%d组' % k, 17, True) + 38, gy + 20, gw - 160, 30, '对应' + LAYER_NO[k],
             size=17, bold=True, color=th.text, anchor='m', lh=24, wrap=False)
        text(s, x + 22, gy + 60, gw - 44, 32, LAYER_NAME[k], size=22, bold=True, color=INK, lh=30, wrap=False)
        text(s, x + 22, gy + 104, 120, 24, '主要任务', size=16, bold=True, color=MUTED, lh=22, wrap=False)
        text(s, x + 22, gy + 128, gw - 44, 28, task, size=17, color=BODY, lh=26)
        concl(s, x + 16, gy + gh - 16 - 46, gw - 32, 46, '交付：' + deli, th, size=17, pad=12)
    y1 = gy + gh + 28
    # 项目流程：管理节奏（阶梯）
    L, LW = 64, 876
    header(s, L, y1, '项目管理流程', note='双周 → 月度 → 季度 → 年度', w=LW)
    cad = [('双周', '双周例会', '子任务进展同步与问题协调'), ('月度', '月度节点检查', '对照里程碑核查交付物'),
           ('季度', '季度评审', '技术方案评审与阶段成果确认'), ('年度', '年度总结', '目标达成评估与下年度计划调整')]
    cw = (LW - 3 * 12) / 4.0
    for i, (f, n, t) in enumerate(cad):
        x = L + i * (cw + 12)
        hh = 226 + i * 40
        top = 990 - hh
        rect(s, x, top, cw, hh, fill=M1.bg if i == 3 else PANEL, line=LINE)
        rect(s, x, top, cw, STYLE['concl_bar'], fill=M1.main)
        oval(s, x + 20, top + 24, 60, 60, fill=M1.soft)
        text(s, x + 20, top + 24, 60, 60, f, size=18, bold=True, color=M1.text, align='c', anchor='m', lh=24,
             wrap=False)
        text(s, x + 20, top + 98, cw - 40, 34, n, size=22, bold=True, color=INK, lh=30, wrap=False)
        text(s, x + 20, top + 138, cw - 36, 80, t, size=19, color=BODY, lh=28)
    # 时间安排：按实际月数等比例的阶段条 + 里程碑
    R, RW = 980, 876
    header(s, R, y1, '时间安排与执行保障', note='里程碑绑定阶段交付物，按节点考核', w=RW)
    mw = RW / 24.0
    sy, sh_ = y1 + 50, 58
    for name, short, a, b, col, per, goal, deli, mk, mdate, mtext in STAGES:
        x = R + a * mw + (1 if a else 0)
        w = (b - a) * mw - (1 if a else 0) - (1 if b < 24 else 0)
        fg = STAGE_ON[STAGE.index(col)] if STAGE_ON else on_color(col)
        rect(s, x, sy, w, sh_, fill=col)
        p0, p1 = per.split('—')
        per_s = p0 + '—' + (p1[5:] if p1[:4] == p0[:4] else p1)
        text(s, x, sy, w, sh_, [dict(t=name, size=17, bold=True, lh=24),
                                  dict(t=per_s, size=14, font=MONO, bold=False, lh=20)],
             color=fg, align='c', anchor='m', wrap=False)
        lw_ = max(text_width('%s · %s' % (mk, mdate), 17, True), text_width(mtext, 15)) + 8
        cx = min(x + w / 2.0, R + RW - lw_ / 2.0)  # 末段较窄，标签不越出右边界
        text(s, cx - 80, sy + sh_ + 12, 160, 26, '%s · %s' % (mk, mdate), size=17, bold=True, color=RED, align='c',
             lh=24, wrap=False)
        text(s, cx - 80, sy + sh_ + 40, 160, 24, mtext, size=15, color=SUB, align='c', lh=22, wrap=False)
    # 执行保障
    ctl = [('质量控制', 'shield-check', '内部评审 ＋ 第三方验证，交付物可核查'),
           ('风险管控', 'alert-triangle', '技术风险（模型精度、实时性、判据知识化完备性）、进度风险、接口风险，均设置应对预案'),
           ('资源保障', 'server', '算力资源、试验与运用数据、样车与试验台资源')]
    y = sy + sh_ + 84
    hs = [52, 86, 52]
    for (t, ic, body), h in zip(ctl, hs):
        line(s, R, y, R + RW, y, LINE, 1)
        icon(s, ic, RED if t == '风险管控' else INK, R, y + 12, 28)
        text(s, R + 40, y + 8, 200, 36, t, size=20, bold=True, color=INK, lh=28, wrap=False)
        text(s, R + 230, y + 10, RW - 230, h - 12, body, size=18, color=BODY, lh=27)
        y += h + 6
    line(s, R, y - 6, R + RW, y - 6, LINE, 1)


# =====================================================================  23 团队研究能力
def s23_team(d):
    s = d.new_slide()
    frame(d, s, '团队研究能力：为何由我单位承担', '牵头单位：********** · 研发平台 · 仿真能力 · 试验能力 · 数据积累 · 在研项目 · 人才队伍', '团队研究能力')
    items = [
        ('研发平台', 'building-factory-2', '【待填：机车总体设计／控制系统开发等平台】'),
        ('仿真能力', 'device-desktop-analytics', '【待填：已有的车辆动力学、牵引电传动、控制等仿真模型与软件】'),
        ('试验能力', 'test-pipe', '现有硬件在环（HIL）试验台，结合经验数据模拟运用工况；【待填：其他试验台架】'),
        ('数据积累', 'database', '【待填：机车试验数据与运用数据积累情况】'),
        ('在研项目', 'clipboard-list', '【待填：**级****项目数量与名称】'),
        ('人才队伍', 'users-group', '【待填：参研人员规模与职称结构】'),
    ]
    cw, ch = (1792 - 2 * 24) / 3.0, 232
    for i, (t, ic, ph) in enumerate(items):
        x = 64 + (i % 3) * (cw + 24)
        y = 214 + (i // 3) * (ch + 22)
        rect(s, x, y, cw, ch, fill=WHITE, line=LINE)
        icon_badge(s, ic, M1, x + 26, y + 24, 64)
        text(s, x + 108, y + 24, cw - 130, 64, t, size=26, bold=True, color=INK, anchor='m', lh=34, wrap=False)
        rect(s, x + 26, y + 108, cw - 52, ch - 132, fill=PH.bg, line=PH.main, lw=1.2, dash='dash')
        text(s, x + 46, y + 108, cw - 92, ch - 132, ph, size=19, color=PH.text, anchor='m', lh=28)
    # 匹配性
    y1 = 724
    header(s, 64, y1, '与本课题的匹配性', w=1792)
    mt = [('控制系统与车辆动力学为牵头单位主营业务方向', '【可补充：具体型号或项目实例】'),
          ('具备支撑智能体研制、引擎开发与示范应用的仿真与试验条件', '【可补充：具体平台与台架】')]
    mw = (1792 - 24) / 2.0
    for i, (a, b) in enumerate(mt):
        x = 64 + i * (mw + 24)
        rect(s, x, y1 + 50, mw, 222, fill=OK.bg)
        rect(s, x, y1 + 50, STYLE['concl_bar'], 222, fill=OK.main)
        icon(s, 'circle-check', OK.main, x + 28, y1 + 76, 34)
        text(s, x + 76, y1 + 72, mw - 104, 70, a, size=23, bold=True, color=INK, lh=32)
        rect(s, x + 76, y1 + 160, mw - 110, 72, fill=PH.bg, line=PH.main, lw=1.2, dash='dash')
        text(s, x + 96, y1 + 160, mw - 150, 72, b, size=19, color=PH.text, anchor='m', lh=28)
    notes(s, '说明：本页需填入单位实际基础，它是“为何由我单位承担”的直接依据。\n虚线框内【待填】/【可补充】为占位内容，请替换为牵头单位实际情况。')


# =====================================================================  24 成果介绍
def s24_results(d):
    s = d.new_slide()
    frame(d, s, '成果介绍：五项成果 ＋ 知识产权与标准', '调度引擎 · 仿真智能体／智能孪生模型 · 列车级数字样机 · 方法体系与技术规范 · 示范应用与验证报告', '成果介绍')
    top = [
        ('成果1', '软件1套', '多智能体协同运行与调度引擎', M1, 'settings-automation', [
            '完成<b>任务理解—任务编排—智能体执行</b>全流程：验证目标解析、环境自装配、保真度与算力调度、执行监控与重规划',
            '支持异构仿真智能体统一接入，<br>解决多专业模型“接不上”的问题',
            '面向性能验证提供<b>实时协同运行能力</b>'],
         ('调度引擎软件', '建议：软件界面截图或系统架构图')),
        ('成果2', '1套', '关键系统仿真智能体／<br>智能孪生模型', M2, 'robot', [
            '覆盖控制系统、车辆动力学与线路等关键系统',
            '由“被动模型”升级为“<b>自主智能体</b>”，具备参数自动标定与结果判断能力',
            '具备标准化接口与可扩展接入机制，支持后续新增系统'],
         ('仿真智能体／孪生模型', '建议：模型结构图或仿真运行界面截图')),
        ('成果3', '1套', '列车级数字样机', M3, 'train', [
            '集成关键系统仿真智能体，形成<b>整车级验证环境</b>',
            '支持列车关键性能的多工况协同验证，并支撑<b>智能运维</b>应用',
            '预留制动、辅助系统等关键系统扩展接口'],
         ('列车级数字样机', '建议：数字样机三维模型或运行界面截图')),
    ]
    pw, ph, y0 = (1792 - 48) / 3.0, 470, 214
    for i, (no, form, name, th, ic, bl, (ft, hint)) in enumerate(top):
        x = 64 + i * (pw + 24)
        rect(s, x, y0, pw, ph, fill=WHITE, line=LINE)
        rect(s, x, y0, pw, BAR, fill=th.main)
        icon_badge(s, ic, th, x + 24, y0 + 26, 56)
        text(s, x + 96, y0 + 22, 300, 26, '<k>%s</k>　<c=%s>%s</c>' % (no, SUB, form), size=17, color=th.text, lh=24,
             wrap=False)
        text(s, x + 96, y0 + 48, pw - 96 - 24 - 80, 62, name, size=23, bold=True, color=INK, lh=30)
        num_mark(s, x + pw - 24 - 80, y0 + 18, 80, '0%d' % (i + 1), th.num, size=36)
        fig_placeholder(s, x + 24, y0 + 122, pw - 48, 160, ft, hint)
        bullets(s, x + 24, y0 + 300, pw - 48, ph - 308, bl, th.main, size=18, lh=27, sa=8)
    # 成果4、5
    y1 = y0 + ph + 22
    LW = 920
    small = [('成果4', '1套', '列车级数字样机构建方法体系与技术规范', M4, 'file-certificate',
              '形成构建方法、集成流程与标准化接口规范，可作为**内推广与后续课题复用的方法支撑'),
             ('成果5', '1份', '示范应用与验证报告', M5, 'report-analytics', '在 FXN5C 机车上完成关键性能协同验证，验证场景为临哈线典型区段，形成验证效能评估结论')]
    sh = (996 - y1 - 14) / 2.0
    for i, (no, form, name, th, ic, txt) in enumerate(small):
        y = y1 + i * (sh + 14)
        rect(s, 64, y, LW, sh, fill=WHITE, line=LINE)
        rect(s, 64, y, BAR, sh, fill=th.main)
        icon_badge(s, ic, th, 90, y + (sh - 60) / 2.0, 60)
        n = wrap_lines(txt, LW - 200, 19)
        top_ = y + (sh - (34 + 6 + n * 28)) / 2.0
        text(s, 172, top_, LW - 200, 34, '<k><c=%s>%s</c></k>　%s　<c=%s>%s</c>' % (th.text, no, name, SUB, form),
             size=21, bold=True, color=INK, lh=30, wrap=False)
        text(s, 172, top_ + 40, LW - 200, n * 28 + 4, txt, size=19, color=BODY, lh=28)
    # 知识产权与标准
    R, RW = 1008, 848
    header(s, R, y1 - 4, '知识产权与标准', w=RW)
    ip = [('4', '项', '发明专利申请', '覆盖任务编排调度机制、保真度与算力调度、实时时序同步、工况自生成等方向'),
          ('2', '篇', '核心论文', '面向多智能体协同仿真与性能验证方法'),
          ('2', '项', '企业标准草案', '列车级数字样机构建与智能体接口相关规范')]
    kw_ = (RW - 2 * 12) / 3.0
    for i, (v, u, lab_, sub) in enumerate(ip):
        x = R + i * (kw_ + 12)
        kpi(s, x, y1 + 40, kw_, 996 - y1 - 40, '%s<s=20><n><c=%s> %s</c></n></s>' % (v, SUB, u), lab_, sub, color=KPI_COLOR,
            value_size=46, label_size=20, sub_size=16, pad=18)
    notes(s, '说明：成果 1—3 中的虚线框为附图预留位置（形状组合）；选中该组合删除后插入实际图片即可。')


# =====================================================================  25 结束页
def s25_end(d):
    s = d.new_slide()
    rect(s, 64, 34, 1792, STYLE['rule_w'], fill=RULE)
    rect(s, 64, 132, 8, 30, fill=RED)
    text(s, 88, 128, 1400, 38, '****“***”科技重大专项课题申报', size=25, bold=True, color=RED, cs=3, anchor='m',
         lh=34, wrap=False)
    text(s, 64, 262, 1792, 140, '恳请各位专家指导', size=104, bold=True, color=INK, lh=134, wrap=False)
    rect(s, 64, 432, STYLE['title_bar'][0] * 1.25, STYLE['title_bar'][1] + 2, fill=RED)
    text(s, 64, 476, 1792, 50, '**********', size=34, bold=True, color=INK, lh=46, wrap=False)
    text(s, 64, 538, 1792, 46, '7.3 轨道交通装备性能验证的多智能体协同运行与调度引擎技术研究', size=28, color=SUB, lh=40, wrap=False)
    ruler(s, 64, 944, 1792)
    loco_train(s, 560, 944, 428, n=3, gap=8, faded=False)
    rect(s, 64, 1010, 1792, STYLE['rule_w'], fill=RULE)
    sep = '<c=%s>   |   </c>' % RULE
    text(s, 64, 1024, 1300, 36, FOOTER + sep + '课题申报汇报' + sep + '谢谢', size=17, color=MUTED, bold=True, cs=2.5,
         anchor='m', lh=24, wrap=False)
    text(s, 1456, 1024, 400, 36, '%02d / %02d' % (s._no, d.total), size=17, color=MUTED, bold=True, font=MONO,
         align='r', anchor='m', lh=24, wrap=False)

SLIDES = [s01_cover, s02_agenda, s03_intro, s04_value, s05_urgency, s06_future, s07_core, s08_layer1,
          s09_layer2, s10_layer3, s11_layer4, s12_methods, s13_route, s14_scenario, s15_innovation,
          s16_benchmark, s17_schedule, s18_tasks, s19_goals, s20_metrics, s21_budget, s22_management,
          s23_team, s24_results, s25_end]


def build(out=OUT, only=None):
    d = Deck(TOTAL, FOOTER)
    for fn in SLIDES:
        fn(d)
    cp = d.prs.core_properties
    cp.title = '7.3 轨道交通装备性能验证的多智能体协同运行与调度引擎技术研究 · 课题申报汇报'
    cp.subject = '课题申报汇报'
    cp.author = ''
    cp.last_modified_by = ''
    cp.keywords = '多智能体; 协同运行; 调度引擎; 数字样机; 性能验证'
    d.save(out)
    for w in pptkit.WARNINGS:
        print('⚠', w)
    print('saved', out, 'slides:', d.slide_no)


if __name__ == '__main__':
    build(sys.argv[1] if len(sys.argv) > 1 else OUT)
