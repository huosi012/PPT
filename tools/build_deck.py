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
LOCO = os.path.join(ROOT, 'assets', 'illustrations', 'locomotive.png')
LOCO_FADED = os.path.join(ROOT, 'assets', 'illustrations', 'locomotive_faded.png')
TOTAL = 25
FOOTER = '7.3 多智能体协同运行与调度引擎'


# 经费明细（第 3、21 页共用）：序号、科目、金额、测算依据、色块、条内标签
BUDGET = [
    ('1', '直接投入费用', 50, 'AI算力服务器（GPU+NPU异构）1台 40万；仿真与调度一体化工作站 2台 10万', BLUE.main, '直接投入 50'),
    ('2', '人员人工费用', 30, '项目参研人员薪酬费用，按投入人月测算', '4A78FF', '人员 30'),
    ('3', '固定资产相关费用（折旧）', 5, '现有仿真与测试设备折旧分摊', '56637A', None),
    ('4', '试验检验及试制外协费用', 5, '建模数据采集处理 2万；仿真模型校验与第三方验证 3万', '727F95', None),
    ('5', '研发成果相关费用', 5, '发明专利申请 2万；论文版面及标准草案编制 3万', '8E9AAE', None),
    ('6', '与研发活动直接相关的<br>其他费用', 5, '差旅费 3万；会议费 2万', 'AAB5C5', None),
]


def loco_train(s, x, y_rail, w_each, n=3, gap=10, faded=True):
    """在标尺线（轨道）上摆放 n 节机车插画。y_rail 为轨面位置。"""
    h = w_each * 0.25
    for i in range(n):
        pic(s, LOCO_FADED if faded else LOCO, x + i * (w_each + gap), y_rail - h * 0.992, w_each, h)


# =====================================================================  01 封面
def s01_cover(d):
    s = d.new_slide()
    rect(s, 64, 34, 1792, 2, fill=RULE)
    rect(s, 64, 132, 8, 30, fill=RED)
    text(s, 88, 128, 1400, 38, '****“***”科技重大专项课题申报', size=25, bold=True, color=RED, cs=3, anchor='m',
         lh=38, wrap=False)
    text(s, 64, 196, 1792, 230, ['<r>7.3</r> 轨道交通装备性能验证的', '多智能体协同运行与调度引擎技术研究'],
         size=78, bold=True, color=INK, lh=110)
    rect(s, 64, 446, 120, 8, fill=RED)
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
    rect(s, 64, 1010, 1792, 2, fill=RULE)
    sep = '<c=%s>   |   </c>' % RULE
    text(s, 64, 1024, 1300, 36, FOOTER + sep + '课题申报汇报' + sep + '研究周期 2027.01—2028.12', size=17,
         color=MUTED, bold=True, cs=2.5, anchor='m', lh=36, wrap=False)


# =====================================================================  02 提纲
def s02_agenda(d):
    s = d.new_slide()
    frame(d, s, '汇报提纲：七个部分，核心是研究内容、技术路线与方法', '共 25 页 · 第三部分“主要研究内容、技术路线与方法”为核心章节（第 6—17 页）',
          '汇报提纲')
    rows = [
        ('01', '项目简介', '基本信息 · 在项目7中的位置 · 主要成果物 · 经费概览', '第 3 页', 'file-description'),
        ('02', '研究意义与价值', '背景 · 目的 · 价值 · 效益 ｜ 立项必要性与紧迫性', '第 4—5 页', 'bulb'),
        ('03', '主要研究内容、技术路线与方法',
         '核心命题 · 研究内容①—④ · 研究方法 · 技术路线 · 典型验证场景 · 创新内核 · 技术对标 · 未来价值 · 阶段进度',
         '第 6—17 页', 'route'),
        ('04', '预期成果与效益', '任务分解与考核方式 · 预期目标与成果对应 · 量化指标与效益评估', '第 18—20 页', 'target-arrow'),
        ('05', '经费预算合理性', '预算明细 · 分年度 · 合理性说明', '第 21 页', 'coins'),
        ('06', '项目管理计划', '组织 · 分配 · 流程 · 控制 ｜ 团队研究能力', '第 22—23 页', 'sitemap'),
        ('07', '成果介绍', '五项成果 ＋ 知识产权与标准', '第 24 页', 'trophy'),
    ]
    y0, rh = 222, 110
    for i, (no, title, sub, pages, ic) in enumerate(rows):
        y = y0 + i * rh
        core = (no == '03')
        if core:
            rect(s, 64, y + 4, 1792, rh - 8, fill=RED_BG)
            rect(s, 64, y + 4, 5, rh - 8, fill=RED)
        text(s, 92, y, 110, rh, no, size=46, bold=True, color=RED if core else 'C9D2DF', font=MONO, anchor='m',
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
    frame(d, s, '项目简介：研究重心是引擎，数字样机是引擎的集成产物', '基本信息 · 在项目7中的位置与边界 · 主要成果物 · 经费概览', '项目简介')
    # ---- 左：基本信息
    L, LW = 64, 840
    header(s, L, 212, '基本信息', w=LW)
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
    header(s, L, y1, '主要成果物', note='软件 · 模型 · 系统 · 规范 · 报告 · 知识产权', w=LW)
    items = [
        ('多智能体协同运行与调度引擎', '软件1套', BLUE, 'settings-automation'),
        ('关键系统仿真智能体／<br>智能孪生模型', '模型1套', PURPLE, 'robot'),
        ('列车级数字样机', '系统1套', GREEN, 'train'),
        ('列车级数字样机构建方法体系与技术规范', '规范1套', GREEN, 'file-certificate'),
        ('示范应用与验证报告', '报告1份', ORANGE, 'report-analytics'),
        ('发明专利／论文／<br>企业标准草案', '4项／2篇／2项', ORANGE, 'certificate'),
    ]
    cw, ch, gx, gy = (LW - 24) / 3.0, 116, 12, 12
    for i, (name, form, th, ic) in enumerate(items):
        cx = L + (i % 3) * (cw + gx)
        cy = y1 + 46 + (i // 3) * (ch + gy)
        rect(s, cx, cy, cw, ch, fill=WHITE, line=LINE)
        rect(s, cx, cy, cw, 4, fill=th.main)
        icon(s, ic, th.main, cx + 16, cy + 18, 30)
        tag(s, cx + cw - 16 - (text_width(form, 16, True) + 20), cy + 18, form, fill=th.bg, color=th.text, size=16,
            h=28, padx=10)
        text(s, cx + 16, cy + 56, cw - 28, 52, name, size=18, bold=True, color=INK, lh=25)
    # ---- 左：经费概览（堆叠条）
    y2 = y1 + 46 + 2 * ch + gy + 22
    header(s, L, y2, '经费概览', note='总预算 100 万元', w=LW)
    by = y2 + 48
    stacked_bar(s, L, by, LW, 40, [(v, c, lab, WHITE) for _, _, v, _, c, lab in BUDGET])
    unit = (LW - 2 * 5) / 100.0
    text(s, L, by + 46, 50 * unit, 28, 'GPU+NPU 异构算力服务器 40 · 一体化工作站 10', size=16, color=SUB, lh=28,
         wrap=False)
    text(s, L + 440, by + 46, LW - 440, 28, '人员 30 · 折旧 · 外协 · 成果 · 其他 各 5', size=16, color=SUB,
         align='r', lh=28, wrap=False)

    # ---- 右：在项目7中的位置与边界
    R, RW = 944, 912
    header(s, R, 212, '在项目7中的位置与相邻课题边界', note='项目7设5个子课题 · 研究重心是“引擎”', w=RW)
    rows = [
        ('7.1', '高质量数据集', '边界：仅采用其数据规范与接口，不承担构建（<b>7.1 建库、7.3 按需生成</b>）', 'n'),
        ('7.2', '多物理域建模与仿真智能体', '边界：为验证需要自建控制系统、车辆动力学与线路等必要智能体，<b>不覆盖弓网/牵引/制动等全部关键系统</b>',
         'n'),
        ('7.3', '协同运行与调度引擎（本课题）',
         '<b>核心攻关多智能体协同运行与调度引擎</b>，<br>解决多专业智能体“接得上、跑得动、调得灵”的问题；<br>'
         '<b>列车级数字样机是引擎承载智能体后形成的集成产物</b>', 'me'),
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
    frame(d, s, '研究意义与价值：实现列车关键性能的快速、可复现、全覆盖验证', '背景与必要性 → 研究目的 → 应用价值与推动作用 · 潜在社会经济效益', '研究意义与价值')
    X, XW = 300, 1556
    side_label(s, 64, 222, 206, '背景与必要性', '实物试验为主、模型各自为战，难以支撑列车级性能验证')
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
    rect(s, X, 414, 5, 88, fill=RED)
    text(s, X + 30, 414, XW - 60, 88,
         '突破多智能体协同运行与调度引擎技术，构建包含关键系统仿真智能体的列车级数字样机，实现列车关键性能的<r>快速、可复现、全覆盖</r>验证。',
         size=23, bold=True, color=INK, anchor='m', lh=36)
    # 价值与效益
    side_label(s, 64, 540, 206, '价值与效益', '应用价值与推动作用<br>潜在社会经济效益')
    y0, ph = 528, 468
    pw = (XW - 32) / 2.0
    yy = panel(s, X, y0, pw, ph, BLUE, '应用价值与对领域发展的推动作用', '01', 'trending-up', title_size=28, badge=76)
    bullets(s, X + 40, yy + 20, pw - 80, 210, [
        '应用于<b>机车研发验证与运维测试</b>，覆盖速度跟踪控制品质、牵引能耗与再生能量利用、运行平稳性等关键性能；支撑复杂工况的<b>低成本、高频次</b>验证',
        '推动研发验证从<b>单系统、离线</b>向<b>列车级、实时协同、智能调度</b>演进；将多智能体协同从通用平台层引入<b>装备性能验证</b>领域',
    ], BLUE.main, size=20, lh=31, sa=12)
    concl(s, X + 40, y0 + ph - 32 - 72, pw - 80, 72, '输出列车级数字样机构建方法体系与技术规范，<br>为行业提供可复用的方法支撑', BLUE,
          size=20)
    x2 = X + pw + 32
    yy = panel(s, x2, y0, pw, ph, GREEN, '潜在社会经济效益', '02', 'coins', title_size=30, badge=76)
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
        rect(s, x2 + 40, ey, 56, 56, fill=GREEN.bg)
        icon(s, ic, GREEN.main, x2 + 40 + 12, ey + 12, 32)
        text(s, x2 + 116, ey - 2, 200, 30, lab_, size=21, bold=True, color=GREEN.text, lh=30, wrap=False)
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
    text(s, R + 470, 258, 300, 30, '目标形态', size=17, bold=True, color=GREEN.text, lh=30, wrap=False)
    gy = 292
    for dim, now, goal in gaps:
        text(s, R, gy, 140, 50, dim, size=20, bold=True, color=INK, anchor='m', lh=30, wrap=False)
        chip(s, R + 140, gy, now, color=SUB, fill=GRAYBG, line=None, size=19, h=50, w=280, align='l', radius=0)
        tri(s, R + 440, gy + 25, 16, 22, TRI)
        chip(s, R + 462, gy, goal, color=GREEN.text, fill=GREEN.bg, line=None, size=19, h=50, w=RW - 462, bold=True,
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


# =====================================================================  06 核心命题
def s06_core(d):
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
        rect(s, L, y, 6, bh, fill=th.main)
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
        rect(s, gx, gy, gw, 4, fill=LAYER[k].main)
        icon(s, ic, LAYER[k].main, gx + 18, gy + 26, 30)
        text(s, gx + 58, gy + 22, gw - 66, 38, g, size=21, bold=True, color=INK, anchor='m', lh=30, wrap=False)
        text(s, gx + 18, gy + 68, gw - 24, 30, '→ ' + t, size=18, bold=True, color=LAYER[k].text, lh=26, wrap=False)
    # 价值锚点
    summary(s, L, 848, 1792, 140, '价值锚点：设计改一次，全工况列车关键性能验证自动重跑一遍',
            '全流程跑在<b>可复现沙箱</b>中——验证可回放、可追溯；于是<b>设计改一次，全工况列车关键性能验证自动重跑一遍</b>，支撑研发验证由“实物试验为主”转向“数字样机为主”。',
            title_size=28, desc_size=21)



# =====================================================================  研究内容 ①—④ 共用：左侧四层导航
def content_left(s, layer, label, body, y=None):
    yy = layer_nav(s, 64, 214, 260, layer)
    y = y or yy + 22
    text(s, 64, y, 260, 34, label, size=22, bold=True, color=INK, lh=32, wrap=False)
    text(s, 64, y + 46, 260, 1000 - y - 60, body, size=19, color=SUB, lh=30)


def card_title(s, x, y, w, title, th, size=25, uw=None):
    """居中标题 + 同色下划线（参考页四栏卡片）。"""
    text(s, x, y, w, 40, title, size=size, bold=True, color=INK, align='c', anchor='m', lh=round(size * 1.3),
         wrap=False)
    uw = uw or min(w - 40, text_width(title, size, True) + 12)
    rect(s, x + (w - uw) / 2.0, y + 48, uw, 3, fill=th.main)


# =====================================================================  07 研究内容① 基础层
def s07_layer1(d):
    s = d.new_slide()
    th = LAYER[1]
    frame(d, s, '研究内容① 测试知识库与任务工作流：为上层提供知识与规则', '研究什么：验证判据知识化 · 任务工作流建模 · 知识与流程的迭代机制', '研究内容① 基础层')
    content_left(s, 1, '定位', '为上层提供<b>知识与规则</b>——没有判据知识库，任务理解只能靠通用大模型“猜”；没有任务工作流，编排调度无章可循。')
    X, XW = 356, 1500
    cw, ch, y0 = (XW - 48) / 3.0, 520, 214
    cards = [
        ('验证判据知识化', 'list-check', '从设计规范、运用要求与历史案例中提取判据，形式化为<b>可自动比对</b>的判据条目，构建验证判据知识库'),
        ('任务工作流建模', 'hierarchy-3', '定义验证任务的分解规则、执行顺序、流转条件与异常分支，形成<b>可配置、可复用</b>的任务工作流'),
        ('知识与流程的迭代机制', 'refresh', '依据验证结果回流，持续更新判据条目与工作流定义'),
    ]
    for i, (t, ic, desc) in enumerate(cards):
        x = X + i * (cw + 24)
        rect(s, x, y0, cw, ch, fill=WHITE, line=LINE)
        num_mark(s, x + cw - 24 - 100, y0 + 16, 100, '0%d' % (i + 1), th.num, size=40)
        icon(s, ic, th.main, x + (cw - 64) / 2.0, y0 + 34, 64)
        card_title(s, x, y0 + 112, cw, t, th)
        text(s, x + 32, y0 + 184, cw - 64, 100, desc, size=20, color=BODY, lh=31)
        # 迷你示意
        vx, vy, vw, vh = x + 24, y0 + 312, cw - 48, 180
        rect(s, vx, vy, vw, vh, fill=PANEL)
        if i == 0:
            for k, t2 in enumerate(['设计规范', '运用要求', '历史案例']):
                chip(s, vx + 18, vy + 20 + k * 46, t2, w=104, h=38, size=17, color=SUB, line=LINE)
            tri(s, vx + 142, vy + vh / 2.0, 12, 16, TRI)
            chip(s, vx + 160, vy + vh / 2.0 - 19, '判据条目', w=100, h=38, size=17, color=th.text, line=th.main,
                 bold=True)
            tri(s, vx + 280, vy + vh / 2.0, 12, 16, TRI)
            chip(s, vx + 298, vy + vh / 2.0 - 24, '判据知识库', w=vw - 316, h=48, size=18, color=WHITE, fill=th.main,
                 line=None, bold=True)
        elif i == 1:
            for k, t2 in enumerate(['分解规则', '执行顺序', '流转条件', '异常分支']):
                chip(s, vx + 18 + (k % 2) * 112, vy + 36 + (k // 2) * 54, t2, w=104, h=42, size=17, color=SUB,
                     line=LINE)
            tri(s, vx + 256, vy + vh / 2.0, 12, 16, TRI)
            chip(s, vx + 274, vy + vh / 2.0 - 24, '任务工作流', w=vw - 292, h=48, size=18, color=WHITE, fill=th.main,
                 line=None, bold=True)
        else:
            chip(s, vx + 18, vy + 22, '验证结果回流', w=150, h=42, size=17, color=SUB, line=LINE)
            chip(s, vx + 18, vy + vh - 64, '更新判据条目', w=150, h=42, size=17, color=th.text, line=th.main, bold=True)
            chip(s, vx + vw - 168, vy + vh - 64, '更新工作流定义', w=150, h=42, size=17, color=th.text, line=th.main,
                 bold=True)
            line(s, vx + 93, vy + 64, vx + 93, vy + vh - 66, th.main, 2, tail='triangle')
            poly(s, [(vx + 168, vy + 43), (vx + vw - 93, vy + 43), (vx + vw - 93, vy + vh - 66)], th.main, 2,
                 tail='triangle')
            text(s, vx + 168, vy + 12, vw - 261, 26, '持续迭代', size=16, bold=True, color=th.text, align='c', lh=22,
                 wrap=False)
    # 底部：产出 + 与上层的关系
    y1 = 772
    header(s, X, y1, '产出', w=700)
    concl(s, X, y1 + 48, 700, 60, '验证判据知识库', th, size=22, check=True)
    concl(s, X, y1 + 122, 700, 60, '任务工作流定义与配置规范', th, size=22, check=True)
    R = X + 740
    RW = XW - 740
    header(s, R, y1, '与上层的关系', w=RW)
    chip(s, R, y1 + 58, '① 知识库与工作流', w=250, h=56, size=20, color=WHITE, fill=th.main, line=None, bold=True)
    line(s, R + 262, y1 + 86, R + RW - 262, y1 + 86, TRI, 2.5, tail='triangle')
    text(s, R + 262, y1 + 50, RW - 524, 30, '知识与规则来源', size=17, bold=True, color=SUB, align='c', lh=24,
         wrap=False)
    chip(s, R + RW - 250, y1 + 58, '② 任务理解与编排调度', w=250, h=56, size=20, color=WHITE, fill=BLUE.main, line=None,
         bold=True)
    text(s, R, y1 + 132, RW, 60, '知识库与工作流是“任务理解与编排调度”的<b>知识与规则来源</b>。', size=20, color=BODY, lh=30)


# =====================================================================  08 研究内容② 决策层
def s08_layer2(d):
    s = d.new_slide()
    th = LAYER[2]
    frame(d, s, '研究内容② 任务理解与智能编排调度：引擎在此落地', '研究什么：验证目标解析与任务分解 · 测试环境自动装配 · 模型保真度与算力自适应调度 · 执行监控与失败重规划',
          '研究内容② 决策层')
    content_left(s, 2, '定位', '本课题的<b>核心决策层</b>——“多智能体协同运行与调度引擎”在此落地。')
    X, XW = 356, 1500
    cw, ch, y0 = (XW - 60) / 4.0, 630, 214
    cards = [
        ('验证目标解析与任务分解', 'target-arrow', '把“验证该机车某项性能是否满足要求”拆解为可执行验证任务，并匹配知识库中的判据与工作流'),
        ('测试环境自动装配', 'plug-connected', '按验证目标自动完成智能体组合、接口连接与算力资源配置'),
        ('模型保真度与算力自适应调度', 'adjustments', '按工况特征匹配模型精度与算力配置'),
        ('执行监控与失败重规划', 'activity', '运行过程监控、异常中断后重规划，支撑指标实时推演'),
    ]
    for i, (t, ic, desc) in enumerate(cards):
        x = X + i * (cw + 20)
        rect(s, x, y0, cw, ch, fill=WHITE, line=LINE)
        num_mark(s, x + cw - 20 - 90, y0 + 14, 90, '0%d' % (i + 1), th.num, size=36)
        icon(s, ic, th.main, x + (cw - 60) / 2.0, y0 + 32, 60)
        card_title(s, x, y0 + 104, cw, t, th, size=23 if len(t) > 11 else 24)
        text(s, x + 26, y0 + 176, cw - 52, 130, desc, size=20, color=BODY, lh=31)
        vx, vy, vw = x + 20, y0 + 322, cw - 40
        vh = ch - 322 - 20
        rect(s, vx, vy, vw, vh, fill=PANEL)
        if i == 0:
            text(s, vx, vy + 18, vw, 26, '拆解为可执行验证任务', size=16, bold=True, color=MUTED, align='c', lh=24,
                 wrap=False)
            for k, t2 in enumerate(['验证对象', '关键性能', '判据条目']):
                chip(s, vx + 24, vy + 58 + k * 60, t2, w=vw - 48, h=46, size=19, color=th.text, line=th.main,
                     bold=True)
        elif i == 1:
            text(s, vx, vy + 18, vw, 26, '按验证目标自动完成', size=16, bold=True, color=MUTED, align='c', lh=24,
                 wrap=False)
            for k, t2 in enumerate(['智能体组合', '接口连接', '算力资源配置']):
                cy = vy + 58 + k * 70
                chip(s, vx + 24, cy, t2, w=vw - 48, h=46, size=19, color=th.text, line=th.main, bold=True)
                if k < 2:
                    text(s, vx, cy + 44, vw, 28, '＋', size=18, bold=True, color=MUTED, align='c', lh=24, wrap=False)
        elif i == 2:
            text(s, vx, vy + 18, vw, 26, '保真度分配', size=16, bold=True, color=MUTED, align='c', lh=24, wrap=False)
            rows = [('常规工况', '降阶模型提速', th.text, th.bg), ('极端工况', '强制全保真保精度', RED, RED_BG)]
            for k, (a, b, fc, bg) in enumerate(rows):
                cy = vy + 60 + k * 110
                text(s, vx + 20, cy, vw - 40, 30, a, size=18, bold=True, color=INK, lh=26, wrap=False)
                chip(s, vx + 20, cy + 34, b, w=vw - 40, h=46, size=19, color=fc, fill=bg, line=fc, bold=True)
        else:
            text(s, vx, vy + 18, vw, 26, '监控 → 重规划闭环', size=16, bold=True, color=MUTED, align='c', lh=24,
                 wrap=False)
            steps = ['运行过程监控', '异常中断', '重规划']
            for k, t2 in enumerate(steps):
                cy = vy + 58 + k * 70
                chip(s, vx + 24, cy, t2, w=vw - 84, h=46, size=19, color=th.text if k != 1 else RED,
                     line=th.main if k != 1 else RED, bold=True)
                if k < 2:
                    line(s, vx + 24 + (vw - 84) / 2.0, cy + 46, vx + 24 + (vw - 84) / 2.0, cy + 70, TRI, 2,
                         tail='triangle')
            poly(s, [(vx + vw - 60, vy + 58 + 140 + 23), (vx + vw - 30, vy + 58 + 140 + 23),
                     (vx + vw - 30, vy + 58 + 23), (vx + vw - 60, vy + 58 + 23)], th.main, 2, tail='triangle')
    # 底部：产出
    y1 = 874
    concl(s, X, y1, 720, 76, '产出：多智能体协同运行与调度引擎（软件1套）', th, size=24, check=True)
    fx = X + 760
    chips_row(s, fx, y1 + 16, ['① 判据与工作流', '② 调度引擎', '③ 执行层'], gap=54, arrow=True, arrow_color=TRI,
              size=19, h=44, padx=16, color=INK, line=LINE)
    text(s, fx, y1 + 74, XW - 760, 40, '向下承接知识与规则输入，向执行层下发执行方案（智能体组合、工况、算力）', size=17, color=SUB, lh=24)


# =====================================================================  09 研究内容③ 执行层
def s09_layer3(d):
    s = d.new_slide()
    th = LAYER[3]
    frame(d, s, '研究内容③ 智能体执行：一条完整的验证执行链', '仿真智能体 → 测试用例生成 → 测试识别 → 测试报告 → 数据回归，结果回流、支持回归重跑', '研究内容③ 执行层')
    content_left(s, 3, '对接项目7总体目标',
                 [{'t': '<b>模型自动标定</b><br>→ 01 仿真智能体', 'sa': 14}, {'t': '<b>性能自动评估</b><br>→ 03 测试识别'}])
    X, XW = 356, 1500
    gap = 30
    cw, ch, y0 = (XW - 4 * gap) / 5.0, 530, 214
    keys = ['模型自动标定', '按需自生成', '异常定位', '可追溯', '回归重跑']
    steps = [
        ('仿真智能体', '研制与接入', 'robot',
         '研制控制系统、车辆动力学与线路等<b>仿真智能体与智能孪生模型</b>；标准化接入、时序同步、<b>模型自动标定</b>'),
        ('测试用例生成', '用例与线路数据按需自生成', 'file-text', '按覆盖完备性生成用例与线路模型数据（平纵断面、曲线半径、超高、坡度、不平顺谱）'),
        ('测试识别', '结果评判 · 异常定位', 'zoom-check', '逐工况对标判据，<b>识别异常工况并定位原因</b>，输出结果可信性判断'),
        ('测试报告', '报告生成', 'report', '验证结论自动汇总，生成<b>可追溯</b>的验证报告'),
        ('数据回归', '回流迭代', 'refresh', '验证结果回流，迭代模型标定、判据与工况集；支持设计变更后的<b>回归重跑</b>'),
    ]
    cxs = []
    for i, (t, sub, ic, desc) in enumerate(steps):
        x = X + i * (cw + gap)
        cxs.append(x + cw / 2.0)
        rect(s, x, y0, cw, ch, fill=WHITE, line=LINE)
        rect(s, x, y0, cw, 6, fill=th.main)
        num_mark(s, x + cw - 18 - 80, y0 + 18, 80, '0%d' % (i + 1), th.num, size=36)
        icon_badge(s, ic, th, x + 22, y0 + 30, 64)
        text(s, x + 22, y0 + 112, cw - 40, 38, t, size=25, bold=True, color=INK, anchor='m', lh=34, wrap=False)
        text(s, x + 22, y0 + 154, cw - 40, 58, sub, size=17, bold=True, color=th.text, lh=25)
        line(s, x + 22, y0 + 216, x + cw - 22, y0 + 216, LINE, 1.5)
        text(s, x + 22, y0 + 234, cw - 40, ch - 330, desc, size=19, color=BODY, lh=30)
        concl(s, x + 18, y0 + ch - 18 - 52, cw - 36, 52, keys[i], th, size=19, check=True, pad=14)
        if i < 4:
            tri(s, x + cw + gap / 2.0, y0 + ch / 2.0, 20, 26, TRI)
    # 数据回归回流线
    yb = y0 + ch
    poly(s, [(cxs[4], yb), (cxs[4], yb + 40), (cxs[0], yb + 40), (cxs[0], yb + 2)], GREEN.main, 2.5, tail='triangle')
    text(s, cxs[0] + 40, yb + 50, cxs[4] - cxs[0] - 80, 30, '数据回归：验证结果回流，迭代模型标定、判据与工况集 · 设计变更后回归重跑', size=19, bold=True,
         color=GREEN.text, align='c', lh=26, wrap=False)
    # 产出
    y1 = 846
    header(s, X, y1, '产出', w=XW)
    bw = (XW - 40) / 3.0
    for i, t in enumerate(['关键系统仿真智能体／智能孪生模型', '用例与线路数据集', '验证报告与追溯记录']):
        concl(s, X + i * (bw + 20), y1 + 48, bw, 64, t, th, size=21, check=True)


# =====================================================================  10 研究内容④ 集成层
def s10_layer4(d):
    s = d.new_slide()
    th = LAYER[4]
    frame(d, s, '研究内容④ 架构融合与协同运行：融合成一个能跑起来的整体', '研究什么：总体架构与接口规范 · 多智能体协同运行 · 可复现沙箱 · 列车级数字样机', '研究内容④ 集成层')
    content_left(s, 4, '定位', '把前三层<b>融合成一个能跑起来的整体</b>，形成列车级数字样机。')
    X, XW = 356, 1500
    pw, ph = (XW - 24) / 2.0, 290
    pos = [(X, 214), (X + pw + 24, 214), (X, 214 + ph + 22), (X + pw + 24, 214 + ph + 22)]
    heads = [('总体架构与接口规范', 'sitemap'), ('多智能体协同运行', 'topology-star-3'), ('可复现沙箱', 'box'),
             ('列车级数字样机', 'train')]
    tops = []
    for i, ((t, ic), (x, y)) in enumerate(zip(heads, pos)):
        tops.append(panel(s, x, y, pw, ph, th, t, '0%d' % (i + 1), ic, title_size=28, badge=64, pad=32))
    # 01
    x, y = pos[0]
    yy = tops[0] + 18
    text(s, x + 32, yy, pw - 64, 34, '定义四层之间的<b>接口、数据流与控制流</b>，形成<b>可扩展</b>的架构规范', size=20, color=BODY,
         lh=30, wrap=False)
    ex = chips_row(s, x + 32, yy + 58, ['接口', '数据流', '控制流'], gap=12, size=19, h=44, padx=20, color=th.text,
                   line=th.main, bold=True)
    tri(s, ex + 26, yy + 80, 14, 18, TRI)
    chip(s, ex + 46, yy + 58, '可扩展的架构规范', h=44, size=19, padx=20, color=WHITE, fill=th.main, line=None, bold=True)
    # 02
    x, y = pos[1]
    yy = tops[1] + 18
    text(s, x + 32, yy, pw - 64, 34, '异构智能体的<b>实时协同求解、时序同步与状态同步</b>', size=20, color=BODY, lh=30, wrap=False)
    chips_row(s, x + 32, yy + 58, ['实时协同求解', '时序同步', '状态同步'], gap=12, arrow=False, size=19, h=44, padx=20,
              color=th.text, line=th.main, bold=True)
    # 03 沙箱三件事
    x, y = pos[2]
    yy = tops[2] + 16
    bw = (pw - 64 - 24) / 3.0
    for k, (a, b, ic) in enumerate([('隔离', '单路模型发散不扩散', 'shield-check'), ('记录', '参数、模型版本、<br>随机种子', 'file-text'),
                                    ('回放', '任意一次验证可复现', 'player-play')]):
        bx = x + 32 + k * (bw + 12)
        rect(s, bx, yy, bw, 116, fill=th.bg)
        icon(s, ic, th.main, bx + 16, yy + 16, 28)
        text(s, bx + 52, yy + 14, bw - 60, 32, a, size=21, bold=True, color=th.text, lh=30, wrap=False)
        text(s, bx + 16, yy + 54, bw - 28, 56, b, size=18, color=BODY, lh=26)
    # 04 数字样机
    x, y = pos[3]
    yy = tops[3] + 18
    text(s, x + 32, yy, pw - 64, 34, '集成关键系统仿真智能体，支持<b>列车关键性能验证与智能运维</b>', size=20, color=BODY, lh=30,
         wrap=False)
    ruler(s, x + 32, y + ph - 34, pw - 64, step=36, tick=8, lw=2)
    loco_train(s, x + 32 + (pw - 64 - 2 * 300 - 8) / 2.0, y + ph - 34, 300, n=2, gap=8, faded=True)
    # 产出
    y1 = 846
    header(s, X, y1, '产出', w=XW)
    bw = (XW - 20) / 2.0
    concl(s, X, y1 + 48, bw, 64, '列车级数字样机', th, size=21, check=True)
    concl(s, X + bw + 20, y1 + 48, bw, 64, '列车级数字样机构建方法体系与集成技术规范', th, size=21, check=True)


# =====================================================================  11 研究方法
def s11_methods(d):
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
        rect(s, X, y, 6, rh, fill=th.main)
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


# =====================================================================  12 技术路线（架构图版式）
def s12_route(d):
    s = d.new_slide()
    frame(d, s, '技术路线：四层递进 ＋ 沙箱底座，闭环迭代', '四层递进：① 知识 → ② 调度 → ③ 执行 → ④ 集成；沙箱底座：隔离 · 记录 · 回放', '技术路线')
    L, LW = 64, 1112
    header(s, L, 212, '总体技术路线', note='四层递进 ＋ 可复现沙箱底座', w=LW)
    zy, zh = 290, 556
    dashed_zone(s, L, zy, LW, zh, '可复现沙箱底座（隔离 · 记录 · 回放）', label_fill=ARCH, fill='FBFCFE', label_x=L + 28)
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
         wrap=False, shp=rect(s, cxs[2] - tw / 2.0, ly - 14, tw, 28, fill='FBFCFE'))
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
        rect(s, R, y, 6, lh_, fill=col)
        tag(s, R + 26, y + 20, t, fill=col, size=19, h=34, padx=14)
        text(s, R + 150, y + 20, RW - 176, 34, st, size=19, bold=True, color=RED if col == RED else SUB, align='r',
             anchor='m', lh=26, wrap=False)
        text(s, R + 26, y + 66, RW - 52, 60, body, size=20, color=BODY, lh=30)
        y += lh_ + 13
    # 路线闭环
    y1 = 882
    rect(s, 64, y1, 5, 104, fill=RED)
    text(s, 92, y1, 200, 104, '路线闭环', size=28, bold=True, color=INK, anchor='m', lh=36, wrap=False)
    loop = ['判据知识化', '任务生成与调度', '协同执行', '结果识别', '数据回归', '更新知识库']
    ex = chips_row(s, 300, y1 + 30, loop, gap=46, arrow=True, arrow_color=MUTED, size=20, h=46, padx=18, color=INK,
                   line=LINE, bold=True)
    icon(s, 'refresh', RED, ex + 20, y1 + 33, 40)
    text(s, ex + 70, y1 + 30, 1856 - ex - 70, 46, '闭环迭代', size=22, bold=True, color=RED, anchor='m', lh=30,
         wrap=False)


# =====================================================================  13 典型验证场景
def s13_scenario(d):
    s = d.new_slide()
    frame(d, s, '典型验证场景：控速＋能耗验证，全流程交给智能体', '输入一句话需求 → ① 任务理解 → ② 任务编排 → ③ 智能体执行 → 验证报告 ＋ 全要素留痕', '典型验证场景')
    # 输入
    y0 = 214
    shp = rect(s, 64, y0, 250, 84, fill=INK)
    text(s, 64, y0, 250, 84, '输入 · 一句话需求', size=22, bold=True, color=WHITE, align='c', anchor='m', lh=30, shp=shp,
         wrap=False)
    rect(s, 314, y0, 1542, 84, fill=PANEL)
    icon(s, 'quote', MUTED, 340, y0 + 24, 36)
    text(s, 394, y0, 1440, 84, '验证该机车在指定线路上按给定运行图运行时的<r>速度跟踪</r>与<r>牵引能耗</r>是否满足要求。', size=26, bold=True,
         color=INK, anchor='m', lh=36, wrap=False)
    # 三步面板
    py, ph = 326, 556
    ws = [450, 806, 456]
    gap = (1792 - sum(ws)) / 2.0
    xs = [64, 64 + ws[0] + gap, 64 + ws[0] + ws[1] + 2 * gap]
    specs = [('任务理解', '变成可执行的验证任务', BLUE, 'target-arrow'), ('任务编排', '生成数据、工况与环境', PURPLE, 'route'),
             ('智能体执行', '协同求解与结果自评判', GREEN, 'robot')]
    tops = []
    for i, ((t, sub, th, ic), x, w) in enumerate(zip(specs, xs, ws)):
        rect(s, x, py, w, ph, fill=WHITE, line=LINE)
        rect(s, x, py, w, 6, fill=th.main)
        icon_badge(s, ic, th, x + 28, py + 28, 64)
        text(s, x + 110, py + 26, w - 200, 38, '%s' % t, size=27, bold=True, color=INK, lh=36, wrap=False)
        text(s, x + 110, py + 64, w - 200, 30, '→ ' + sub, size=18, bold=True, color=th.text, lh=26, wrap=False)
        num_mark(s, x + w - 28 - 90, py + 18, 90, '0%d' % (i + 1), th.num, size=40)
        line(s, x + 28, py + 114, x + w - 28, py + 114, LINE, 1.5)
        tops.append(py + 132)
        if i < 2:
            tri(s, x + w + gap / 2.0, py + ph / 2.0, 20, 28, TRI)
    # ① 任务理解
    x, w, th = xs[0], ws[0], BLUE
    y = tops[0]
    for lab_, val in [('验证对象', '牵引控制系统<br>（含级位控制、防空转／防滑行）'), ('关键性能', '<b>速度跟踪偏差、牵引能耗</b>'),
                      ('验证判据', '从设计规范与运用要求中提取，形式化为<b>可自动比对</b>的判据条目')]:
        tag(s, x + 28, y, lab_, fill=th.bg, color=th.text, size=18, h=32, padx=12)
        n = wrap_lines(strip_tags(val), w - 56, 20)
        text(s, x + 28, y + 40, w - 56, n * 31, val, size=20, color=BODY, lh=31)
        y += 40 + n * 31 + 30
    # ② 任务编排：2×2 子块
    x, w, th = xs[1], ws[1], PURPLE
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
    x, w, th = xs[2], ws[2], GREEN
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


# =====================================================================  14 引擎创新内核
def s14_innovation(d):
    s = d.new_slide()
    frame(d, s, '引擎创新内核：先进性不在单项技术，而在调度的维度与粒度', '三个技术判断：调度对象是“智能体” · 调度决策多一个维度：模型保真度 · 调度基础是可复现沙箱', '引擎创新内核')
    specs = [
        (BLUE, 'robot', '调度对象是“智能体”，不是“模型”', '调度的是算例', '调度的是<b>可自主标定、自主判断的智能体</b>',
         '调度粒度从“跑模型”提升到“派任务”，人退出执行回路', None),
        (PURPLE, 'adjustments', '调度决策多一个维度：模型保真度', '在“快”与“准”之间取舍',
         '按工况特征动态分配——<b>常规工况用降阶模型提速，关键与极端工况强制全保真保精度</b>', '批量提速与关键工况精度同时成立', ['降阶 · 快', '全保真 · 准']),
        (GREEN, 'history', '调度基础是可复现沙箱', '仿真结果难以复现', '对<b>参数、模型版本、随机种子</b>全程留痕', '验证可回放、可追溯、可回归',
         ['参数', '模型版本', '随机种子']),
    ]
    pw, ph, y0 = (1792 - 64) / 3.0, 676, 214
    for i, (th, ic, title, conv, ours, cap, mini) in enumerate(specs):
        x = 64 + i * (pw + 32)
        rect(s, x, y0, pw, ph, fill=WHITE, line=LINE)
        rect(s, x, y0, pw, 6, fill=th.main)
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


# =====================================================================  15 技术对标与技术定位
def s15_benchmark(d):
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
    trs = [['<b>%s</b>' % a, '<c=%s>✓</c>  %s' % (GREEN.main, b), '<bc=%s>×</bc>  %s' % (RED, c)] for a, b, c in rows]
    table(s, L, 256, cols, trs, header=['现有成熟技术', '已解决', '未解决（本课题的定位）'], size=20, lh=30, pad_y=33,
          pad_x=18, col_styles={0: dict(size=20), 1: dict(color=GREEN.text), 2: dict(color=BODY)})
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
    rect(s, R, by, 5, 136, fill=RED)
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


# =====================================================================  16 引擎的未来价值
def s16_future(d):
    s = d.new_slide()
    frame(d, s, '未来价值：从“一个课题的交付物”到“可生长的验证能力底座”', '能力扩展 · 场景延伸 · 范式转变 · 资产沉淀', '引擎的未来价值')
    items = [
        ('能力扩展', 'puzzle', BLUE, '新增一个关键系统＝新一轮建模与集成', '按标准化接口接入智能体即可，<b>引擎不需重构</b>'),
        ('场景延伸', 'arrows-split-2', PURPLE, '数字样机只服务研发验证', '延伸到<b>智能运维</b>（性能复现、状态评估）；方法体系可迁移至其他装备领域'),
        ('范式转变', 'repeat', GREEN, '验证是项目式、一次性的', '<b>常态化、可回归</b>——设计改一次，全工况验证自动重跑'),
        ('资产沉淀', 'database', ORANGE, '经验在专家个人手里', '判据、工况、结论全程留痕，<b>个人经验沉淀为组织资产</b>'),
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


# =====================================================================  17 阶段目标与进度安排
def s17_schedule(d):
    s = d.new_slide()
    frame(d, s, '阶段目标与进度：四阶段推进，前一阶段交付即后一阶段输入', '研究周期 2027.01—2028.12（24 个月） · 四个里程碑 M1—M4 与阶段交付物对齐', '阶段目标与进度安排')
    X0, X1 = 64, 1844
    mw = (X1 - X0) / 24.0
    stages = [
        ('阶段一', '架构与规范', 0, 6, STAGE[0], '2027.01—2027.06', '需求分析与总体架构设计', '总体架构、验证判据知识化方法与标准化接口规范',
         'M1', '2027.06', '架构与接口规范'),
        ('阶段二', '智能体与引擎', 6, 15, STAGE[1], '2027.07—2028.03', '智能体与引擎研制', '关键系统仿真智能体／智能孪生模型、调度引擎原理样机',
         'M2', '2028.03', '智能体与引擎'),
        ('阶段三', '样机与验证', 15, 21, STAGE[2], '2028.04—2028.09', '样机集成与验证', '列车级数字样机、典型工况协同验证结果', 'M3',
         '2028.09', '数字样机与验证'),
        ('阶段四', '示范与验收', 21, 24, STAGE[3], '2028.10—2028.12', '示范应用与验收', '示范应用与效能评估、方法体系与技术规范、标准草案',
         'M4', '2028.12', '示范应用与验收'),
    ]
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
        fg = INK if col == STAGE[1] else WHITE
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
        rect(s, x, cy0, 6, ch, fill=col)
        tag(s, x + 26, cy0 + 22, name, fill=col, color=INK if col == STAGE[1] else WHITE, size=18, h=32, padx=14)
        text(s, x + 140, cy0 + 22, cw - 160, 32, per, size=18, bold=True, color=MUTED, font=MONO, align='r',
             anchor='m', lh=24, wrap=False)
        text(s, x + 26, cy0 + 72, cw - 52, 40, goal, size=25, bold=True, color=INK, lh=34, wrap=False)
        line(s, x + 26, cy0 + 126, x + cw - 26, cy0 + 126, LINE, 1.5)
        text(s, x + 26, cy0 + 140, cw - 52, 28, '阶段交付', size=17, bold=True, color=MUTED, lh=24, wrap=False)
        parts = [p for p in deli.split('、')]
        bullets(s, x + 26, cy0 + 174, cw - 52, 150, parts, col if col != STAGE[1] else 'D9A300', size=19, lh=29, sa=6)
        rect(s, x + 20, cy0 + ch - 20 - 50, cw - 40, 50, fill=RED_BG)
        text(s, x + 36, cy0 + ch - 20 - 50, cw - 72, 50, '<r>%s</r>（%s）%s' % (mk, mdate, mtext), size=19, color=INK,
             anchor='m', lh=26, bold=True, wrap=False)
        if i < 3:
            tri(s, x + cw + 18, cy0 + ch / 2.0, 18, 24, TRI)
    summary(s, 64, 900, 1792, 92, '阶段间衔接：前一阶段交付物即后一阶段输入',
            '判据与接口规范 → 任务生成与智能体接入；智能体与引擎 → 样机集成；样机 → 性能验证', title_size=25, desc_size=20, gap=2)


# =====================================================================  18 任务分解与考核方式
def s18_tasks(d):
    s = d.new_slide()
    frame(d, s, '任务分解与考核：四个子任务，交付物逐级成为下一任务输入', '子任务1 知识 → 子任务2 调度 → 子任务3 执行 → 子任务4 集成；数据回归更新判据与工作流', '预期成果与效益')
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
        rect(s, x, y0, cw, 6, fill=th.main)
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
    poly(s, [(cxs[3], yb), (cxs[3], yb + 46), (cxs[0], yb + 46), (cxs[0], yb + 2)], GREEN.main, 2.5, tail='triangle')
    text(s, cxs[0], yb + 56, cxs[3] - cxs[0], 30, '数据回归：更新判据与工作流', size=20, bold=True, color=GREEN.text, align='c',
         lh=28, wrap=False)
    summary(s, 64, 872, 1792, 118, '每个子任务的交付物即下一子任务的输入', '形成“<r>知识—调度—执行—集成</r>”的闭环。', title_size=28,
            desc_size=22)


# =====================================================================  19 预期目标与成果对应
def s19_goals(d):
    s = d.new_slide()
    frame(d, s, '预期目标与成果对应：四项目标，每项都有对应成果物', '调度引擎 · 数字样机与性能验证 · 方法体系、智能运维与扩展能力 · 示范应用', '预期成果与效益')
    goals = [
        ('01', '调度引擎', BLUE, '形成<b>多智能体协同运行与调度引擎1套</b>，支持异构仿真智能体的实时协同运行与算力资源动态调度',
         [('多智能体协同运行与调度引擎', '软件1套')]),
        ('02', '数字样机与性能验证', PURPLE,
         '建成包含车辆动力学与线路等关键系统仿真智能体的<b>列车级数字样机</b>，实现对列车速度跟踪控制品质、牵引能耗与再生能量利用、运行平稳性等关键性能的<b>多工况协同验证</b>',
         [('关键系统仿真智能体／智能孪生模型', '1套'), ('列车级数字样机', '1套'), ('示范应用与验证报告', '1份')]),
        ('03', '方法体系、智能运维\n与扩展能力', GREEN,
         '形成列车级数字样机构建方法体系与标准化集成技术规范，<b>支持列车关键性能验证与智能运维</b>，具备向制动、辅助系统等关键系统扩展的能力',
         [('列车级数字样机构建方法体系与技术规范', '1套')]),
        ('04', '示范应用', ORANGE, '在典型机车产品上完成示范应用，关键性能验证的<b>效率与工况覆盖度显著提升</b>，形成可推广的智能验证能力',
         [('示范应用与验证报告', '1份'), ('专利 · 论文 · 标准草案', '4项 · 2篇 · 2项')]),
    ]
    text(s, 64, 212, 400, 32, '预期目标', size=18, bold=True, color=MUTED, lh=26, wrap=False)
    text(s, 1250, 212, 400, 32, '对应成果物', size=18, bold=True, color=MUTED, lh=26, wrap=False)
    y = 250
    rh = 176
    for no, name, th, txt, outs in goals:
        rect(s, 64, y, 1792, rh, fill=WHITE, line=LINE)
        rect(s, 64, y, 6, rh, fill=th.main)
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
    frame(d, s, '量化指标与效益评估：成果交付 ＋ 示范验证 ＋ 对比评估', '说明：量化以成果数量与能力覆盖为主，具体性能限值在课题任务书中进一步细化', '预期成果与效益')
    L, LW = 64, 872
    header(s, L, 212, '量化指标', note='成果数量', w=LW)
    kp = [('软件', '1', '套'), ('模型', '1', '套'), ('系统', '1', '套'), ('规范', '1', '套'), ('报告', '1', '份'),
          ('专利', '4', '项'), ('论文', '2', '篇'), ('标准草案', '2', '项')]
    tw_, th_ = (LW - 3 * 12) / 4.0, 124
    for i, (lab_, v, u) in enumerate(kp):
        x = L + (i % 4) * (tw_ + 12)
        y = 256 + (i // 4) * (th_ + 12)
        rect(s, x, y, tw_, th_, fill=GRAYBG)
        text(s, x + 22, y + 14, 150, 64, '%s<s=20><n><c=%s> %s</c></n></s>' % (v, SUB, u), size=50, bold=True,
             color=RED if i >= 5 else INK, font=MONO, lh=64, wrap=False)
        text(s, x + 22, y + 82, tw_ - 30, 28, lab_, size=19, bold=True, color=SUB, lh=26, wrap=False)
    rows = [
        ('能力覆盖', '支持<b>异构</b>仿真智能体接入；支持速度跟踪控制品质、牵引能耗与再生能量利用、运行平稳性等关键性能的<b>多工况</b>协同验证'),
        ('扩展能力', '形成<b>可扩展</b>的智能体接入机制，支持<b>智能运维</b>应用，具备向制动、辅助系统等关键系统扩展的能力'),
        ('效率提升', '关键性能验证的<b>效率与工况覆盖度显著提升</b>'),
    ]
    y = 256 + 2 * th_ + 12 + 24
    ns = [wrap_lines(txt, LW - 130, 20) for _, txt in rows]
    base = [n * 31 + 36 for n in ns]
    pad = (256 + 4 * 144 + 3 * 12 - y - sum(base)) / float(len(rows))
    for (lab_, txt), n, b in zip(rows, ns, base):
        h = b + pad
        line(s, L, y, L + LW, y, LINE, 1)
        top = y + (h - n * 31) / 2.0
        tag(s, L, top - 1, lab_, fill=BLUE.bg, color=BLUE.text, size=18, h=32, padx=12)
        text(s, L + 130, top, LW - 130, n * 31 + 4, txt, size=20, color=BODY, lh=31)
        y += h
    line(s, L, y, L + LW, y, LINE, 1)
    # 右：效益评估
    R, RW = 976, 880
    header(s, R, 212, '效益评估', note='效益 · 评估方式', w=RW)
    bens = [
        ('技术效益', 'bulb', BLUE, '形成列车级数字样机构建方法体系与技术规范，填补多智能体协同调度的领域化应用空白', '方法体系与规范交付、示范应用验证'),
        ('经济效益', 'coins', GREEN, '减少实物试验与样车试制投入，缩短研制迭代周期', '试验次数与周期对比评估'),
        ('管理效益', 'chart-bar', PURPLE, '提升关键性能验证的效率与工况覆盖度，支撑研发流程标准化', '验证效率与覆盖度对比'),
        ('推广效益', 'world', ORANGE, '方法体系支持智能运维并具备向制动、辅助系统扩展能力，可在**内推广', '扩展性验证与推广方案'),
    ]
    y = 256
    bh = 144
    for t, ic, th, txt, how in bens:
        rect(s, R, y, RW, bh, fill=WHITE, line=LINE)
        rect(s, R, y, 6, bh, fill=th.main)
        icon(s, ic, th.main, R + 26, y + 22, 32)
        text(s, R + 70, y + 18, 200, 40, t, size=22, bold=True, color=INK, anchor='m', lh=30, wrap=False)
        text(s, R + 26, y + 62, RW - 52, 40, txt, size=19, color=BODY, lh=28, wrap=False)
        text(s, R + 26, y + 100, RW - 52, 32, '<k>评估方式</k>　' + how, size=18, color=th.text, lh=26, wrap=False)
        y += bh + 12
    summary(s, 64, 908, 1792, 84, '效益评估以“<r>成果交付＋示范验证＋对比评估</r>”三层支撑，每项成果均对应一项研究目标。', title_size=27)


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
    ], BLUE.main, size=20, lh=31, sa=20)


# =====================================================================  22 项目管理计划
def s22_management(d):
    s = d.new_slide()
    frame(d, s, '项目管理计划：按四层分组承担，按里程碑绑定考核', '组织架构与任务分配 · 管理流程 · 进度控制与执行力保障', '项目管理计划')
    # 组织架构
    header(s, 64, 212, '组织架构与任务分配', w=900)
    bx, bw = 960 - 170, 340
    rect(s, bx, 214, bw, 58, fill=INK)
    icon(s, 'user-star', WHITE, bx + 70, 214 + 14, 30)
    text(s, bx + 110, 214, bw - 130, 58, '项目负责人', size=23, bold=True, color=WHITE, anchor='m', lh=30, wrap=False)
    groups = [
        (1, '判据知识化、工作流建模、知识与流程迭代', '判据知识库、工作流定义', None),
        (2, '任务解析、环境装配、保真度与算力调度', '协同运行与调度引擎', None),
        (3, '智能体研制、用例生成、识别与报告、数据回归', '仿真智能体／孪生模型、验证报告', '仿真智能体／用例／识别／报告／回归'),
        (4, '架构设计、协同运行、沙箱、样机集成', '列车级数字样机、方法体系与规范', None),
    ]
    gw = (1792 - 3 * 24) / 4.0
    gy = 322
    line(s, 960, 272, 960, 296, MUTED, 1.5)
    line(s, 64 + gw / 2.0, 296, 64 + 3 * (gw + 24) + gw / 2.0, 296, MUTED, 1.5)
    for i, (k, task, deli, extra) in enumerate(groups):
        th = LAYER[k]
        x = 64 + i * (gw + 24)
        line(s, x + gw / 2.0, 296, x + gw / 2.0, gy - 2, MUTED, 1.5, tail='triangle')
        rect(s, x, gy, gw, 284, fill=WHITE, line=LINE)
        rect(s, x, gy, gw, 6, fill=th.main)
        tag(s, x + 22, gy + 22, '子任务%d组' % k, fill=th.main, size=17, h=30, padx=12)
        text(s, x + 22 + text_width('子任务%d组' % k, 17, True) + 38, gy + 22, gw - 160, 30, '对应' + LAYER_NO[k],
             size=17, bold=True, color=th.text, anchor='m', lh=24, wrap=False)
        text(s, x + 22, gy + 62, gw - 44, 34, LAYER_NAME[k], size=22, bold=True, color=INK, lh=30, wrap=False)
        text(s, x + 22, gy + 96, gw - 44, 26, extra or '', size=16, color=SUB, lh=22, wrap=False)
        text(s, x + 22, gy + 134, 120, 26, '主要任务', size=16, bold=True, color=MUTED, lh=22, wrap=False)
        text(s, x + 22, gy + 160, gw - 44, 56, task, size=17, color=BODY, lh=26)
        concl(s, x + 16, gy + 284 - 16 - 48, gw - 32, 48, '交付：' + deli, th, size=17, pad=12)
    # 管理流程
    y1 = 634
    L, LW = 64, 876
    header(s, L, y1, '管理流程', note='双周 → 月度 → 季度 → 年度', w=LW)
    cad = [('双周', '双周例会', '子任务进展同步与问题协调'), ('月度', '月度节点检查', '对照里程碑核查交付物'),
           ('季度', '季度评审', '技术方案评审与阶段成果确认'), ('年度', '年度总结', '目标达成评估与下年度计划调整')]
    cw = (LW - 3 * 12) / 4.0
    for i, (f, n, t) in enumerate(cad):
        x = L + i * (cw + 12)
        hh = 206 + i * 36
        top = 990 - hh
        rect(s, x, top, cw, hh, fill=BLUE.bg if i == 3 else PANEL, line=LINE)
        rect(s, x, top, cw, 5, fill=BLUE.main)
        oval(s, x + 20, top + 24, 60, 60, fill=BLUE.soft)
        text(s, x + 20, top + 24, 60, 60, f, size=18, bold=True, color=BLUE.text, align='c', anchor='m', lh=24,
             wrap=False)
        text(s, x + 20, top + 98, cw - 40, 34, n, size=22, bold=True, color=INK, lh=30, wrap=False)
        text(s, x + 20, top + 138, cw - 36, 80, t, size=19, color=BODY, lh=28)
    # 进度控制与执行力保障
    R, RW = 980, 876
    header(s, R, y1, '进度控制与执行力保障', w=RW)
    ctl = [('里程碑绑定考核', 'flag', None), ('质量控制', 'shield-check', '内部评审 ＋ 第三方验证，交付物可核查'),
           ('风险管控', 'alert-triangle', '技术风险（模型精度、实时性、判据知识化完备性）、进度风险、接口风险，均设置应对预案'),
           ('资源保障', 'server', '算力资源、试验与运用数据、样车与试验台资源')]
    y = y1 + 48
    hs = [72, 58, 92, 58]
    for (t, ic, body), h in zip(ctl, hs):
        line(s, R, y, R + RW, y, LINE, 1)
        icon(s, ic, RED if t == '风险管控' else INK, R, y + 16, 28)
        text(s, R + 40, y + 12, 200, 36, t, size=20, bold=True, color=INK, lh=28, wrap=False)
        if body:
            text(s, R + 230, y + 14, RW - 230, h - 16, body, size=18, color=BODY, lh=27)
        else:
            chips_row(s, R + 230, y + 14, ['M1 架构与接口规范', 'M2 智能体与引擎', 'M3 数字样机与验证', 'M4 示范应用与验收'], gap=8,
                      size=15, h=32, padx=8, color=RED, fill=RED_BG, line=None, bold=True)
        y += h + 6
    line(s, R, y - 6, R + RW, y - 6, LINE, 1)


# =====================================================================  23 团队研究能力
def s23_team(d):
    s = d.new_slide()
    frame(d, s, '团队研究能力：为何由我单位承担', '牵头单位：********** · 研发平台 · 仿真能力 · 试验能力 · 数据积累 · 在研项目 · 人才队伍', '团队研究能力')
    items = [
        ('研发平台', 'building-factory-2', '【待填：机车总体设计／控制系统开发等平台】'),
        ('仿真能力', 'device-desktop-analytics', '【待填：已有的车辆动力学、牵引电传动、控制等仿真模型与软件】'),
        ('试验能力', 'test-pipe', '【待填：*****、牵引／制动试验台、******等】'),
        ('数据积累', 'database', '【待填：机车试验数据与运用数据积累情况】'),
        ('在研项目', 'clipboard-list', '【待填：**级****项目数量与名称】'),
        ('人才队伍', 'users-group', '【待填：参研人员规模与职称结构】'),
    ]
    cw, ch = (1792 - 2 * 24) / 3.0, 232
    for i, (t, ic, ph) in enumerate(items):
        x = 64 + (i % 3) * (cw + 24)
        y = 214 + (i // 3) * (ch + 22)
        rect(s, x, y, cw, ch, fill=WHITE, line=LINE)
        icon_badge(s, ic, BLUE, x + 26, y + 24, 64)
        text(s, x + 108, y + 24, cw - 130, 64, t, size=26, bold=True, color=INK, anchor='m', lh=34, wrap=False)
        rect(s, x + 26, y + 108, cw - 52, ch - 132, fill=ORANGE.bg, line=ORANGE.main, lw=1.2, dash='dash')
        text(s, x + 46, y + 108, cw - 92, ch - 132, ph, size=19, color=ORANGE.text, anchor='m', lh=28)
    # 匹配性
    y1 = 724
    header(s, 64, y1, '与本课题的匹配性', w=1792)
    mt = [('控制系统与车辆动力学为牵头单位主营业务方向', '【可补充：具体型号或项目实例】'),
          ('具备支撑智能体研制、引擎开发与示范应用的仿真与试验条件', '【可补充：具体平台与台架】')]
    mw = (1792 - 24) / 2.0
    for i, (a, b) in enumerate(mt):
        x = 64 + i * (mw + 24)
        rect(s, x, y1 + 50, mw, 222, fill=GREEN.bg)
        rect(s, x, y1 + 50, 5, 222, fill=GREEN.main)
        icon(s, 'circle-check', GREEN.main, x + 28, y1 + 76, 34)
        text(s, x + 76, y1 + 72, mw - 104, 70, a, size=23, bold=True, color=INK, lh=32)
        rect(s, x + 76, y1 + 160, mw - 110, 72, fill=ORANGE.bg, line=ORANGE.main, lw=1.2, dash='dash')
        text(s, x + 96, y1 + 160, mw - 150, 72, b, size=19, color=ORANGE.text, anchor='m', lh=28)
    notes(s, '说明：本页需填入单位实际基础，它是“为何由我单位承担”的直接依据。\n橙色虚线框与【待填】/【可补充】为占位内容，请替换为牵头单位实际情况。')


# =====================================================================  24 成果介绍
def s24_results(d):
    s = d.new_slide()
    frame(d, s, '成果介绍：五项成果 ＋ 知识产权与标准', '调度引擎 · 仿真智能体／智能孪生模型 · 列车级数字样机 · 方法体系与技术规范 · 示范应用与验证报告', '成果介绍')
    top = [
        ('成果1', '软件1套', '多智能体协同运行与调度引擎', BLUE, 'settings-automation', [
            '完成<b>任务理解—任务编排—智能体执行</b>全流程：验证目标解析、环境自装配、保真度与算力调度、执行监控与重规划',
            '支持异构仿真智能体统一接入，<br>解决多专业模型“接不上”的问题',
            '面向性能验证提供<b>实时协同运行能力</b>']),
        ('成果2', '1套', '关键系统仿真智能体／<br>智能孪生模型', PURPLE, 'robot', [
            '覆盖控制系统、车辆动力学与线路等关键系统',
            '由“被动模型”升级为“<b>自主智能体</b>”，具备参数自动标定与结果判断能力',
            '具备标准化接口与可扩展接入机制，支持后续新增系统']),
        ('成果3', '1套', '列车级数字样机', GREEN, 'train', [
            '集成关键系统仿真智能体，形成<b>整车级验证环境</b>',
            '支持列车关键性能的多工况协同验证，并支撑<b>智能运维</b>应用',
            '预留制动、辅助系统等关键系统扩展接口']),
    ]
    pw, ph, y0 = (1792 - 48) / 3.0, 430, 214
    for i, (no, form, name, th, ic, bl) in enumerate(top):
        x = 64 + i * (pw + 24)
        rect(s, x, y0, pw, ph, fill=WHITE, line=LINE)
        rect(s, x, y0, pw, 6, fill=th.main)
        icon_badge(s, ic, th, x + 28, y0 + 30, 72)
        text(s, x + 118, y0 + 28, 300, 30, '<k>%s</k>　<c=%s>%s</c>' % (no, SUB, form), size=18, color=th.text, lh=26,
             wrap=False)
        text(s, x + 118, y0 + 58, pw - 118 - 28 - 96, 70, name, size=24, bold=True, color=INK, lh=32)
        num_mark(s, x + pw - 28 - 90, y0 + 22, 90, '0%d' % (i + 1), th.num, size=40)
        line(s, x + 28, y0 + 132, x + pw - 28, y0 + 132, LINE, 1.5)
        bullets(s, x + 28, y0 + 152, pw - 56, ph - 170, bl, th.main, size=20, lh=31, sa=12)
    # 成果4、5
    y1 = y0 + ph + 22
    LW = 920
    small = [('成果4', '1套', '列车级数字样机构建方法体系与技术规范', ORANGE, 'file-certificate',
              '形成构建方法、集成流程与标准化接口规范，可作为**内推广与后续课题复用的方法支撑'),
             ('成果5', '1份', '示范应用与验证报告', CYAN, 'report-analytics', '在典型机车产品上完成关键性能协同验证，形成验证效能评估结论')]
    sh = (996 - y1 - 14) / 2.0
    for i, (no, form, name, th, ic, txt) in enumerate(small):
        y = y1 + i * (sh + 14)
        rect(s, 64, y, LW, sh, fill=WHITE, line=LINE)
        rect(s, 64, y, 6, sh, fill=th.main)
        icon_badge(s, ic, th, 90, y + (sh - 60) / 2.0, 60)
        n = wrap_lines(txt, LW - 200, 19)
        top = y + (sh - (34 + 6 + n * 28)) / 2.0
        text(s, 172, top, LW - 200, 34, '<k><c=%s>%s</c></k>　%s　<c=%s>%s</c>' % (th.text, no, name, SUB, form),
             size=21, bold=True, color=INK, lh=30, wrap=False)
        text(s, 172, top + 40, LW - 200, n * 28 + 4, txt, size=19, color=BODY, lh=28)
    # 知识产权与标准
    R, RW = 1008, 848
    header(s, R, y1 - 4, '知识产权与标准', w=RW)
    ip = [('4', '项', '发明专利申请', '覆盖任务编排调度机制、保真度与算力调度、实时时序同步、工况自生成等方向'),
          ('2', '篇', '核心论文', '面向多智能体协同仿真与性能验证方法'),
          ('2', '项', '企业标准草案', '列车级数字样机构建与智能体接口相关规范')]
    kw_ = (RW - 2 * 12) / 3.0
    for i, (v, u, lab_, sub) in enumerate(ip):
        x = R + i * (kw_ + 12)
        kpi(s, x, y1 + 40, kw_, 996 - y1 - 40, '%s<s=20><n><c=%s> %s</c></n></s>' % (v, SUB, u), lab_, sub, color=RED,
            value_size=46, label_size=20, sub_size=16, pad=18)


# =====================================================================  25 结束页
def s25_end(d):
    s = d.new_slide()
    rect(s, 64, 34, 1792, 2, fill=RULE)
    rect(s, 64, 132, 8, 30, fill=RED)
    text(s, 88, 128, 1400, 38, '****“***”科技重大专项课题申报', size=25, bold=True, color=RED, cs=3, anchor='m',
         lh=34, wrap=False)
    text(s, 64, 262, 1792, 140, '恳请各位专家指导', size=104, bold=True, color=INK, lh=134, wrap=False)
    rect(s, 64, 432, 120, 8, fill=RED)
    text(s, 64, 476, 1792, 50, '**********', size=34, bold=True, color=INK, lh=46, wrap=False)
    text(s, 64, 538, 1792, 46, '7.3 轨道交通装备性能验证的多智能体协同运行与调度引擎技术研究', size=28, color=SUB, lh=40, wrap=False)
    ruler(s, 64, 944, 1792)
    loco_train(s, 560, 944, 428, n=3, gap=8, faded=False)
    rect(s, 64, 1010, 1792, 2, fill=RULE)
    sep = '<c=%s>   |   </c>' % RULE
    text(s, 64, 1024, 1300, 36, FOOTER + sep + '课题申报汇报' + sep + '谢谢', size=17, color=MUTED, bold=True, cs=2.5,
         anchor='m', lh=24, wrap=False)
    text(s, 1456, 1024, 400, 36, '%02d / %02d' % (s._no, d.total), size=17, color=MUTED, bold=True, font=MONO,
         align='r', anchor='m', lh=24, wrap=False)

SLIDES = [s01_cover, s02_agenda, s03_intro, s04_value, s05_urgency, s06_core, s07_layer1, s08_layer2,
          s09_layer3, s10_layer4, s11_methods, s12_route, s13_scenario, s14_innovation, s15_benchmark,
          s16_future, s17_schedule, s18_tasks, s19_goals, s20_metrics, s21_budget, s22_management,
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
