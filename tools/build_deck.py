# -*- coding: utf-8 -*-
"""根据《项目汇报PPT文本.md》（v6）生成课题申报汇报 PPT。

用法：python3 tools/build_deck.py [输出路径]
依赖：python-pptx、Pillow；图标首次生成需 node + tools/package.json 中的依赖（已生成的 PNG 在 assets/icons）。
每页“备注（讲稿）”写入 PPT 备注栏；示意图为形状组合，整组删除即可换成实际图片。
"""
import os
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from pptkit import *  # noqa: E402,F401,F403
from deck_components import *  # noqa: E402,F401,F403
import pptkit  # noqa: E402

OUT = os.path.join(ROOT, '项目汇报PPT.pptx')
LOCO = os.path.join(ROOT, 'assets', 'illustrations', LOCO_NAME + '.png')
LOCO_FADED = os.path.join(ROOT, 'assets', 'illustrations', LOCO_NAME + '_faded.png')
TOTAL = 21
FOOTER = '7.3 多智能体协同运行与调度引擎'
LAYER_NAME.update({1: '测试知识库与任务工作流', 2: '任务理解与智能编排调度', 3: '智能体执行', 4: '架构融合与协同运行'})


# 经费明细（第 3、17 页共用）：序号、科目、金额、测算依据、色块、条内标签
BUDGET = [
    ('1', '直接投入费用', 50, 'AI算力服务器（GPU+NPU异构）1台 40万；仿真与调度一体化工作站 2台 10万', M1.main, '直接投入 50'),
    ('2', '人员人工费用', 30, '按投入人月测算：<br><bc=%s>【待填】人·月 × 【待填】万元/人·月</bc>' % PH.text, BUDGET2, '人员 30'),
    ('3', '固定资产相关费用（折旧）', 5, '现有仿真与测试设备折旧分摊', GRAYS[0], None),
    ('4', '试验检验及试制外协费用', 5, '建模数据采集处理 2万；模型校验与第三方验证 3万', GRAYS[1], None),
    ('5', '研发成果相关费用', 5, '发明专利申请 2万；论文与标准草案编制 3万', GRAYS[2], None),
    ('6', '与研发活动直接相关的<br>其他费用', 5, '差旅费 3万；会议费 2万', GRAYS[3], None),
]

# 研究阶段（第 14、18 页共用）：名称、简称、起止月（0—24）、色、起止时间、阶段目标、阶段交付、里程碑、日期、里程碑名
STAGES = [
    ('阶段一', '架构与规范', 0, 6, STAGE[0], '2027.01—2027.06', '完成总体架构与接口规范设计', '总体架构、验证判据知识化方法与标准化接口规范',
     'M1', '2027.06', '架构与接口规范'),
    ('阶段二', '智能体与引擎', 6, 15, STAGE[1], '2027.07—2028.03', '突破实时协同与保真度调度技术，研制智能体与引擎原理样机',
     '关键系统仿真智能体／智能孪生模型、调度引擎原理样机', 'M2', '2028.03', '智能体与引擎'),
    ('阶段三', '样机与验证', 15, 21, STAGE[2], '2028.04—2028.09', '建成 FXN5C 列车级数字样机', '列车级数字样机、典型工况协同验证结果', 'M3',
     '2028.09', '数字样机与验证'),
    ('阶段四', '示范与验收', 21, 24, STAGE[3], '2028.10—2028.12', '在临哈线典型区段完成示范应用', '示范应用与效能评估、方法体系与技术规范、标准草案',
     'M4', '2028.12', '示范应用与验收'),
]


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
    text(s, 64, 474, 1792, 40, '所属项目：7. 基于智能体的轨道交通装备性能数字样机关键技术研究', size=26, color=SUB,
         anchor='m', lh=40, wrap=False)
    text(s, 64, 528, 1792, 50, '让现场才暴露的问题，在研发阶段先跑出来。', size=34, bold=True, color=RED, anchor='m', lh=50,
         wrap=False)
    cells = [('牵头单位', '**********'), ('项目负责人', '***'), ('申报层级', '课题级'),
             ('研究周期', '2027年1月—2028年12月')]
    ws = [330, 270, 250, 520]
    x = 64
    for i, ((lab, val), w) in enumerate(zip(cells, ws)):
        xx = x + (30 if i else 0)
        if i:
            rect(s, x, 636, 2, 88, fill=LINE)
        text(s, xx, 632, w - 30, 32, lab, size=20, bold=True, color=MUTED, lh=32, cs=1, wrap=False)
        text(s, xx, 670, w - 30, 54, val, size=32, bold=True, color=INK, lh=54, wrap=False)
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
    frame(d, s, '汇报提纲：七个部分，核心是研究内容、技术路线与方法', '共 21 页 · 第三部分“主要研究内容、技术路线与方法”为核心章节（第 7—14 页）',
          '汇报提纲')
    rows = [
        ('01', '项目简介', '项目基本信息 · 项目子课题及预计成果物 · 项目预计经费', '第 3 页', 'file-description'),
        ('02', '研究意义与价值', '背景与目的 · 应用价值与效益 · 为什么是现在', '第 4—6 页', 'bulb'),
        ('03', '主要研究内容、技术路线与方法', '总体路线 · 研究内容①—④ · 典型场景 · 创新点 · 阶段与检验', '第 7—14 页', 'route'),
        ('04', '预期成果与效益', '预期成果与考核方式 · 量化指标与效益评估', '第 15—16 页', 'target-arrow'),
        ('05', '经费预算合理性', '预算明细与测算依据 · 分年度 · 合理性说明', '第 17 页', 'coins'),
        ('06', '项目管理计划', '任务分配 · 项目流程 · 时间安排 · 执行保障 ｜ 团队研究能力', '第 18—19 页', 'sitemap'),
        ('07', '成果介绍', '五项成果 ＋ 知识产权与标准', '第 20 页', 'trophy'),
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
    lab = dict(bold=True, color=SUB, size=17)
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
    header(s, L, y1, '预计成果物', note='数字样机由智能体经引擎集成而成', w=LW)
    items = [
        ('多智能体协同运行与调度引擎', '软件1套', M1, 'settings-automation'),
        ('关键系统仿真智能体／<br>智能孪生模型', '模型1套', M2, 'robot'),
        ('列车级数字样机', '系统1套', M3, 'train'),
        ('列车级数字样机构建方法体系与技术规范', '规范1套', M3, 'file-certificate'),
        ('示范应用与验证报告', '报告1份', M4, 'report-analytics'),
        ('发明专利／论文／<br>企业标准草案', '4项／2篇／2项', M4, 'certificate'),
    ]
    cw, gx, gy = (LW - 24) / 3.0, 12, 12
    ch = (990 - 122 - 22 - 46 - gy - (y1)) / 2.0   # 使经费块底边落在 990
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
    header(s, R, 212, '项目子课题与本课题边界', note='项目7设 5 个子课题 · 本课题攻关“引擎”', w=RW)
    rows = [
        ('7.1', '高质量数据集', '采用其数据规范与接口，不承担建库（7.1 建库、7.3 按需生成）', 'n'),
        ('7.2', '多物理域建模与仿真智能体', '只建验证必需的智能体（控制系统、车辆动力学、线路），不覆盖全部关键系统', 'n'),
        ('7.3', '协同运行与调度引擎（本课题）',
         '<b>核心攻关引擎</b>，解决多专业智能体“接得上、跑得动、调得灵”；数字样机是引擎承载智能体后的集成产物<br>'
         '分 <b>4 个子任务</b>：知识 → 调度 → 执行 → 集成', 'me'),
        ('7.4', '数字样机构建', '承接智能体集成与协同运行', 'n'),
        ('7.5', '方案优化与智能生成', '为优化提供验证支撑', 'n'),
        ('6.2', '通用智能体调度平台', '本课题为<b>面向性能验证的领域化编排</b>，强调实时协同与多物理域时序', 'ext'),
    ]
    nb, tgap = 76, 20
    tw_ = RW - nb - tgap - 22
    ext_gap = 20
    groups = [34 + 8 + wrap_lines(desc, tw_, 19) * 29 for _, _, desc, _ in rows]
    pad = (990 - 256 - 10 * (len(rows) - 1) - ext_gap - sum(groups)) / float(len(rows))
    y = 256
    for (no, name, desc, kind), g in zip(rows, groups):
        h = g + pad
        if kind == 'ext':
            line(s, R, y + 5, R + RW, y + 5, RULE, 1.5, dash='dash')
            y += ext_gap
        me = (kind == 'me')
        rect(s, R, y, RW, h, fill=RED_BG if me else WHITE, line=RED if me else LINE, lw=2 if me else 1,
             dash='dash' if kind == 'ext' else None)
        shp = rect(s, R, y, nb, h, fill=RED if me else WHITE)
        if not me:
            line(s, R + nb, y + 12, R + nb, y + h - 12, LINE, 1)
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
    notes(s, '备注（讲稿）：边界细节——7.1 采用其数据规范与接口，不承担建库；7.2 仅自建验证必需的智能体，不覆盖弓网/牵引/制动等全部关键系统；'
             '7.4 承接智能体集成与协同运行；7.5 为优化提供验证支撑；6.2 通用智能体调度平台是通用编排，本课题强调实时协同与多物理域时序。'
             '数字样机由关键系统仿真智能体／智能孪生模型经调度引擎集成而成：智能体可单独核验、复用，数字样机作为整车级验证系统交付。')


# =====================================================================  04 研究背景与重要性
def s04_background(d):
    s = d.new_slide()
    frame(d, s, '研究背景与重要性：一次验证 40 人、2 个月，问题仍到现场才暴露', '验证现状 · 真实案例 · 症结 · 研究目的', '研究意义与价值')
    L, LW = 64, 872
    big_numbers(s, L, 214, LW, [('40', '人', '一次投入'), ('<s=18><n>约</n></s>2', '个月', '周期'), ('−10', '℃', '案例环境')],
                h=84, label_w=90)
    text(s, L, 304, LW, 26, '40 人 ＝ 研发 20 ＋ 测试 10 ＋ 现场 10', size=16, color=MUTED, lh=24, wrap=False)
    header(s, L, 346, '验证现状', w=LW)
    text(s, L, 390, LW, 64, '控制软件每次迭代：HIL 台架＋经验数据模拟运用工况，重点工况再到现场验证。', size=20, color=BODY, lh=31)
    header(s, L, 470, '即便如此，问题还是在现场才暴露', w=LW)
    rect(s, L, 514, LW, 184, fill=RED_BG)
    rect(s, L, 514, STYLE['concl_bar'], 184, fill=RED)
    icon(s, 'quote', MUTED, L + 24, 536, 32)
    text(s, L + 72, 514, LW - 100, 184,
         '2024 年 12 月，FXN5C 在临哈线长大下坡、约 −10℃：回手柄后风扇仍以 40Hz 运行，低温水 25℃ → 0℃，高温水 78℃，'
         '<b>散热器冻结</b>。多系统耦合的长时序问题，台架没有复现出来。', size=20, color=INK, anchor='m', lh=32)
    R, RW = 976, 880
    header(s, R, 214, '症结', w=RW)
    cards = [('复杂工况模拟不了', '多系统耦合、长时序、低温／长大坡道／湿轨组合，台架覆盖不到'),
             ('验证靠人，跟不上迭代', '工况人工编、环境人工搭、结果人工看'),
             ('经验散在人手里', '模型各自为战，判据分散在标准、大纲和专家经验中')]
    for i, (t, b) in enumerate(cards):
        problem_card(s, R, 258 + i * 152, RW, 136, t, b, title_size=24, body_size=19, pad=26)
    summary(s, 64, 770, 1792, 150, '研究目的：让控制软件每改一版，复杂运用工况都能在数字样机上自动跑一遍、自动判读',
            '问题在研发阶段暴露，现场只做确认。', title_size=27, desc_size=22, gap=8)
    notes(s, '备注（讲稿）：技术表述——研制多智能体协同运行与调度引擎，以 FXN5C 为对象建成列车级数字样机，实现列车关键性能的快速、可复现、全覆盖验证。'
             '冻结案例：控制、冷却、动力、线路四个系统耦合，且是回手柄后持续一段时间才出现的长时序过程，HIL 只有控制器在环、其他系统靠经验数据近似，'
             '因此复现不出来。')


# =====================================================================  05 应用价值与社会经济效益
def s05_value(d):
    s = d.new_slide()
    frame(d, s, '应用价值：一句话需求进去，全工况报告出来', '用起来是什么样 · 用在三个环节 · 社会经济效益', '研究意义与价值')
    header(s, 64, 212, '用起来是什么样', note='以一次控制软件迭代为例', w=1792)
    steps = [('输入', '“验证新版软件在临哈线全工况下的速度跟踪与牵引能耗”', 'message'),
             ('引擎自动', '生成工况 → 装配智能体 → 分配精度与算力 → 协同求解 → 对标判据', 'settings-automation'),
             ('输出', '验证报告与异常定位，全程可回放；工程师只看异常项', 'report-analytics')]
    cw = (1792 - 2 * 40) / 3.0
    for i, (t, b, ic) in enumerate(steps):
        x = 64 + i * (cw + 40)
        rect(s, x, 256, cw, 158, fill=WHITE, line=LINE)
        rect(s, x, 256, cw, BAR, fill=M1.main)
        icon_badge(s, ic, M1, x + 24, 280, 56)
        text(s, x + 96, 280, cw - 120, 56, '<c=%s><m>0%d</m></c>  %s' % (M1.text, i + 1, t), size=23, bold=True, color=INK,
             anchor='m', lh=30, wrap=False)
        text(s, x + 24, 346, cw - 48, 60, b, size=18, color=BODY, lh=27)
        if i < 2:
            tri(s, x + cw + 20, 335, 16, 22, TRI)
    header(s, 64, 440, '用在三个环节', w=1792)
    rows = [['<b>研发：软件迭代</b>', '台架＋现场验证', '全工况在数字样机上自动回归，减少现场验证'],
            ['<b>试验：上线前预试验</b>', '低温等工况受季节与线路限制', '按试验大纲先在数字样机上预演，缩减上线项目'],
            ['<b>运用：问题复现</b>', '靠经验排查', '用运行记录复现、定位原因，支撑智能运维']]
    table(s, 64, 484, [340, 560, 892], rows, header=['环节', '现状', '应用引擎后'], size=19, lh=29, pad_y=17,
          col_styles={0: dict(color=INK), 2: dict(color=INK)})
    header(s, 64, 754, '社会经济效益', w=1792)
    effs = [('经济', 'coins', '减少现场验证与线路试验投入，缩短迭代周期'), ('社会', 'shield-check', '复杂工况覆盖度与可信度提升，减少运用故障'),
            ('产业', 'building-factory-2', '方法体系可向制动、辅助系统扩展，在**内主机企业推广')]
    ew = (1792 - 2 * 40) / 3.0
    for i, (t, ic, b) in enumerate(effs):
        x = 64 + i * (ew + 40)
        rect(s, x, 798, ew, 110, fill=M3.bg)
        icon(s, ic, M3.main, x + 24, 822, 32)
        text(s, x + 72, 812, 200, 30, t + '效益', size=20, bold=True, color=M3.text, lh=28, wrap=False)
        text(s, x + 72, 846, ew - 96, 54, b, size=18, color=BODY, lh=27)
    notes(s, '备注（讲稿）：对领域发展的推动作用——性能验证从“单系统、离线、靠现场”转向“列车级、实时协同、研发阶段可回归”，成为研发体系中可持续使用的基础能力；'
             '输出的构建方法体系与技术规范，主机企业可按同一流程构建自己的数字样机；新增系统按标准接口接入智能体即可，引擎不需重构。')


# =====================================================================  06 为什么是现在
def s06_now(d):
    s = d.new_slide()
    frame(d, s, '立项必要性与紧迫性：为什么是现在', '技术条件 · 工程需求 · 自主可控 ｜ 三个条件刚刚同时具备', '研究意义与价值')
    items = [('技术条件刚成熟', 'cpu', M1, '智能体技术让“任务理解—编排—执行”可自动化；GPU＋NPU 异构算力让实时协同与全保真计算可负担'),
             ('工程需求在加压', 'trending-up', M2, '软件迭代加快<bc=%s>【可补充：近年迭代次数】</bc>，复杂工况问题增多，“台架＋现场”已跟不上' % PH.text),
             ('自主可控有窗口', 'shield-check', M3, '研发设计类工业软件国产化率约 5%；验证引擎这一层现在自己做，就掌握在自己手里')]
    cw = (1792 - 2 * 32) / 3.0
    for i, (t, ic, th, b) in enumerate(items):
        x = 64 + i * (cw + 32)
        yy = panel(s, x, 214, cw, 430, th, t, '0%d' % (i + 1), ic, title_size=28, badge=72)
        text(s, x + 40, yy + 24, cw - 80, 220, b, size=22, color=BODY, lh=36)
    summary(s, 64, 720, 1792, 150, '结论：本课题要做的是<r>验证环节的自主引擎</r>', '晚一年布局，就多一年被动。', title_size=28, desc_size=22, gap=8)
    notes(s, '备注（讲稿）：《“***”铁路科技创新规划》（****〔****〕**号）将数字孪生列为智能铁路关键技术，要求开展数字孪生平台研发应用，“到****年智能铁路技术全面突破”；'
             '用户端年软件支出达数百万元且年均涨价约 5%；指南已将“研发设计软件受制于人”“设计仿真数据割裂”列为待解决事项。'
             '数字孪生与智能体技术正从概念验证走向工程落地，等国外工具链先做成，验证环节将重演工业软件受制于人的路径。')


# =====================================================================  07 总体技术路线
def s07_route(d):
    s = d.new_slide()
    frame(d, s, '总体技术路线：四层引擎，把验证工作流交给智能体', '目标：像软件回归测试一样——设计改一次，全工况在数字样机上自动重跑一遍', '总体技术路线')
    L, LW = 64, 952
    header(s, L, 212, '四层架构与分工', note='第 8—11 页逐层展开', w=LW)
    roles = {1: '标准、大纲、案例 → 可自动比对的判据、可复用的流程', 2: '一句话需求 → 任务；自动装配智能体，按工况分配精度与算力',
             3: '跑工况、自动判读、生成报告、结果回流', 4: '多系统智能体在可复现沙箱中协同运行，形成列车级数字样机'}
    y, bh, gap = 256, 100, 10
    for i in (4, 3, 2, 1):
        th = LAYER[i]
        rect(s, L, y, LW, bh, fill=WHITE, line=LINE)
        rect(s, L, y, 250, bh, fill=th.bg)
        rect(s, L, y, BAR, bh, fill=th.main)
        text(s, L + 26, y + 16, 214, 34, LAYER_NO[i] + ' ' + LAYER_KIND[i], size=24, bold=True, color=th.text, lh=34, wrap=False)
        text(s, L + 26, y + 54, 218, 30, LAYER_NAME[i], size=17, bold=True, color=INK, lh=26, wrap=False)
        text(s, L + 276, y, LW - 300, bh, roles[i], size=20, color=BODY, anchor='m', lh=31)
        y += bh + gap
    R, RW = 1056, 800
    header(s, R, 212, '接入现有研发验证体系', note='不另起炉灶', w=RW)
    rows = [['设计规范、试验大纲、问题记录', '提取为判据与流程模板（①）'], ['已有仿真模型、HIL 台架', '经标准接口封装接入，不重复建模（③④）'],
            ['型式试验、线路试验与运用数据', '智能体自动标定与结果可信性评判（③）'], ['设计评审、软件放行、故障分析', '验证报告与复现结论直接作为依据']]
    yy = table(s, R, 256, [360, 440], rows, header=['现有资源与环节', '衔接方式'], size=18, lh=27, pad_y=13,
               col_styles={0: dict(color=INK, bold=True)})
    header(s, R, 600, '闭环', note='全程留痕，可回放、可回归', w=RW)
    chips_row(s, R, 646, ['判据与流程', '分解与调度', '协同执行与判读', '报告', '数据回归'], gap=26, arrow=True, arrow_color=TRI,
              size=16, h=40, padx=12, color=INK, line=LINE, bold=True)
    summary(s, 64, 760, 1792, 150, '怎么算成功：阶段三，数字样机复现临哈线散热器冻结工况并定位原因；阶段四，一次软件迭代完成全工况回归',
            '验证周期由约 2 个月缩短至<bc=%s>【待填】</bc>；各阶段检验见第 14 页。' % PH.text, title_size=25, desc_size=21, gap=8)
    notes(s, '备注（讲稿）：与课题要求的对应——高效实时协同运行 → ③④；调度引擎 → ②；列车级数字样机及构建方法体系 → ④；支持关键性能验证与智能运维 → ④。'
             '对接项目7总体目标——模型自动标定 → ③；指标实时推演 → ④；性能自动评估 → ③。')


# =====================================================================  08 研究内容① 基础层
def s08_layer1(d):
    s = d.new_slide()
    th = LAYER[1]
    frame(d, s, '研究内容① 测试知识库与任务工作流：把判据和流程从人手里拿出来', '现状与问题 · 研究内容 · 研究方法 · 示意图', '研究内容① 基础层')
    content_left(s, 1, '与上层的关系', '知识库与工作流是②“任务理解与编排调度”的<b>知识与规则来源</b>。',
                 outs=['验证判据知识库', '任务工作流定义<br>与配置规范'])
    X, XW = 356, 1500
    scene_band(s, X, 214, XW, 106,
               '判定一个工况是否合格，依据分散在国标、行标、企标、试验大纲和专家经验里，查找比对靠人工。<bc=%s>【可替换为实际案例】</bc>' % PH.text,
               '<b>知识工程</b>——标准规范形式化、领域知识建模、工作流建模；难点是判据多以经验和文档形式存在。',
               scene_label='现状与问题', role_label='研究方法')
    cw = (XW - 48) / 3.0
    cards = [
        ('验证判据知识化', '从规范、大纲与历次问题记录中提取判据，形式化为<b>可自动比对</b>的条目；首批覆盖速度跟踪、牵引能耗、运行平稳性'),
        ('任务工作流建模', '定义任务的分解、顺序、流转条件与异常分支，形成<b>可配置</b>的流程模板'),
        ('知识与流程迭代', '验证结果与新问题回流，持续补充判据、修订流程'),
    ]
    for i, (t, desc) in enumerate(cards):
        point_card(s, X + i * (cw + 24), 330, cw, 148, '0%d' % (i + 1), t, desc, th, desc_size=18)
    # 附图：判据知识化（上）→ 任务工作流（下）→ 结果回流（右）
    f, py = fig_area(s, X, 496, XW, 494, '从规范到判据、从判据到工作流')
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
    notes(s, '备注（讲稿）：没有判据知识库，任务理解只能靠通用大模型“猜”；没有工作流，编排调度无章可循。判据条目示例：验证对象——牵引控制系统；'
             '关键性能——速度跟踪偏差；适用工况——长大下坡、湿轨；判定规则——|实测 − 目标| ≤ 允许偏差；来源——设计规范条款。\\n' + FIG_NOTE)


# =====================================================================  09 研究内容② 决策层
def s09_layer2(d):
    s = d.new_slide()
    th = LAYER[2]
    frame(d, s, '研究内容② 任务理解与智能编排调度：引擎在此落地', '现状与问题 · 研究内容 · 研究方法 · 示意图',
          '研究内容② 决策层')
    content_left(s, 2, '上下游关系', '向下承接①的<b>知识与规则</b>，向③下发<b>执行方案</b>（智能体组合、工况、算力）。',
                 outs=['多智能体协同运行与调度引擎（软件1套）'])
    X, XW = 356, 1500
    scene_band(s, X, 214, XW, 106,
               '工况是区段 × 温度 × 轨面 × 载重 × 手柄序列的组合，全用高精度模型算不完，全用简化模型关键工况又不可信。<bc=%s>【可替换为实际案例】</bc>' % PH.text,
               '<b>智能调度</b>——任务分解与匹配、多目标寻优、保真度—算力协同配置；难点是在线权衡精度与效率，极端工况不能外推。',
               scene_label='现状与问题', role_label='研究方法')
    cw = (XW - 60) / 4.0
    cards = [
        ('目标解析与任务分解', '一句话需求 → 验证对象、关键性能、判据条目，匹配知识库'),
        ('测试环境自动装配', '自动完成智能体组合、接口连接与算力配置'),
        ('保真度与算力自适应调度', '常规工况降阶提速，关键与极端工况强制全保真；GPU 跑多物理域并行仿真，NPU 做推理与编排决策'),
        ('执行监控与失败重规划', '异常中断后自动重规划，支撑指标实时推演'),
    ]
    for i, (t, desc) in enumerate(cards):
        point_card(s, X + i * (cw + 20), 330, cw, 150, '0%d' % (i + 1), t, desc, th)
    # 附图：一次验证任务中的调度过程
    f, py = fig_area(s, X, 496, XW, 494, '一次验证中的调度过程：按工况切换保真度、分配算力，异常时重规划')
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
    notes(s, '备注（讲稿）：算力为本课题配置的 GPU＋NPU 异构服务器 1 台、一体化工作站 2 台。\\n' + FIG_NOTE)


# =====================================================================  10 研究内容③ 执行层
def s10_layer3(d):
    s = d.new_slide()
    th = LAYER[3]
    frame(d, s, '研究内容③ 智能体执行：把 −10℃ 现场的问题在研发室里跑出来', '一条执行链：仿真智能体 → 用例生成 → 测试识别 → 报告与回归', '研究内容③ 执行层')
    content_left(s, 3, '对接项目7总体目标', [{'t': '<b>模型自动标定</b> → 仿真智能体', 'sa': 8}, {'t': '<b>性能自动评估</b> → 测试识别'}],
                 outs=['关键系统仿真智能体／<br>智能孪生模型', '用例与线路数据集', '验证报告与追溯记录'])
    X, XW = 356, 1500
    scene_band(s, X, 214, XW, 106,
               '仿真结果靠人看曲线，多系统、长时序问题（如临哈线散热器冻结）难发现、难定位。',
               '<b>实时协同与可信评判</b>——分布式协同仿真、时序同步、自动标定与不确定度量化；难点是步长差异大、极端工况数据少。',
               scene_label='现状与问题', role_label='研究方法')
    gap = 24
    cw, ch, y0 = (XW - 3 * gap) / 4.0, 206, 330
    steps = [
        ('仿真智能体', '研制与接入', '控制系统、车辆动力学、线路等智能体与孪生模型，标准化接入、时序同步，用试验与运用数据<b>自动标定</b>'),
        ('用例生成', '按需自生成', '按线路资料与覆盖要求生成用例与线路数据（平纵断面、曲线、超高、坡度、不平顺谱）'),
        ('测试识别', '结果评判 · 异常定位', '逐工况对标判据，<b>识别异常、定位原因</b>，给出可信性判断'),
        ('报告与回归', '可追溯 · 回归重跑', '结论自动汇总、可追溯；结果回流迭代标定与判据，设计变更后<b>回归重跑</b>'),
    ]
    cxs = []
    for i, (t, sub, desc) in enumerate(steps):
        x = X + i * (cw + gap)
        cxs.append(x + cw / 2.0)
        point_card(s, x, y0, cw, ch, '0%d' % (i + 1), t, desc, th, sub=sub)
        if i < 3:
            tri(s, x + cw + gap / 2.0, y0 + ch / 2.0, 14, 20, TRI)
    yb = y0 + ch
    poly(s, [(cxs[3], yb), (cxs[3], yb + 22), (cxs[0], yb + 22), (cxs[0], yb + 2)], OK.main, 2, tail='triangle')
    lab = '数据回归：结果回流，迭代模型标定、判据与工况集 · 设计变更后回归重跑'
    lw_ = text_width(lab, 16, True) + 28
    lx_ = (cxs[0] + cxs[3]) / 2.0 - lw_ / 2.0
    rect(s, lx_, yb + 10, lw_, 24, fill=WHITE)
    text(s, lx_, yb + 10, lw_, 24, lab, size=16, bold=True, color=OK.text, align='c', anchor='m', lh=22, wrap=False)
    # 附图：用现场案例说明测试识别
    f, py = fig_area(s, X, 584, XW, 406, '以临哈线散热器冻结为例：逐工况对标判据，识别异常并定位原因', note='示意图 · 数据据现场记录 · 防冻限值【待填】')
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
    notes(s, '备注（讲稿）：标定数据来自型式试验、线路试验与运用记录。测试识别输出结果可信性判断，是为了解决极端工况实测数据少、结论难以自证的问题。'
             '设计修改后，低温工况全量回归重跑。示意图：数据据 2024 年 12 月临哈线现场记录，曲线为示意；防冻限值线仅示意位置，请按设计规范取值；'
             '如有现场记录曲线，选中附图组合删除后替换即可。')


# =====================================================================  11 研究内容④ 集成层
def s11_layer4(d):
    s = d.new_slide()
    th = LAYER[4]
    frame(d, s, '研究内容④ 架构融合与协同运行：融合成一个能跑起来的整体', '现状与问题 · 研究内容 · 研究方法 · 示意图', '研究内容④ 集成层')
    content_left(s, 4, '对接项目7总体目标', '<b>指标实时推演</b> → 协同运行',
                 outs=['列车级数字样机', '列车级数字样机构建方法体系与集成技术规范'])
    X, XW = 356, 1500
    scene_band(s, X, 214, XW, 106,
               '台架以控制器为核心，其他系统靠经验数据近似，控制、冷却、动力、线路等系统难以在同一环境中协同运行，复杂场景无法模拟。',
               '<b>系统集成</b>——标准化接口、服务化封装、沙箱化运行与全程留痕；难点是异构模型实时协同下的稳定性与可复现。',
               scene_label='现状与问题', role_label='研究方法')
    cw = (XW - 60) / 4.0
    cards = [
        ('总体架构与接口规范', '定义四层的<b>接口、数据流与控制流</b>；已有模型与 HIL 台架经标准接口接入，不重复建设'),
        ('多智能体协同运行', '异构智能体<b>实时协同求解、时序同步与状态同步</b>'),
        ('可复现沙箱', '<b>隔离</b>（单路发散不扩散）、<b>记录</b>（参数、模型版本、随机种子）、<b>回放</b>（任意一次验证可复现）'),
        ('列车级数字样机', '以 <b>FXN5C</b> 为示范对象集成关键系统智能体，支持性能验证与智能运维'),
    ]
    for i, (t, desc) in enumerate(cards):
        point_card(s, X + i * (cw + 20), 330, cw, 150, '0%d' % (i + 1), t, desc, th)
    # 附图：数字样机构成
    f, py = fig_area(s, X, 496, XW, 494, '列车级数字样机构成：关键系统智能体经标准化接口接入引擎，在沙箱中协同运行')
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


# =====================================================================  12 典型应用场景
def s12_scenario(d):
    s = d.new_slide()
    frame(d, s, '典型应用场景：FXN5C 临哈线控速＋能耗验证，全流程交给智能体', '输入一句话需求 → 引擎自动完成五步 → 验证报告 ＋ 全要素留痕', '典型应用场景')
    y0 = 214
    shp = rect(s, 64, y0, 250, 84, fill=INK)
    text(s, 64, y0, 250, 84, '输入 · 一句话需求', size=22, bold=True, color=WHITE, align='c', anchor='m', lh=30, shp=shp, wrap=False)
    rect(s, 314, y0, 1542, 84, fill=PANEL)
    icon(s, 'quote', MUTED, 340, y0 + 24, 36)
    text(s, 394, y0, 1440, 84, '验证 FXN5C 在临哈线上按给定运行图运行时的<r>速度跟踪</r>与<r>牵引能耗</r>是否满足要求。', size=26, bold=True,
         color=INK, anchor='m', lh=36, wrap=False)
    header(s, 64, 330, '引擎自动完成', w=1792)
    steps = [('任务理解', '对象——牵引控制系统；性能——速度跟踪偏差、牵引能耗；判据取自知识库'),
             ('工况生成', '临哈线线路数据 × 轨面（干／湿）× 载荷（满／空）× 线路条件（长大坡道／小半径曲线／隧道）'),
             ('装配与调度', '牵引控制、车辆动力学、线路智能体自动装配，按工况分配精度与算力'),
             ('执行与判读', '协同求解、自动标定；定位异常（如某曲线段速度超差），追溯原因'),
             ('输出', '验证报告＋全要素留痕，可回放、可回归')]
    gap = 24
    cw = (1792 - 4 * gap) / 5.0
    for i, (t, b) in enumerate(steps):
        x = 64 + i * (cw + gap)
        point_card(s, x, 374, cw, 240, '0%d' % (i + 1), t, b, M1, desc_size=18)
        if i < 4:
            tri(s, x + cw + gap / 2.0, 494, 14, 20, TRI)
    header(s, 64, 654, '与现有做法对比', w=1792)
    rows = [['<b>工况准备</b>', '人工编写', '按线路资料与判据自动生成'], ['<b>环境搭建</b>', '台架＋人工联调', '自动装配智能体与算力'],
            ['<b>结果判读</b>', '人工看曲线', '自动比对判据、定位原因']]
    table(s, 64, 698, [340, 560, 892], rows, header=['环节', '现状', '应用引擎后'], size=19, lh=29, pad_y=15,
          col_styles={2: dict(color=INK)})
    notes(s, '备注（讲稿）：这一次验证里，工况没人手写、线路数据没人准备、环境没人搭、异常没人看图判断。追溯原因示例：黏着利用不足？控制参数不匹配？')


# =====================================================================  13 创新点与技术定位
def s13_innovation(d):
    s = d.new_slide()
    frame(d, s, '创新点与技术定位：先进性不在单项技术，而在调度的维度与粒度', '三项突破 · 技术定位：不替代现有工具，把它们接起来、调度起来', '创新点')
    header(s, 64, 212, '三项突破', w=1792)
    rows = [['<b>① 调度对象升级为智能体</b>', '智能体自主标定、自主判读，引擎按任务调度智能体而非算例', '<b>人退出执行回路</b>'],
            ['<b>② 调度多一个维度：保真度</b>', '工况驱动的降阶／全保真切换，GPU、NPU 算力随之调配', '<b>批量提速与关键工况精度同时成立</b>'],
            ['<b>③ 调度基础是可复现沙箱</b>', '参数、模型版本、随机种子全程留痕与回放', '<b>现场问题可复现，改进后可回归</b>']]
    yy = table(s, 64, 256, [440, 860, 492], rows, header=['突破', '技术机制', '带来的变化'], size=20, lh=30, pad_y=19,
               col_styles={0: dict(color=RED), 2: dict(color=INK)})
    header(s, 64, yy + 34, '技术定位：不替代现有工具，把它们接起来、调度起来', w=1792)
    rows = [('商用仿真软件', '单专业高精度', '列车级协同'), ('协同仿真标准', '模型封装与交换', '智能编排、实时性'),
            ('HIL 试验台', '控制器实时验证', '工况受限、不可批量'), ('线路试验', '结果权威', '周期长、不可复现'),
            ('降阶模型', '提速', '极端工况不能外推'), ('通用智能体框架', '任务编排', '多物理域实时协同')]
    trs = [['<b>%s</b>' % a, '<c=%s>✓</c>  %s' % (OK.main, b), '<bc=%s>×</bc>  %s' % (RED, c)] for a, b, c in rows]
    cw = (1792 - 40) / 2.0
    for k in (0, 1):
        table(s, 64 + k * (cw + 40), yy + 78, [300, 280, cw - 580], trs[k * 3:(k + 1) * 3], header=['现有手段', '能做', '做不到'],
              size=19, lh=28, pad_y=14, col_styles={1: dict(color=OK.text), 2: dict(color=BODY)})
    summary(s, 64, 862, 1792, 108, '在成熟技术之上，补上“<r>面向列车级性能验证的智能编排与实时协同</r>”这一层',
            '不做仿真软件、不做接口标准、不做试验台。', title_size=25, desc_size=20, gap=6)
    notes(s, '备注（讲稿）：与同类工作的差异——现有工作多聚焦单一系统仿真或通用智能体平台，本课题把协同运行与调度做成可交付的引擎和可复用的方法体系。')


# =====================================================================  14 阶段目标与进度

# 各里程碑的检验（考题）
CHECKS = {
    'M1': '规范通过评审；判据首批覆盖三项关键性能',
    'M2': '原理样机完成<bc=%s>【待填】</bc>个工况自动装配与调度；智能体与实测比对达标' % PH.text,
    'M3': '<b>复现临哈线散热器冻结工况并定位原因</b>',
    'M4': '一次软件迭代完成全工况回归，验证周期由约 2 个月缩短至<bc=%s>【待填】</bc>；第三方验证' % PH.text,
}


def s14_stages(d):
    s = d.new_slide()
    frame(d, s, '阶段目标与进度：四个阶段，每个里程碑一道考题', '研究周期 2027.01—2028.12（24 个月） · 完成 → 突破 → 建成 → 实现', '阶段目标与进度')
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
    # 目标 + 考题
    rows = []
    for name, short, a, b, col, per, goal, deli, mk, mdate, mtext in stages:
        rows.append(['<b>%s</b><br><g>%s</g>' % (name, per), goal, CHECKS[mk]])
    table(s, 64, 470, [300, 700, 792], rows, header=['阶段', '目标', '检验（考题）'], size=19, lh=29, pad_y=13,
          col_styles={0: dict(color=INK), 2: dict(color=INK)})
    notes(s, '备注（讲稿）：阶段交付——一：总体架构、判据知识化方法、接口规范；二：关键系统仿真智能体／智能孪生模型、调度引擎原理样机；'
             '三：列车级数字样机、典型工况协同验证结果；四：示范应用与效能评估、方法体系与技术规范、标准草案。前一阶段交付物即后一阶段输入。')


# =====================================================================  15 预期成果与考核方式
def s15_results_assess(d):
    s = d.new_slide()
    frame(d, s, '预期成果与考核方式：四项目标，每项有成果物和考核方式', '目标 → 成果物 → 考核方式 ｜ 知识产权与标准', '预期成果与效益')
    rows = [['<b>1 调度引擎</b>：实现异构智能体实时协同运行与算力动态调度', '引擎软件 1 套', '软件演示、功能核验'],
            ['<b>2 数字样机与性能验证</b>：建成 FXN5C 数字样机，实现速度跟踪、牵引能耗、运行平稳性等关键性能的多工况协同验证',
             '智能体／孪生模型 1 套<br>数字样机 1 套（由智能体集成而成）', '智能体——与实测比对<br>数字样机——整车级协同验证'],
            ['<b>3 方法体系与扩展能力</b>：形成构建方法体系与集成规范，可向制动、辅助系统扩展', '方法体系与技术规范 1 套<br>判据知识库', '知识库核验、报告评审'],
            ['<b>4 示范应用</b>：在临哈线典型区段完成 FXN5C 示范，实现全工况自动回归', '示范应用与验证报告 1 份', '示范验收、第三方验证']]
    yy = table(s, 64, 214, [800, 520, 472], rows, header=['预期目标', '成果物', '考核方式'], size=19, lh=29, pad_y=17,
               col_styles={1: dict(color=INK), 2: dict(color=BODY)})
    header(s, 64, yy + 40, '知识产权与标准', w=1792)
    ip = [('4', '项', '发明专利申请', '任务编排调度、保真度与算力调度、实时时序同步、工况自生成'), ('2', '篇', '核心论文', '多智能体协同仿真与性能验证方法'),
          ('2', '项', '企业标准草案', '列车级数字样机构建、智能体接口')]
    kw_ = (1792 - 2 * 24) / 3.0
    for i, (v, u, lab_, sub) in enumerate(ip):
        kpi(s, 64 + i * (kw_ + 24), yy + 84, kw_, 200, '%s<s=20><n><c=%s> %s</c></n></s>' % (v, SUB, u), lab_, sub,
            color=KPI_COLOR, value_size=46, label_size=20, sub_size=17, pad=22)
    notes(s, '备注（讲稿）：子任务分工——子任务1 知识库与工作流（目标 1、3）；子任务2 调度引擎（目标 1）；子任务3 智能体执行（目标 2）；'
             '子任务4 集成与示范（目标 2、3、4）。各子任务交付物逐级成为下一任务输入，验证数据回流更新判据与工作流。')


# =====================================================================  16 量化指标与效益评估
def s16_metrics(d):
    s = d.new_slide()
    frame(d, s, '量化指标与效益评估：与现状对比', '指标：现状 → 目标 ｜ 效益评估：成果交付 ＋ 示范验证 ＋ 与现状对比', '预期成果与效益')
    L, LW = 64, 872
    header(s, L, 212, '指标：与现状对比', note='目标值由课题组提供', w=LW)
    ph = '<bc=%s>【待填】</bc>' % PH.text
    rows = [['一次软件迭代的验证周期', '约 2 个月', '缩短 ≥ ' + ph + '%'], ['现场验证人员投入', '10 人', '减少 ≥ ' + ph + '%'],
            ['工况覆盖度', '受台架与现场条件限制', '提升 ≥ ' + ph + '%'], ['接入异构仿真智能体', '—', '≥ ' + ph + ' 类'],
            ['协同验证典型工况', '—', '≥ ' + ph + ' 个']]
    table(s, L, 256, [330, 250, 292], rows, header=['指标', '现状', '目标'], size=19, lh=30, pad_y=29,
          col_styles={0: dict(color=INK, bold=True), 2: dict(color=INK, bold=True)})
    R, RW = 976, 880
    header(s, R, 212, '效益评估', note='成果交付 ＋ 示范验证 ＋ 与现状对比', w=RW)
    bens = [('技术效益', 'bulb', M1, '方法体系与规范交付；专利 4 项、论文 2 篇、标准草案 2 项', None),
            ('经济效益', 'coins', M3, '减少实物试验与样车试制投入，缩短研制迭代周期', [('减少实物试验', '≥【待填】次/年'), ('节约费用', '≥【待填】万元/年')]),
            ('管理与推广效益', 'world', M2, '验证流程标准化，经验沉淀为判据库；可向制动、辅助系统扩展，在**内推广', None)]
    y = 256
    for t, ic, th, txt, tg in bens:
        bh = 180 if tg else 138
        rect(s, R, y, RW, bh, fill=WHITE, line=LINE)
        rect(s, R, y, BAR, bh, fill=th.main)
        icon(s, ic, th.main, R + 26, y + 22, 32)
        text(s, R + 70, y + 18, 300, 40, t, size=22, bold=True, color=INK, anchor='m', lh=30, wrap=False)
        text(s, R + 26, y + 66, RW - 52, 60, txt, size=19, color=BODY, lh=29)
        if tg:
            xx = R + 26
            for lb, v in tg:
                xx += target_chip(s, xx, y + 126, lb, v) + 10
        y += bh + 14


# =====================================================================  17 经费预算
def s17_budget(d):
    s = d.new_slide()
    frame(d, s, '经费预算：总预算 100 万元，算力是核心投入', '预算明细与测算依据 · 分年度 · 合理性说明', '经费预算合理性')
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
    bullets(s, R, 424, RW, 400, [
        '<b>算力与任务强相关</b>：NPU 做智能体推理与编排决策，GPU 做多物理域并行仿真，满足实时协同与保真度切换',
        '<b>人员占比 30%</b>：以算法与软件开发为主，属智力密集型研究',
        '<r>无对外技术合作费</r>，全部用于自主研发；分年度投入均衡，与四阶段进度匹配',
    ], M1.main, size=21, lh=33, sa=26)
    notes(s, '备注（讲稿）：折旧——利用现有仿真与测试设备，避免重复购置；外协——数据采集与第三方验证；成果——专利、论文与标准草案；'
             '其他——示范应用差旅与会议。人员人工费用测算依据中【待填】为占位，请按“投入人月 × 人月费用标准”填写，合计应为 30 万元。')


# =====================================================================  18 项目管理计划
def s18_management(d):
    s = d.new_slide()
    frame(d, s, '项目管理计划：按四层分组承担，按里程碑绑定考核', '任务分配 · 项目流程 · 时间安排 · 执行保障', '项目管理计划')
    # 任务分配：组织架构
    header(s, 64, 212, '组织架构与任务分配', w=900)
    bx, bw = 960 - 170, 340
    rect(s, bx, 214, bw, 58, fill=INK)
    icon(s, 'user-star', WHITE, bx + 70, 214 + 14, 30)
    text(s, bx + 110, 214, bw - 130, 58, '项目负责人', size=23, bold=True, color=WHITE, anchor='m', lh=30, wrap=False)
    groups = [
        (1, '判据知识化、工作流建模', '判据知识库、工作流定义'),
        (2, '任务解析、环境装配、保真度与算力调度', '协同运行与调度引擎'),
        (3, '智能体研制、用例生成、识别与报告', '仿真智能体、验证报告'),
        (4, '架构、沙箱、样机集成、FXN5C 示范', '数字样机、方法体系与规范'),
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
        concl(s, x + 22, gy + gh - 18 - 46, gw - 44, 46, '交付：' + deli, th, size=17, pad=12)
    y1 = gy + gh + 28
    # 项目流程：管理节奏（阶梯）
    L, LW = 64, 2 * gw + 24
    header(s, L, y1, '项目管理流程', note='双周 → 月度 → 季度 → 年度', w=LW)
    cad = [('双周', '双周例会', '子任务进展同步与问题协调'), ('月度', '月度节点检查', '对照里程碑核查交付物'),
           ('季度', '季度评审', '技术方案评审与阶段成果确认'), ('年度', '年度总结', '目标达成评估与下年度计划调整')]
    cw = (LW - 3 * 24) / 4.0
    for i, (f, n, t) in enumerate(cad):
        x = L + i * (cw + 24)
        hh = 226 + i * 40
        top = 990 - hh
        rect(s, x, top, cw, hh, fill=M1.bg if i == 3 else PANEL, line=LINE)
        rect(s, x, top, cw, STYLE['concl_bar'], fill=M1.main)
        oval(s, x + 22, top + 24, 60, 60, fill=M1.soft)
        text(s, x + 22, top + 24, 60, 60, f, size=18, bold=True, color=M1.text, align='c', anchor='m', lh=24,
             wrap=False)
        text(s, x + 22, top + 98, cw - 44, 34, n, size=22, bold=True, color=INK, lh=30, wrap=False)
        text(s, x + 22, top + 138, cw - 40, 80, t, size=19, color=BODY, lh=28)
    # 时间安排：按实际月数等比例的阶段条 + 里程碑
    R, RW = 64 + 2 * (gw + 24), 2 * gw + 24
    header(s, R, y1, '时间安排与执行保障', note='里程碑绑定阶段交付物，按节点考核', w=RW)
    mw = RW / 24.0
    sy, sh_ = y1 + 62, 58
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
    ctl = [('质量', 'shield-check', '内部评审 ＋ 第三方验证，交付物可核查'),
           ('风险 → 应对', 'alert-triangle', '判据不完备 → 首批聚焦三项性能、持续回流｜实时同步失真 → 分层步长、关键工况全保真校核｜'
                                       '极端工况数据少 → 试验与运用数据标定、第三方验证'),
           ('资源', 'server', '算力平台、HIL 台架、FXN5C 试验与运用数据')]
    y = sy + sh_ + 84
    hs = [52, 112, 52]
    for (t, ic, body), h in zip(ctl, hs):
        line(s, R, y, R + RW, y, LINE, 1)
        icon(s, ic, RED if t.startswith('风险') else INK, R, y + 12, 28)
        text(s, R + 40, y + 8, 200, 36, t, size=20, bold=True, color=INK, lh=28, wrap=False)
        text(s, R + 230, y + 10, RW - 230, h - 12, body, size=18, color=BODY, lh=27)
        y += h + 6
    line(s, R, y - 6, R + RW, y - 6, LINE, 1)
    notes(s, '备注（讲稿）：双周例会——子任务进展同步与问题协调；月度——对照里程碑核查交付物；季度——技术方案与阶段成果评审；年度——目标评估与计划调整。'
             '接口与进度风险——接口规范在阶段一先行交付，里程碑绑定考核。')


# =====================================================================  19 团队研究能力
def s19_team(d):
    s = d.new_slide()
    frame(d, s, '团队研究能力：为什么由我们承担', '牵头单位：********** · 研发平台 · 仿真能力 · 试验能力 · 数据积累 · 在研项目 · 人才队伍', '团队研究能力')
    rect(s, 64, 214, 1792, 96, fill=RED_BG)
    rect(s, 64, 214, STYLE['concl_bar'], 96, fill=RED)
    text(s, 92, 214, 1740, 96, '控制系统与车辆动力学是主营方向；FXN5C 的研制、试验与运用数据在手；有 HIL 台架和现场问题记录——'
         '<b>判据来源、标定数据、验证对象都是现成的</b>。<c=%s>【按实际情况调整】</c>' % PH.text, size=21, color=INK, anchor='m', lh=32)
    items = [
        ('研发平台', 'building-factory-2', '【待填：机车总体设计／控制系统开发等平台】'),
        ('仿真能力', 'device-desktop-analytics', '【待填：已有的车辆动力学、牵引电传动、控制等仿真模型与软件】'),
        ('试验能力', 'test-pipe', 'HIL 试验台，结合经验数据模拟运用工况；【待填：其他试验台架】'),
        ('数据积累', 'database', '【待填：FXN5C 型式试验、线路试验数据与运用记录】'),
        ('在研项目', 'clipboard-list', '【待填：**级****项目数量与名称】'),
        ('人才队伍', 'users-group', '【待填：参研人员规模与职称结构】'),
    ]
    cw, ch = (1792 - 2 * 24) / 3.0, 313
    for i, (t, ic, ph) in enumerate(items):
        x = 64 + (i % 3) * (cw + 24)
        y = 340 + (i // 3) * (ch + 24)
        rect(s, x, y, cw, ch, fill=WHITE, line=LINE)
        icon_badge(s, ic, M1, x + 26, y + 24, 64)
        text(s, x + 108, y + 24, cw - 130, 64, t, size=26, bold=True, color=INK, anchor='m', lh=34, wrap=False)
        rect(s, x + 26, y + 108, cw - 52, ch - 132, fill=PH.bg, line=PH.main, lw=1.2, dash='dash')
        text(s, x + 46, y + 108, cw - 92, ch - 132, ph, size=19, color=PH.text, anchor='m', lh=28)
    notes(s, '备注（讲稿）：本页需填入单位实际基础，它是“为什么由我们承担”的直接依据。虚线框内【待填】为占位内容，请替换为牵头单位实际情况。')


# =====================================================================  20 成果介绍
def s20_results(d):
    s = d.new_slide()
    frame(d, s, '成果介绍：五项成果 ＋ 知识产权与标准', '调度引擎 · 仿真智能体／智能孪生模型 · 列车级数字样机 · 方法体系与技术规范 · 示范应用与验证报告', '成果介绍')
    top = [
        ('成果1', '软件1套', '多智能体协同运行与调度引擎', M1, 'settings-automation', [
            '输入验证需求，自动完成<b>分解、装配、调度、监控与报告</b>',
            '支持异构仿真智能体统一接入'],
         ('调度引擎软件', '建议：软件界面截图或系统架构图')),
        ('成果2', '1套', '关键系统仿真智能体／<br>智能孪生模型', M2, 'robot', [
            '控制系统、车辆动力学、线路等关键系统',
            '可按实测数据<b>自动标定</b>，标准接口可扩展'],
         ('仿真智能体／孪生模型', '建议：模型结构图或仿真运行界面截图')),
        ('成果3', '1套', '列车级数字样机', M3, 'train', [
            '成果 2 经成果 1 集成，以 <b>FXN5C</b> 为对象的整车级验证环境',
            '用于回归验证与问题复现（智能运维），预留制动、辅助系统接口'],
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
              '构建方法、集成流程与接口规范，可在**内推广、后续课题复用'),
             ('成果5', '1份', '示范应用与验证报告', M5, 'report-analytics', 'FXN5C 临哈线典型区段关键性能协同验证与效能评估')]
    sh = (990 - y1 - 14) / 2.0
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
    ip = [('4', '项', '发明专利申请', '任务编排调度、保真度与算力调度、实时时序同步、工况自生成'),
          ('2', '篇', '核心论文', '多智能体协同仿真与性能验证方法'),
          ('2', '项', '企业标准草案', '列车级数字样机构建、智能体接口')]
    kw_ = (RW - 2 * 12) / 3.0
    for i, (v, u, lab_, sub) in enumerate(ip):
        x = R + i * (kw_ + 12)
        kpi(s, x, y1 + 40, kw_, 990 - y1 - 40, '%s<s=20><n><c=%s> %s</c></n></s>' % (v, SUB, u), lab_, sub, color=KPI_COLOR,
            value_size=46, label_size=20, sub_size=16, pad=18)
    notes(s, '备注（讲稿）：成果 1—3 中的虚线框为附图预留位置（形状组合）；选中该组合删除后插入实际图片即可。')


# =====================================================================  21 结束页
def s21_end(d):
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


SLIDES = [s01_cover, s02_agenda, s03_intro, s04_background, s05_value, s06_now, s07_route, s08_layer1, s09_layer2,
          s10_layer3, s11_layer4, s12_scenario, s13_innovation, s14_stages, s15_results_assess, s16_metrics,
          s17_budget, s18_management, s19_team, s20_results, s21_end]


def build(out=OUT):
    d = Deck(TOTAL, FOOTER, report_type='课题申报汇报')
    for f in SLIDES:
        f(d)
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
