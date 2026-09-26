# -*- coding: utf-8 -*-
"""根据《项目汇报PPT文本.md》（v7）生成课题申报汇报 PPT。

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
    frame(d, s, '汇报提纲：七个部分，核心是研究内容、技术路线与方法', '七个部分 · 第三部分为核心章节（第 7—14 页）', '汇报提纲')
    rows = [
        ('01', '项目简介', '基本信息 · 子课题与成果物 · 预计经费', '第 3 页', 'file-description'),
        ('02', '研究意义与价值', '背景与目的 · 应用价值与效益 · 为什么是现在', '第 4—6 页', 'bulb'),
        ('03', '主要研究内容、技术路线与方法', '总体路线 · 研究内容①—④ · 典型场景 · 创新点 · 阶段与检验', '第 7—14 页', 'route'),
        ('04', '预期成果与效益', '成果与考核 · 量化指标与效益', '第 15—16 页', 'target-arrow'),
        ('05', '经费预算合理性', '预算明细 · 分年度 · 合理性', '第 17 页', 'coins'),
        ('06', '项目管理计划', '任务 · 流程 · 时间 · 保障 ｜ 团队能力', '第 18—19 页', 'sitemap'),
        ('07', '成果介绍', '五项成果 · 知识产权与标准', '第 20 页', 'trophy'),
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
    frame(d, s, '项目简介：研究重心是引擎，数字样机是引擎的集成产物', '项目基本信息 · 子课题与本课题边界 · 预计成果物 · 预计经费',
          '项目简介')
    # ---- 左：基本信息（项目、课题名称见封面，这里只列要素）
    L, LW = 64, 840
    header(s, L, 212, '项目基本信息', w=LW)
    lab = dict(bold=True, color=SUB, size=18)
    val = dict(color=INK, size=20)
    yy = table(s, L, 256, [108, 296, 124, 312], [
        ['牵头单位', '**********', '项目负责人', '***'],
        ['主管部门', '********部（***）', '产业领域', '轨道交通装备'],
        ['项目目的', '基础前瞻共性技术研究', '研究周期', '2027.01—2028.12'],
    ], size=20, lh=28, pad_x=10, pad_y=13, col_styles={0: lab, 1: val, 2: lab, 3: val})
    # ---- 左：预计成果物（清单：名称 + 形式，不用卡片和图标）
    y1 = yy + 36
    header(s, L, y1, '预计成果物', w=LW)
    items = [('多智能体协同运行与调度引擎', '软件 1 套'), ('关键系统仿真智能体／智能孪生模型', '模型 1 套'),
             ('列车级数字样机', '系统 1 套'), ('列车级数字样机构建方法体系与技术规范', '规范 1 套'),
             ('示范应用与验证报告', '报告 1 份'), ('知识产权与标准', '专利 4 · 论文 2 · 标准 2')]
    rh = 54
    ly = y1 + 46
    line(s, L, ly, L + LW, ly, LINE, 1)
    for r, (name, form) in enumerate(items):
        ry = ly + r * rh
        text(s, L + 10, ry, LW - 250, rh, name, size=20, color=INK, anchor='m', lh=28, wrap=False)
        text(s, L + LW - 240, ry, 230, rh, form, size=18, color=SUB, align='r', anchor='m', lh=26, wrap=False)
        line(s, L, ry + rh, L + LW, ry + rh, LINE, 1)
    # ---- 左：预计经费（细条，条内不写字，数字写在下方）
    y2 = ly + len(items) * rh + 36
    header(s, L, y2, '预计经费', note='总预算 100 万元', w=LW)
    by = y2 + 50
    stacked_bar(s, L, by, LW, 22, [(v, c, None, WHITE) for _, _, v, _, c, _ in BUDGET])
    text(s, L, by + 30, LW, 28, '<b>直接投入 50</b>　算力服务器 40 · 工作站 10　　<b>人员 30</b>　　折旧／外协／成果／其他 各 5',
         size=17, color=SUB, lh=26, wrap=False)
    BOT = by + 58                                 # 左栏底线，右栏与之对齐

    # ---- 右：子课题与本课题边界（清单式，只有“本课题”一行着色）
    R, RW = 944, 912
    header(s, R, 212, '项目子课题与本课题边界', w=RW)
    rows = [
        ('7.1', '高质量数据集', '用其数据规范与接口，不重复建库', 'n'),
        ('7.2', '多物理域建模与仿真智能体', '只建验证必需的智能体', 'n'),
        ('7.3', '协同运行与调度引擎',
         '核心攻关：多专业智能体“接得上、跑得动、调得灵”<br>4 个子任务：知识 → 调度 → 执行 → 集成', 'me'),
        ('7.4', '数字样机构建', '承接智能体集成与协同运行', 'n'),
        ('7.5', '方案优化与智能生成', '为优化提供验证支撑', 'n'),
        ('6.2', '通用智能体调度平台', '本课题专注面向性能验证的领域化编排', 'ext'),
    ]
    nb, tgap = 84, 20
    tw_ = RW - nb - tgap - 24
    ext_gap = 24
    groups = [34 + 6 + wrap_lines(desc, tw_, 19) * 29 for _, _, desc, _ in rows]
    pad = (BOT - 256 - ext_gap - sum(groups)) / float(len(rows))
    y = 256
    line(s, R, y, R + RW, y, LINE, 1)
    for (no, name, desc, kind), g in zip(rows, groups):
        h = g + pad
        ext = (kind == 'ext')
        me = (kind == 'me')
        if ext:
            y += ext_gap
            line(s, R, y, R + RW, y, RULE, 1.5, dash='dash')
        shp = None
        if me:
            rect(s, R, y, RW, h, fill=RED_BG)
            shp = rect(s, R, y, nb, h, fill=RED)
        text(s, R, y, nb, h, no, size=26, bold=True, color=WHITE if me else INK, font=MONO, align='c', anchor='m',
             lh=34, shp=shp, wrap=False)
        tx = R + nb + tgap
        top = y + (h - g) / 2.0
        text(s, tx, top, 640, 34, name, size=22, bold=True, color=RED if me else INK, lh=34, wrap=False)
        if me:
            tag(s, tx + text_width(name, 22, True) + 14, top + 3, '本课题', fill=RED, size=16, h=28, padx=10)
        if ext:
            tag(s, tx + text_width(name, 22, True) + 14, top + 3, '相邻项目', fill=GRAYBG, color=SUB, size=16, h=28,
                padx=10)
        text(s, tx, top + 40, tw_, g - 38, desc, size=19, color=BODY if me else SUB, lh=29)
        y += h
        line(s, R, y, R + RW, y, RULE if ext else LINE, 1.5 if ext else 1, dash='dash' if ext else None)
    notes(s, '备注（讲稿）：项目名称——7. 基于智能体的轨道交通装备性能数字样机关键技术研究；本课题——7.3 轨道交通装备性能验证的'
             '多智能体协同运行与调度引擎技术研究。总预算 100 万元，当年（2027 年）预算 50 万元；直接投入 50 万元为 GPU+NPU 异构算力服务器 40 万、'
             '仿真与调度一体化工作站 10 万，人员 30 万，折旧、外协、成果、其他各 5 万。\n'
             '边界细节——7.1 采用其数据规范与接口，不承担建库（7.1 建库、7.3 按需生成）；7.2 仅自建控制系统、车辆动力学与线路等验证必需的智能体，'
             '不覆盖弓网/牵引/制动等全部关键系统；7.3 核心攻关引擎，数字样机是引擎承载智能体后的集成产物；7.4 承接智能体集成与协同运行；'
             '7.5 为优化提供验证支撑；6.2 通用智能体调度平台是通用编排，本课题强调实时协同与多物理域时序。\n'
             '数字样机由关键系统仿真智能体／智能孪生模型经调度引擎集成而成：智能体可单独核验、复用，数字样机作为整车级验证系统交付。')


# =====================================================================  04 研究背景与重要性
def s04_background(d):
    s = d.new_slide()
    frame(d, s, '研究背景与重要性：一次验证 40 人、2 个月，问题仍到现场才暴露', '验证现状 · 真实案例 · 症结 · 研究目的', '研究意义与价值')
    L, LW = 64, 872
    big_numbers(s, L, 214, LW, [('40', '人', '研发＋测试＋现场'), ('<s=18><n>约</n></s>2', '个月', '一次迭代'),
                                ('−10', '℃', '现场案例')], h=96, label_w=140)
    header(s, L, 362, '即便如此，问题还是在现场才暴露', w=LW)
    qy, qh = 406, 278
    rect(s, L, qy, LW, qh, fill=RED_BG)
    rect(s, L, qy, STYLE['concl_bar'], qh, fill=RED)
    icon(s, 'quote', MUTED, L + 24, qy + 26, 32)
    text(s, L + 72, qy, LW - 104, qh,
         '2024 年 12 月，FXN5C 在临哈线长大下坡，约 −10℃：回手柄后风扇仍以 40Hz 运行，低温水 25℃ → 0℃、高温水 78℃，'
         '<b>散热器冻结</b>。<br>台架没有复现出来。', size=23, color=INK, anchor='m', lh=38)
    R, RW = 976, 880
    header(s, R, 214, '症结', w=RW)
    cards = [('复杂工况模拟不了', '多系统耦合、长时序，台架覆盖不到'),
             ('验证靠人，跟不上迭代', '工况人工编、环境人工搭、结果人工看'),
             ('经验散在人手里', '判据散在标准、大纲和专家经验里')]
    for i, (t, b) in enumerate(cards):
        problem_card(s, R, 258 + i * 148, RW, 130, t, b, title_size=26, body_size=20, pad=28)
    summary(s, 64, 744, 1792, 150, '研究目的：让控制软件每改一版，复杂运用工况都能在数字样机上自动跑一遍、自动判读',
            '问题在研发阶段暴露，现场只做确认。', title_size=28, desc_size=22, gap=8)
    notes(s, '备注（讲稿）：验证现状——控制软件每次迭代，HIL 台架＋经验数据模拟运用工况，重点工况再到现场验证；一次投入 40 人'
             '（研发 20 ＋ 测试 10 ＋ 现场 10），周期约 2 个月。\n'
             '冻结案例：控制、冷却、动力、线路四个系统耦合，且是回手柄后持续一段时间才出现的长时序过程，HIL 只有控制器在环、其他系统靠经验数据近似，'
             '因此复现不出来。\n技术表述——研制多智能体协同运行与调度引擎，以 FXN5C 为对象建成列车级数字样机，实现列车关键性能的快速、可复现、全覆盖验证。')


# =====================================================================  05 应用价值与社会经济效益
def s05_value(d):
    s = d.new_slide()
    frame(d, s, '应用价值：一句话需求进去，全工况报告出来', '用在三个环节 · 对领域发展的推动 · 社会经济效益', '研究意义与价值')
    header(s, 64, 212, '用在三个环节', w=1792)
    stages = [('研发', '软件迭代', '台架＋现场验证', '全工况在数字样机上自动回归'),
              ('试验', '上线前预试验', '低温等工况受季节、线路限制', '先在数字样机上预演，缩减上线项目'),
              ('运用', '问题复现', '靠经验排查', '用运行记录复现问题、定位原因')]
    cw = (1792 - 2 * 40) / 3.0
    y0, ch = 256, 266
    for i, (a, b, now, aft) in enumerate(stages):
        x = 64 + i * (cw + 40)
        rect(s, x, y0, cw, ch, fill=WHITE, line=LINE)
        rect(s, x, y0, cw, BAR, fill=M1.main)
        text(s, x + 28, y0 + 24, cw - 56, 44, '%s<c=%s> · </c>%s' % (a, MUTED, b), size=28, bold=True, color=INK,
             anchor='m', lh=38, wrap=False)
        text(s, x + 28, y0 + 90, cw - 56, 26, '现在', size=17, bold=True, color=MUTED, lh=24, wrap=False)
        text(s, x + 28, y0 + 118, cw - 56, 32, now, size=20, color=SUB, lh=30, wrap=False)
        line(s, x + 28, y0 + 168, x + cw - 28, y0 + 168, LINE, 1)
        text(s, x + 28, y0 + 184, cw - 56, 26, '应用引擎后', size=17, bold=True, color=INK, lh=24, wrap=False)
        text(s, x + 28, y0 + 212, cw - 56, 66, aft, size=22, bold=True, color=INK, lh=32)
        if i < 2:
            tri(s, x + cw + 20, y0 + ch / 2.0, 16, 22, TRI)
    header(s, 64, 566, '社会经济效益', w=1792)
    effs = [('经济效益', '减少现场验证与线路试验投入，缩短迭代周期'),
            ('社会效益', '复杂工况覆盖更全、结论更可信，减少运用故障'),
            ('产业效益', '方法可扩展到制动、辅助系统，在**内主机企业推广')]
    ew = (1792 - 2 * 40) / 3.0
    for i, (t, b) in enumerate(effs):
        x = 64 + i * (ew + 40)
        rect(s, x, 610, ew, 112, fill=WHITE, line=LINE)
        rect(s, x, 610, BAR, 112, fill=M3.main)
        text(s, x + 28, 624, ew - 56, 32, t, size=21, bold=True, color=INK, lh=30, wrap=False)
        text(s, x + 28, 662, ew - 56, 52, b, size=19, color=BODY, lh=28)
    summary(s, 64, 766, 1792, 120, '推动领域发展：性能验证从“单系统、离线、靠现场”，走向“列车级、实时协同、研发阶段可回归”',
            None, title_size=26)
    notes(s, '备注（讲稿）：用起来的样子——输入一句话需求，引擎自动生成工况、装配智能体、分配精度与算力、协同求解、对标判据，'
             '输出验证报告与异常定位，全程可回放，工程师只看异常项（完整流程见第 12 页典型应用场景）。\n'
             '对领域发展的推动作用——性能验证成为研发体系中可持续使用的基础能力；输出的构建方法体系与技术规范，主机企业可按同一流程构建自己的数字样机；'
             '新增系统按标准接口接入智能体即可，引擎不需重构。')


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
    header(s, L, 212, '四层架构与分工', w=LW)
    roles = {1: '标准、大纲、案例 → 可自动比对的判据', 2: '需求 → 任务；自动装配智能体、分配精度与算力',
             3: '跑工况、自动判读、生成报告', 4: '智能体在沙箱中协同运行，集成为数字样机'}
    y, bh, gap = 256, 104, 12
    for i in (4, 3, 2, 1):
        th = LAYER[i]
        rect(s, L, y, LW, bh, fill=WHITE, line=LINE)
        rect(s, L, y, 250, bh, fill=th.bg)
        rect(s, L, y, BAR, bh, fill=th.main)
        text(s, L + 26, y + 18, 214, 34, LAYER_NO[i] + ' ' + LAYER_KIND[i], size=24, bold=True, color=th.text, lh=34,
             wrap=False)
        text(s, L + 26, y + 56, 218, 30, LAYER_NAME[i], size=17, bold=True, color=INK, lh=26, wrap=False)
        text(s, L + 276, y, LW - 300, bh, roles[i], size=21, color=BODY, anchor='m', lh=32)
        y += bh + gap
    bot = y - gap                                 # 四层底边，右侧表格与之对齐
    R, RW = 1056, 800
    header(s, R, 212, '接入现有体系，不另起炉灶', w=RW)
    rows = [['设计规范、试验大纲', '提取为判据与流程模板'], ['已有仿真模型、HIL 台架', '标准接口接入，不重复建模'],
            ['试验与运用数据', '用于智能体自动标定'], ['验证报告', '作为设计评审、软件放行的依据']]
    table(s, R, 256, [340, 460], rows, header=['现有资源', '接入方式'], size=20, lh=30,
          pad_y=((bot - 256 - 44) / 4.0 - 30) / 2.0, col_styles={0: dict(color=INK, bold=True)})
    summary(s, 64, 770, 1792, 130, '怎么算成功：数字样机复现临哈线散热器冻结；一次软件迭代完成全工况回归',
            '验证周期由约 2 个月缩短至<bc=%s>【待填】</bc>' % PH.text, title_size=26, desc_size=21, gap=8)
    notes(s, '备注（讲稿）：闭环——判据与流程 → 分解与调度 → 协同执行与判读 → 报告 → 数据回归、更新判据；全程留痕，可回放、可回归。'
             '各阶段检验见第 14 页。\n'
             '与课题要求的对应——高效实时协同运行 → ③④；调度引擎 → ②；列车级数字样机及构建方法体系 → ④；支持关键性能验证与智能运维 → ④。'
             '对接项目7总体目标——模型自动标定 → ③；指标实时推演 → ④；性能自动评估 → ③。')


# =====================================================================  08 研究内容① 基础层
def s08_layer1(d):
    s = d.new_slide()
    th = LAYER[1]
    frame(d, s, '研究内容① 测试知识库与任务工作流：把判据和流程从人手里拿出来', '现状与问题 · 研究内容 · 研究方法 · 示意图', '研究内容① 基础层')
    content_left(s, 1, outs=['验证判据知识库', '任务工作流定义<br>与配置规范'], outs_bottom=990)
    X, XW = 356, 1500
    scene_band(s, X, 214, XW, 106,
               '判一个工况合不合格，依据散在国标、行标、企标、试验大纲和专家经验里，查找比对全靠人工。',
               '<b>知识工程</b>：规范形式化、领域知识建模、工作流建模。难点：判据多藏在经验和文档里。',
               scene_label='现状与问题', role_label='研究方法')
    cw = (XW - 48) / 3.0
    cards = [
        ('验证判据知识化', '从规范、大纲与问题记录中提取判据，变成<b>可自动比对</b>的条目'),
        ('任务工作流建模', '定义分解、顺序、流转与异常分支，形成<b>可配置</b>的流程模板'),
        ('知识与流程迭代', '验证结果回流，持续补充判据、修订流程'),
    ]
    for i, (t, desc) in enumerate(cards):
        point_card(s, X + i * (cw + 24), 330, cw, 150, '0%d' % (i + 1), t, desc, th, desc_size=20)
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
    notes(s, '备注（讲稿）：知识库与工作流是②“任务理解与编排调度”的知识与规则来源；判据首批覆盖速度跟踪、牵引能耗、运行平稳性。'
             '没有判据知识库，任务理解只能靠通用大模型“猜”；没有工作流，编排调度无章可循。判据条目示例：验证对象——牵引控制系统；'
             '关键性能——速度跟踪偏差；适用工况——长大下坡、湿轨；判定规则——|实测 − 目标| ≤ 允许偏差；来源——设计规范条款。\n' + FIG_NOTE)


# =====================================================================  09 研究内容② 决策层
def s09_layer2(d):
    s = d.new_slide()
    th = LAYER[2]
    frame(d, s, '研究内容② 任务理解与智能编排调度：引擎在此落地', '现状与问题 · 研究内容 · 研究方法 · 示意图',
          '研究内容② 决策层')
    content_left(s, 2, outs=['多智能体协同运行与调度引擎（软件1套）'], outs_bottom=990)
    X, XW = 356, 1500
    scene_band(s, X, 214, XW, 106,
               '区段、温度、轨面、载重、手柄序列组合起来：全用高精度模型算不完，全用简化模型又不可信。',
               '<b>智能调度</b>：任务分解与匹配、多目标寻优、保真度—算力协同配置。难点：在线权衡精度与效率。',
               scene_label='现状与问题', role_label='研究方法')
    cw = (XW - 60) / 4.0
    cards = [
        ('目标解析与任务分解', '一句话需求 → 验证对象、关键性能、判据条目'),
        ('测试环境自动装配', '自动组合智能体、连接接口、配置算力'),
        ('保真度与算力自适应调度', '常规工况降阶提速，关键工况<b>强制全保真</b>'),
        ('执行监控与失败重规划', '异常中断后自动重规划，支撑指标实时推演'),
    ]
    for i, (t, desc) in enumerate(cards):
        point_card(s, X + i * (cw + 20), 330, cw, 150, '0%d' % (i + 1), t, desc, th, desc_size=19)
    # 附图：一次验证任务中的调度过程
    f, py = fig_area(s, X, 496, XW, 494, '一次验证中的调度过程：按工况切换保真度、分配算力')
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
    notes(s, '备注（讲稿）：上下游——向下承接①的知识与规则，向③下发执行方案（智能体组合、工况、算力）。目标解析时匹配知识库中的判据条目；'
             '保真度调度中关键与极端工况强制全保真（极端工况不能外推），GPU 跑多物理域并行仿真，NPU 做推理与编排决策。'
             '算力为本课题配置的 GPU＋NPU 异构服务器 1 台、一体化工作站 2 台。\n' + FIG_NOTE)


# =====================================================================  10 研究内容③ 执行层
def s10_layer3(d):
    s = d.new_slide()
    th = LAYER[3]
    frame(d, s, '研究内容③ 智能体执行：把 −10℃ 现场的问题在研发室里跑出来', '现状与问题 · 研究内容 · 研究方法 · 示意图', '研究内容③ 执行层')
    content_left(s, 3, outs=['关键系统仿真智能体／<br>智能孪生模型', '用例与线路数据集', '验证报告与追溯记录'], outs_bottom=990)
    X, XW = 356, 1500
    scene_band(s, X, 214, XW, 106,
               '仿真结果靠人看曲线，多系统、长时序问题（如临哈线散热器冻结）难发现、难定位。',
               '<b>实时协同与可信评判</b>：分布式协同仿真、时序同步、自动标定。难点：步长差异大、极端工况数据少。',
               scene_label='现状与问题', role_label='研究方法')
    gap = 20
    cw, ch, y0 = (XW - 3 * gap) / 4.0, 150, 330
    steps = [
        ('仿真智能体', '控制、车辆动力学、线路等智能体标准化接入，实测数据<b>自动标定</b>'),
        ('用例生成', '按线路资料与覆盖要求，自动生成用例与线路数据'),
        ('测试识别', '逐工况对标判据，<b>识别异常、定位原因</b>'),
        ('报告与回归', '结论自动汇总可追溯，设计变更后<b>回归重跑</b>'),
    ]
    for i, (t, desc) in enumerate(steps):
        x = X + i * (cw + gap)
        point_card(s, x, y0, cw, ch, '0%d' % (i + 1), t, desc, th, desc_size=19)
        if i < 3:
            tri(s, x + cw + gap / 2.0, y0 + ch / 2.0, 14, 20, TRI)
    # 附图：用现场案例说明测试识别（附图区与其他研究内容页同位同高，图内内容整体下移居中）
    f, py = fig_area(s, X, 496, XW, 494, '以临哈线散热器冻结为例：识别异常、定位原因', note='示意图 · 防冻限值【待填】')
    py += 52
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
    notes(s, '备注（讲稿）：对接项目7总体目标——模型自动标定 → 仿真智能体；性能自动评估 → 测试识别。标定数据来自型式试验、线路试验与运用记录；'
             '用例生成的线路数据包括平纵断面、曲线、超高、坡度、不平顺谱。测试识别同时输出结果可信性判断，解决极端工况实测数据少、结论难以自证的问题。'
             '数据回归：结果回流，迭代模型标定、判据与工况集；设计修改后，低温工况全量回归重跑。\n示意图：数据据 2024 年 12 月临哈线现场记录，曲线为示意；防冻限值线仅示意位置，请按设计规范取值；'
             '如有现场记录曲线，选中附图组合删除后替换即可。')


# =====================================================================  11 研究内容④ 集成层
def s11_layer4(d):
    s = d.new_slide()
    th = LAYER[4]
    frame(d, s, '研究内容④ 架构融合与协同运行：融合成一个能跑起来的整体', '现状与问题 · 研究内容 · 研究方法 · 示意图', '研究内容④ 集成层')
    content_left(s, 4, outs=['列车级数字样机', '列车级数字样机构建方法体系与集成技术规范'], outs_bottom=990)
    X, XW = 356, 1500
    scene_band(s, X, 214, XW, 106,
               '台架以控制器为核心，其他系统靠经验数据近似，多系统难以同台协同，复杂场景无法模拟。',
               '<b>系统集成</b>：标准化接口、服务化封装、沙箱化运行与全程留痕。难点：异构模型实时协同的稳定与可复现。',
               scene_label='现状与问题', role_label='研究方法')
    cw = (XW - 60) / 4.0
    cards = [
        ('总体架构与接口规范', '定义四层接口与数据流；已有模型、HIL 台架<b>标准接入</b>'),
        ('多智能体协同运行', '异构智能体<b>实时协同求解</b>，时序与状态同步'),
        ('可复现沙箱', '<b>隔离、记录、回放</b>：任意一次验证都能复现'),
        ('列车级数字样机', '以 <b>FXN5C</b> 为对象集成关键系统智能体'),
    ]
    for i, (t, desc) in enumerate(cards):
        point_card(s, X + i * (cw + 20), 330, cw, 150, '0%d' % (i + 1), t, desc, th, desc_size=19)
    # 附图：数字样机构成
    f, py = fig_area(s, X, 496, XW, 494, '列车级数字样机构成')
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
    notes(s, '备注（讲稿）：对接项目7总体目标——指标实时推演 → 协同运行。总体架构定义四层的接口、数据流与控制流，已有模型与 HIL 台架经标准接口接入，'
             '不重复建设；可复现沙箱——隔离（单路发散不扩散）、记录（参数、模型版本、随机种子）、回放（任意一次验证可复现）；'
             '数字样机以 FXN5C 为示范对象，支持性能验证与智能运维。\n' + FIG_NOTE)


# =====================================================================  12 典型应用场景
def s12_scenario(d):
    s = d.new_slide()
    frame(d, s, '典型应用场景：FXN5C 临哈线控速＋能耗验证，全流程交给智能体', '输入一句话需求 → 引擎自动完成五步 → 验证报告 ＋ 全要素留痕', '典型应用场景')
    y0 = 214
    shp = rect(s, 64, y0, 250, 84, fill=ARCH)
    text(s, 64, y0, 250, 84, '输入 · 一句话需求', size=22, bold=True, color=WHITE, align='c', anchor='m', lh=30, shp=shp, wrap=False)
    rect(s, 314, y0, 1542, 84, fill=PANEL)
    icon(s, 'quote', MUTED, 340, y0 + 24, 36)
    text(s, 394, y0, 1440, 84, '验证 FXN5C 在临哈线上按给定运行图运行时的<r>速度跟踪</r>与<r>牵引能耗</r>是否满足要求。', size=26, bold=True,
         color=INK, anchor='m', lh=36, wrap=False)
    header(s, 64, 338, '引擎自动完成', w=1792)
    steps = [('任务理解', '对象：牵引控制系统<br>性能：速度跟踪、牵引能耗'),
             ('工况生成', '临哈线数据 × 干／湿轨面 × 满／空载 × 坡道、曲线、隧道'),
             ('装配与调度', '自动装配控制、动力学、线路智能体，按工况分配精度与算力'),
             ('执行与判读', '协同求解，定位异常（如某曲线段速度超差）'),
             ('输出', '验证报告＋全程留痕，可回放、可回归')]
    gap = 24
    cw = (1792 - 4 * gap) / 5.0
    for i, (t, b) in enumerate(steps):
        x = 64 + i * (cw + gap)
        point_card(s, x, 382, cw, 168, '0%d' % (i + 1), t, b, M1, desc_size=19)
        if i < 4:
            tri(s, x + cw + gap / 2.0, 382 + 84, 14, 20, TRI)
    header(s, 64, 594, '与现有做法对比', w=1792)
    rows = [['<b>工况准备</b>', '人工编写', '按线路资料与判据自动生成'], ['<b>环境搭建</b>', '台架＋人工联调', '自动装配智能体与算力'],
            ['<b>结果判读</b>', '人工看曲线', '自动比对判据、定位原因']]
    table(s, 64, 638, [340, 560, 892], rows, header=['环节', '现状', '应用引擎后'], size=20, lh=30, pad_y=20,
          col_styles={2: dict(color=INK)})
    notes(s, '备注（讲稿）：这一次验证里，工况没人手写、线路数据没人准备、环境没人搭、异常没人看图判断。判据取自知识库；'
             '工况组合为临哈线线路数据 × 轨面（干／湿）× 载荷（满／空）× 线路条件（长大坡道／小半径曲线／隧道）；执行中自动标定。'
             '追溯原因示例：黏着利用不足？控制参数不匹配？')


# =====================================================================  13 创新点与技术定位
def s13_innovation(d):
    s = d.new_slide()
    frame(d, s, '创新点与技术定位：先进性不在单项技术，而在调度的维度与粒度', '三项突破 · 技术定位', '创新点')
    header(s, 64, 212, '三项突破', w=1792)
    rows = [['<b>① 调度对象升级为智能体</b>', '智能体自主标定、自主判读，引擎调度的是智能体而非算例', '<b>人退出执行回路</b>'],
            ['<b>② 调度多一个维度：保真度</b>', '按工况切换降阶／全保真，GPU、NPU 算力随之调配', '<b>批量提速与关键工况精度兼得</b>'],
            ['<b>③ 调度基础是可复现沙箱</b>', '参数、模型版本、随机种子全程留痕与回放', '<b>现场问题可复现，改进后可回归</b>']]
    yy = table(s, 64, 256, [440, 860, 492], rows, header=['突破', '技术机制', '带来的变化'], size=20, lh=30, pad_y=18,
               col_styles={0: dict(color=RED), 2: dict(color=INK)})
    header(s, 64, yy + 40, '技术定位', note='不替代现有工具，把它们接起来、调度起来', w=1792)
    rows = [('商用仿真软件', '单专业高精度', '列车级协同'), ('协同仿真标准', '模型封装与交换', '智能编排、实时性'),
            ('HIL 试验台', '控制器实时验证', '工况受限、不可批量'), ('通用智能体框架', '任务编排', '多物理域实时协同')]
    trs = [['<b>%s</b>' % a, '<c=%s>✓</c>  %s' % (OK.main, b), '<bc=%s>×</bc>  %s' % (RED, c)] for a, b, c in rows]
    yy2 = table(s, 64, yy + 84, [440, 660, 692], trs, header=['现有手段', '能做', '做不到'], size=20, lh=30, pad_y=12,
                col_styles={1: dict(color=OK.text), 2: dict(color=BODY)})
    summary(s, 64, yy2 + 30, 1792, 90, '在成熟技术之上，补上“<r>面向列车级性能验证的智能编排与实时协同</r>”这一层', None,
            title_size=25)
    notes(s, '备注（讲稿）：不做仿真软件、不做接口标准、不做试验台。其他现有手段——线路试验：结果权威，但周期长、不可复现；'
             '降阶模型：提速，但极端工况不能外推。\n与同类工作的差异——现有工作多聚焦单一系统仿真或通用智能体平台，'
             '本课题把协同运行与调度做成可交付的引擎和可复用的方法体系。')


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
    for yi, yr in enumerate(['2027 年', '2028 年']):
        text(s, X0 + yi * 12 * mw, 212, 12 * mw, 30, yr, size=20, bold=True, color=INK, align='c', lh=28, wrap=False)
    line(s, X0 + 12 * mw, 212, X0 + 12 * mw, 350, RULE, 1.5, dash='dash')
    by = 252
    for name, short, a, b, col, *_ in STAGES:
        x = X0 + a * mw + (2 if a else 0)
        w = (b - a) * mw - (2 if a else 0)
        shp = rect(s, x, by, w, 56, fill=col)
        fg = STAGE_ON[STAGE.index(col)] if STAGE_ON else on_color(col)
        full = '%s · %s' % (name, short)
        text(s, x, by, w, 56, full if text_width(full, 20, True) + 24 <= w else name, size=20, bold=True, color=fg,
             align='c', anchor='m', lh=28, shp=shp, wrap=False)
    ty = by + 64
    rect(s, X0, ty, X1 - X0, 3, fill=RULER)
    for m in range(25):
        xx = X0 + m * mw
        rect(s, xx - 1, ty + 3, 2, 12 if m % 12 == 0 else (8 if m % 3 == 0 else 5), fill=RULER)
    for m in range(2, 24, 3):                     # 只标季末月份
        text(s, X0 + (m + 0.5) * mw - 24, ty + 14, 48, 24, '%02d' % (m % 12 + 1), size=15, color=MUTED, font=MONO,
             align='c', lh=20, wrap=False)
    my = ty + 50
    for name, short, a, b, col, per, goal, deli, mk, mdate, mtext in STAGES:
        xx = X0 + b * mw
        shape(s, MSO_SHAPE.DIAMOND, xx - 12, my, 24, 24, fill=RED)
        al = 'r' if b == 24 else 'c'
        bx = xx - 230 if al == 'r' else xx - 115
        text(s, bx, my + 30, 230, 28, '%s · %s' % (mk, mdate), size=19, bold=True, color=RED, align=al, lh=26,
             wrap=False, font=FONT)
    rows = [['<b>%s</b>' % name, goal, CHECKS[mk]] for name, short, a, b, col, per, goal, deli, mk, mdate, mtext in STAGES]
    table(s, 64, my + 96, [220, 760, 812], rows, header=['阶段', '目标', '检验（考题）'], size=20, lh=30, pad_y=28,
          col_styles={0: dict(color=INK), 2: dict(color=INK)})
    notes(s, '备注（讲稿）：阶段起止与里程碑——阶段一 2027.01—06（M1 架构与接口规范）；阶段二 2027.07—2028.03（M2 智能体与引擎）；'
             '阶段三 2028.04—09（M3 数字样机与验证）；阶段四 2028.10—12（M4 示范应用与验收）。\n'
             '阶段交付——一：总体架构、判据知识化方法、接口规范；二：关键系统仿真智能体／智能孪生模型、调度引擎原理样机；'
             '三：列车级数字样机、典型工况协同验证结果；四：示范应用与效能评估、方法体系与技术规范、标准草案。前一阶段交付物即后一阶段输入。')


# =====================================================================  15 预期成果与考核方式
def s15_results_assess(d):
    s = d.new_slide()
    frame(d, s, '预期成果与考核方式：四项目标，每项有成果物和考核方式', '目标 → 成果物 → 考核方式 ｜ 知识产权与标准', '预期成果与效益')
    rows = [['<b>1 调度引擎</b>：异构智能体实时协同运行与算力动态调度', '引擎软件 1 套', '软件演示、功能核验'],
            ['<b>2 数字样机与性能验证</b>：建成 FXN5C 数字样机，速度跟踪、牵引能耗、运行平稳性多工况协同验证',
             '智能体／孪生模型 1 套<br>数字样机 1 套', '智能体：与实测比对<br>数字样机：整车级协同验证'],
            ['<b>3 方法体系与扩展能力</b>：构建方法体系与集成规范，可向制动、辅助系统扩展', '方法体系与技术规范 1 套<br>判据知识库',
             '知识库核验、报告评审'],
            ['<b>4 示范应用</b>：临哈线典型区段完成 FXN5C 示范，全工况自动回归', '示范应用与验证报告 1 份', '示范验收、第三方验证']]
    yy = table(s, 64, 214, [820, 500, 472], rows, header=['预期目标', '成果物', '考核方式'], size=20, lh=30, pad_y=24,
               col_styles={1: dict(color=INK), 2: dict(color=BODY)})
    header(s, 64, yy + 44, '知识产权与标准', w=1792)
    ip = [('4', '项', '发明专利申请'), ('2', '篇', '核心论文'), ('2', '项', '企业标准草案')]
    kw_ = (1792 - 2 * 24) / 3.0
    for i, (v, u, lab_) in enumerate(ip):
        kpi(s, 64 + i * (kw_ + 24), yy + 88, kw_, 132, '%s<s=20><n><c=%s> %s</c></n></s>' % (v, SUB, u), lab_, None,
            color=KPI_COLOR, value_size=46, label_size=21, pad=22)
    notes(s, '备注（讲稿）：知识产权方向——发明专利：任务编排调度、保真度与算力调度、实时时序同步、工况自生成；论文：多智能体协同仿真与性能验证方法；'
             '标准草案：列车级数字样机构建、智能体接口。\n子任务分工——子任务1 知识库与工作流（目标 1、3）；子任务2 调度引擎（目标 1）；'
             '子任务3 智能体执行（目标 2）；子任务4 集成与示范（目标 2、3、4）。各子任务交付物逐级成为下一任务输入，验证数据回流更新判据与工作流。')


# =====================================================================  16 量化指标与效益评估
def s16_metrics(d):
    s = d.new_slide()
    frame(d, s, '量化指标与效益评估：与现状对比', '指标：现状 → 目标 · 效益评估', '预期成果与效益')
    L, LW = 64, 872
    header(s, L, 212, '指标：与现状对比', w=LW)
    ph = '<bc=%s>【待填】</bc>' % PH.text
    rows = [['一次软件迭代的验证周期', '约 2 个月', '缩短 ≥ ' + ph + '%'], ['现场验证人员投入', '10 人', '减少 ≥ ' + ph + '%'],
            ['工况覆盖度', '受台架与现场条件限制', '提升 ≥ ' + ph + '%'], ['接入异构仿真智能体', '—', '≥ ' + ph + ' 类'],
            ['协同验证典型工况', '—', '≥ ' + ph + ' 个']]
    tb = table(s, L, 256, [330, 250, 292], rows, header=['指标', '现状', '目标'], size=20, lh=30, pad_y=41,
               col_styles={0: dict(color=INK, bold=True), 2: dict(color=INK, bold=True)})
    R, RW = 976, 880
    header(s, R, 212, '效益评估', w=RW)
    bens = [('技术效益', M1, '复杂工况可复现、可回归，形成列车级性能验证能力', None),
            ('经济效益', M3, '减少实物试验与样车试制投入，缩短研制迭代周期', [('减少实物试验', '≥【待填】次/年'), ('节约费用', '≥【待填】万元/年')]),
            ('管理与推广效益', M2, '验证流程标准化；可扩展到制动、辅助系统，在**内推广', None)]
    y, cg = 256, 14
    hs = [1.0, 1.4, 1.0]                          # 带目标值的“经济效益”卡高一些；三卡总高与左表齐平
    unit = (tb - 256 - cg * 2) / sum(hs)
    for (t, th, txt, tg), k in zip(bens, hs):
        bh = unit * k
        ch_ = 36 + 12 + 30 + (56 if tg else 0)   # 标题 + 间距 + 一行说明 +（目标值）
        top = y + (bh - ch_) / 2.0
        rect(s, R, y, RW, bh, fill=WHITE, line=LINE)
        rect(s, R, y, BAR, bh, fill=th.main)
        text(s, R + 28, top, 400, 36, t, size=22, bold=True, color=INK, anchor='m', lh=30, wrap=False)
        text(s, R + 28, top + 48, RW - 56, 32, txt, size=20, color=BODY, lh=30)
        if tg:
            xx = R + 28
            for lb, v in tg:
                xx += target_chip(s, xx, top + 96, lb, v) + 10
        y += bh + cg
    notes(s, '备注（讲稿）：目标值由课题组提供（【待填】）。效益评估依据——成果交付 ＋ 示范验证 ＋ 与现状对比；'
             '技术效益的成果形式见第 15 页：发明专利 4 项、核心论文 2 篇、企业标准草案 2 项。')


# =====================================================================  17 经费预算
def s17_budget(d):
    s = d.new_slide()
    frame(d, s, '经费预算：总预算 100 万元，算力是核心投入', '预算明细与测算依据 · 分年度 · 合理性说明', '经费预算合理性')
    L, LW = 64, 1064
    # 总额 + 堆叠条（比例只在这里画一次，表格里不再重复画迷你条）
    text(s, L, 206, 240, 80, '100<s=24><n><c=%s> 万元</c></n></s>' % SUB, size=60, bold=True, color=INK, font=MONO,
         lh=76, wrap=False)
    text(s, L, 280, 240, 28, '总预算', size=18, bold=True, color=MUTED, lh=24, wrap=False)
    stacked_bar(s, L + 250, 226, LW - 250, 56, [(v, c, lab, WHITE) for _, _, v, _, c, lab in BUDGET], size=19)
    rows = []
    for no, name, v, basis, col, _ in BUDGET:
        rows.append([no, name, '%d' % v, '%d%%' % v, basis])
    rows.append(['', '<b>合计</b>', '<b>100</b>', '<b>100%</b>', '—'])
    cols = [64, 300, 136, 96, 468]
    tb = table(s, L, 332, cols, rows, header=['序号', '科目', '金额（万元）', '占比', '测算依据'], size=19, lh=28, pad_x=14,
          pad_y=20, head_size=16,
          col_styles={0: dict(align='c', font=MONO, bold=True, color=SUB), 1: dict(bold=True, color=INK),
                      2: dict(align='c', font=MONO, bold=True, color=INK), 3: dict(align='r', font=MONO, color=SUB),
                      4: dict(color=BODY, size=18, lh=27)},
          fills={6: PANEL})
    # 右：分年度 + 合理性
    R, RW = 1168, 688
    header(s, R, 212, '分年度', w=RW)
    tw_ = (RW - 16) / 2.0
    for i, (yr, amt) in enumerate([('2027 年', '50'), ('2028 年', '50')]):
        x = R + i * (tw_ + 16)
        rect(s, x, 256, tw_, 96, fill=WHITE, line=LINE)
        text(s, x + 22, 256, 140, 96, yr, size=22, bold=True, color=SUB, anchor='m', lh=30, wrap=False)
        text(s, x + 140, 256, tw_ - 160, 96, '%s<s=20><n><c=%s> 万元</c></n></s>' % (amt, SUB), size=44, bold=True,
             color=INK, font=MONO, align='r', anchor='m', lh=56, wrap=False)
    header(s, R, 380, '合理性说明', w=RW)
    why = [('算力与任务强相关', 'NPU 做推理与编排决策，GPU 做多物理域并行仿真'),
           ('人员占比 30%', '以算法与软件开发为主'),
           ('无对外技术合作费', '全部用于自主研发')]
    cg = 16
    chh = (tb - 424 - cg * (len(why) - 1)) / float(len(why))
    for i, (t, b) in enumerate(why):
        y = 424 + i * (chh + cg)
        rect(s, R, y, RW, chh, fill=WHITE, line=LINE)
        rect(s, R, y, BAR, chh, fill=M1.main)
        n = wrap_lines(b, RW - 56, 20)
        top = y + (chh - (32 + 8 + n * 30)) / 2.0
        text(s, R + 28, top, RW - 56, 32, t, size=22, bold=True, color=INK, lh=30, wrap=False)
        text(s, R + 28, top + 40, RW - 56, n * 30 + 4, b, size=20, color=BODY, lh=30)
    notes(s, '备注（讲稿）：算力满足实时协同与保真度切换；人员属智力密集型研究。折旧——利用现有仿真与测试设备，避免重复购置；'
             '外协——数据采集与第三方验证；成果——专利、论文与标准草案；其他——示范应用差旅与会议。'
             '人员人工费用测算依据中【待填】为占位，请按“投入人月 × 人月费用标准”填写，合计应为 30 万元。')


# =====================================================================  18 项目管理计划
def s18_management(d):
    s = d.new_slide()
    frame(d, s, '项目管理计划：按四层分组承担，按里程碑绑定考核', '任务分配 · 项目流程 · 时间安排 · 执行保障', '项目管理计划')
    # 任务分配：组织架构（每组只写对应研究内容与交付物，主要任务进备注）
    header(s, 64, 212, '组织架构与任务分配', w=900)
    bx, bw = 960 - 170, 340
    rect(s, bx, 214, bw, 58, fill=ARCH)
    icon(s, 'user-star', WHITE, bx + 70, 214 + 14, 30)
    text(s, bx + 110, 214, bw - 130, 58, '项目负责人', size=23, bold=True, color=WHITE, anchor='m', lh=30, wrap=False)
    groups = [(1, '判据知识库、工作流定义'), (2, '协同运行与调度引擎'), (3, '仿真智能体、验证报告'), (4, '数字样机、方法体系与规范')]
    gw = (1792 - 3 * 24) / 4.0
    gy, gh = 322, 150
    line(s, 960, 272, 960, 296, MUTED, 1.5)
    line(s, 64 + gw / 2.0, 296, 64 + 3 * (gw + 24) + gw / 2.0, 296, MUTED, 1.5)
    for i, (k, deli) in enumerate(groups):
        th = LAYER[k]
        x = 64 + i * (gw + 24)
        line(s, x + gw / 2.0, 296, x + gw / 2.0, gy - 2, MUTED, 1.5, tail='triangle')
        rect(s, x, gy, gw, gh, fill=WHITE, line=LINE)
        rect(s, x, gy, gw, BAR, fill=th.main)
        tag(s, x + 22, gy + 20, '子任务 %d 组 · %s' % (k, LAYER_NO[k]), fill=th.main, size=17, h=30, padx=12)
        text(s, x + 22, gy + 62, gw - 44, 32, LAYER_NAME[k], size=22, bold=True, color=INK, lh=30, wrap=False)
        text(s, x + 22, gy + 104, gw - 44, 28, '<b>交付</b>　' + deli, size=18, color=SUB, lh=26, wrap=False)
    y1 = gy + gh + 44
    # 项目流程：管理节奏（阶梯，一级一句）
    L, LW = 64, 2 * gw + 24
    header(s, L, y1, '项目管理流程', w=LW)
    cad = [('双周例会', '进展同步<br>问题协调'), ('月度节点检查', '对照里程碑<br>核查交付物'), ('季度评审', '方案评审<br>成果确认'),
           ('年度总结', '目标评估<br>计划调整')]
    cw = (LW - 3 * 24) / 4.0
    top0, step = y1 + 60, 46
    for i, (n, t) in enumerate(cad):
        x = L + i * (cw + 24)
        top = top0 + (3 - i) * step
        rect(s, x, top, cw, 990 - top, fill=WHITE, line=LINE)
        rect(s, x, top, cw, STYLE['concl_bar'], fill=M1.main)
        text(s, x + 22, top + 24, cw - 44, 34, n, size=22, bold=True, color=INK, lh=30, wrap=False)
        text(s, x + 22, top + 70, cw - 40, 60, t, size=19, color=SUB, lh=28)
    # 时间安排：阶段条（只写阶段名）+ 里程碑日期
    R, RW = 64 + 2 * (gw + 24), 2 * gw + 24
    header(s, R, y1, '时间安排与执行保障', w=RW)
    mw = RW / 24.0
    sy, sh_ = y1 + 60, 52
    for name, short, a, b, col, per, goal, deli, mk, mdate, mtext in STAGES:
        x = R + a * mw + (1 if a else 0)
        w = (b - a) * mw - (1 if a else 0) - (1 if b < 24 else 0)
        fg = STAGE_ON[STAGE.index(col)] if STAGE_ON else on_color(col)
        shp = rect(s, x, sy, w, sh_, fill=col)
        text(s, x, sy, w, sh_, name, size=18, bold=True, color=fg, align='c', anchor='m', lh=24, shp=shp, wrap=False)
        lab_ = '%s · %s' % (mk, mdate)
        lw_ = text_width(lab_, 16, True) + 8
        cx = min(x + w / 2.0, R + RW - lw_ / 2.0)    # 末段较窄，标签不越出右边界
        text(s, cx - 80, sy + sh_ + 10, 160, 24, lab_, size=16, bold=True, color=RED, align='c', lh=22, wrap=False)
    # 执行保障：三行，风险一条一行
    ctl = [('质量', 'shield-check', '内部评审 ＋ 第三方验证，交付物可核查'),
           ('风险 → 应对', 'alert-triangle', '判据不全 → 首批聚焦三项关键性能<br>同步失真 → 分层步长、关键工况全保真校核<br>'
                                          '极端工况数据少 → 实测数据标定、第三方验证'),
           ('资源', 'server', '算力平台、HIL 台架、FXN5C 试验与运用数据')]
    y = sy + sh_ + 56
    nl, lh_ = [1, 3, 1], 28
    padr = (990 - y - sum(n * lh_ for n in nl)) / (2.0 * len(nl))
    line(s, R, y, R + RW, y, LINE, 1)
    for (t, ic, body), n in zip(ctl, nl):
        ty_ = y + padr
        icon(s, ic, RED if t.startswith('风险') else INK, R, ty_ - 1, 28)
        text(s, R + 40, ty_ - 2, 200, 32, t, size=20, bold=True, color=INK, lh=28, wrap=False)
        text(s, R + 230, ty_, RW - 230, n * lh_ + 4, body, size=18, color=BODY, lh=lh_)
        y += n * lh_ + 2 * padr
        line(s, R, y, R + RW, y, LINE, 1)
    notes(s, '备注（讲稿）：各组主要任务——子任务1组：判据知识化、工作流建模；子任务2组：任务解析、环境装配、保真度与算力调度；'
             '子任务3组：智能体研制、用例生成、识别与报告；子任务4组：架构、沙箱、样机集成、FXN5C 示范。\n'
             '管理节奏——双周例会：子任务进展同步与问题协调；月度：对照里程碑核查交付物；季度：技术方案与阶段成果评审；年度：目标评估与计划调整。\n'
             '时间安排——M1 架构与接口规范（2027.06）；M2 智能体与引擎（2028.03）；M3 数字样机与验证（2028.09）；M4 示范应用与验收（2028.12）。'
             '接口与进度风险——接口规范在阶段一先行交付，里程碑绑定考核。')


# =====================================================================  19 团队研究能力
def s19_team(d):
    s = d.new_slide()
    frame(d, s, '团队研究能力：为什么由我们承担', '牵头单位：********** · 六项现有基础', '团队研究能力')
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
    frame(d, s, '成果介绍：五项成果 ＋ 知识产权与标准', '软件 · 模型 · 系统 · 规范 · 报告', '成果介绍')
    top = [
        ('成果1', '软件1套', '多智能体协同运行与调度引擎', M1, 'settings-automation',
         '输入验证需求，自动完成<b>分解、装配、调度、监控与报告</b>', ('调度引擎软件', '建议：软件界面截图或系统架构图')),
        ('成果2', '1套', '关键系统仿真智能体／<br>智能孪生模型', M2, 'robot',
         '控制系统、车辆动力学、线路等，按实测数据<b>自动标定</b>', ('仿真智能体／孪生模型', '建议：模型结构图或仿真运行界面截图')),
        ('成果3', '1套', '列车级数字样机', M3, 'train',
         '以 <b>FXN5C</b> 为对象，用于回归验证与问题复现', ('列车级数字样机', '建议：数字样机三维模型或运行界面截图')),
    ]
    pw, ph, y0 = (1792 - 48) / 3.0, 440, 214
    for i, (no, form, name, th, ic, desc, (ft, hint)) in enumerate(top):
        x = 64 + i * (pw + 24)
        rect(s, x, y0, pw, ph, fill=WHITE, line=LINE)
        rect(s, x, y0, pw, BAR, fill=th.main)
        icon_badge(s, ic, th, x + 24, y0 + 26, 56)
        text(s, x + 96, y0 + 22, 300, 26, '<k>%s</k>　<c=%s>%s</c>' % (no, SUB, form), size=17, color=th.text, lh=24,
             wrap=False)
        text(s, x + 96, y0 + 48, pw - 120, 62, name, size=23, bold=True, color=INK, lh=30)
        fig_placeholder(s, x + 24, y0 + 122, pw - 48, 210, ft, hint)
        text(s, x + 24, y0 + 352, pw - 48, 64, desc, size=19, color=BODY, lh=29)
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
    notes(s, '备注（讲稿）：成果 1 支持异构仿真智能体统一接入；成果 2 标准接口可扩展；成果 3 由成果 2 经成果 1 集成，预留制动、辅助系统接口，'
             '支撑智能运维。\n成果 1—3 中的虚线框为附图预留位置（形状组合）；选中该组合删除后插入实际图片即可。')


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
