# -*- coding: utf-8 -*-
"""用 PowerPoint COM 组装 10 页汇报 PPT（+3 页备用附录），并导出每页预览。"""
import os
import time
from pathlib import Path

import win32com.client
from PIL import Image

BASE = Path(__file__).resolve().parents[2]
FIG = BASE / 'outputs' / 'report' / 'figures'
OUT = BASE / 'outputs' / 'report'
PREVIEW = OUT / 'preview'
PREVIEW.mkdir(parents=True, exist_ok=True)
PPT_PATH = OUT / '智慧审计_机器学习_10页汇报.pptx'

# ---------------- 颜色与字体 ----------------
FONT = 'Microsoft YaHei'
DARK = '#1B2631'
TEXT = '#34495E'
GRAY = '#7F8C8D'
LIGHT_GRAY = '#F4F6F7'
BORDER = '#D5D8DC'
PRIMARY = '#2E86C1'
PRIMARY_DARK = '#1F4E79'
PRIMARY_LIGHT = '#AED6F1'
ACCENT = '#E67E22'
ACCENT_LIGHT = '#FDEBD0'
GREEN = '#27AE60'
RED = '#C0392B'
WHITE = '#FFFFFF'


def rgb(hexstr):
    h = hexstr.lstrip('#')
    r, g, b = int(h[0:2], 16), int(h[2:4], 16), int(h[4:6], 16)
    return b << 16 | g << 8 | r  # PowerPoint COM 使用 BGR 顺序


def set_font(tr, size, color, bold=False):
    tr.Font.Name = FONT
    tr.Font.NameFarEast = FONT
    tr.Font.NameAscii = FONT
    tr.Font.Size = size
    tr.Font.Bold = -1 if bold else 0
    tr.Font.Color.RGB = rgb(color)


def add_text(slide, text, left, top, width, height, size=12, color=TEXT, bold=False,
             align=1, anchor=1, spacing=1.05):
    shp = slide.Shapes.AddTextbox(1, left, top, width, height)
    tf = shp.TextFrame
    tf.WordWrap = -1
    tf.MarginLeft = 0
    tf.MarginRight = 0
    tf.MarginTop = 0
    tf.MarginBottom = 0
    tf.VerticalAnchor = anchor
    tr = tf.TextRange
    tr.Text = text
    set_font(tr, size, color, bold)
    tr.ParagraphFormat.Alignment = align
    try:
        tr.ParagraphFormat.SpaceWithin = spacing
    except Exception:
        pass
    return shp


def add_box(slide, left, top, width, height, fill=None, line=None, radius=True, shadow=False):
    shape_type = 5 if radius else 1
    shp = slide.Shapes.AddShape(shape_type, left, top, width, height)
    if fill is None:
        shp.Fill.Visible = 0
    else:
        shp.Fill.Visible = -1
        shp.Fill.Solid()
        shp.Fill.ForeColor.RGB = rgb(fill)
    if line is None:
        shp.Line.Visible = 0
    else:
        shp.Line.Visible = -1
        shp.Line.ForeColor.RGB = rgb(line)
        shp.Line.Weight = 0.9
    if not shadow:
        try:
            shp.Shadow.Visible = 0
        except Exception:
            pass
    return shp


def add_picture_fit(slide, path, left, top, max_w, max_h):
    with Image.open(path) as im:
        iw, ih = im.size
    scale = min(max_w / iw, max_h / ih)
    w, h = int(iw * scale), int(ih * scale)
    x = left + (max_w - w) // 2
    y = top + (max_h - h) // 2
    slide.Shapes.AddPicture(str(path), 0, -1, x, y, w, h)
    return x, y, w, h


def add_title(slide, title, subtitle=None):
    add_box(slide, 40, 30, 6, 34, fill=ACCENT, radius=False)
    add_text(slide, title, 56, 24, 860, 42, size=23, color=DARK, bold=True)
    if subtitle:
        add_text(slide, subtitle, 56, 72, 860, 46, size=12.5, color=GRAY, spacing=1.1)


def add_footer(slide, idx):
    add_box(slide, 40, 518, 880, 1.0, fill=BORDER, radius=False)
    add_text(slide, '智慧审计 × 机器学习  |  期中实践', 40, 523, 420, 14, 8.5, GRAY)
    add_text(slide, f'{idx:02d} / 10', 840, 523, 80, 14, 8.5, GRAY, align=3)


def add_stat_strip(slide, text, y=112):
    add_box(slide, 48, y - 5, 864, 32, fill=LIGHT_GRAY, radius=True)
    add_text(slide, text, 62, y, 836, 22, size=12.5, color=ACCENT, bold=True, anchor=3)


def add_bullets(slide, items, left, top, width, height, size=11, color=TEXT, spacing=1.25):
    add_text(slide, '\r'.join('• ' + x for x in items), left, top, width, height,
             size=size, color=color, spacing=spacing)


def add_appendix_footer(slide, idx, total=3):
    add_box(slide, 40, 518, 880, 1.0, fill=BORDER, radius=False)
    add_text(slide, '智慧审计 × 机器学习  |  期中实践  |  附录 · 备用页（不计入 10 页主线）',
             40, 523, 620, 14, 8.5, GRAY)
    add_text(slide, f'A-{idx} / {total}', 840, 523, 80, 14, 8.5, GRAY, align=3)


def add_table(slide, headers, rows, left, top, col_widths, row_h=38, header_h=34,
              aligns=None, sizes=None, colors=None, bolds=None, highlight_rows=None):
    """手绘表格，避免 COM 表格样式在不同 Office 版本下不一致。"""
    n = len(headers)
    aligns = aligns or [1] * n
    sizes = sizes or [10] * n
    colors = colors or [DARK] * n
    bolds = bolds or [False] * n

    x = left
    for j, (h, w) in enumerate(zip(headers, col_widths)):
        add_box(slide, x, top, w, header_h, fill=PRIMARY_DARK, line=WHITE, radius=False)
        add_text(slide, h, x + 8, top + 7, w - 16, header_h - 10,
                 size=10.5, color=WHITE, bold=True, align=aligns[j], anchor=3)
        x += w

    y = top + header_h
    for i, row in enumerate(rows):
        fill = '#F8F9F9' if i % 2 == 0 else WHITE
        highlight = bool(highlight_rows and i in highlight_rows)
        if highlight:
            fill = '#EBF5FB'
        x = left
        for j, (cell, w) in enumerate(zip(row, col_widths)):
            add_box(slide, x, y, w, row_h, fill=fill, line=BORDER, radius=False)
            add_text(slide, str(cell), x + 8, y + 5, w - 16, row_h - 10,
                     size=sizes[j], color=colors[j], bold=bolds[j] or highlight,
                     align=aligns[j], anchor=3)
            x += w
        y += row_h
    return y


# ---------------- 启动 PowerPoint ----------------
app = win32com.client.DispatchEx('PowerPoint.Application')
app.Visible = True
app.DisplayAlerts = 1
pres = None
try:
    pres = app.Presentations.Add()
    pres.PageSetup.SlideWidth = 960
    pres.PageSetup.SlideHeight = 540

    def new_slide(bg='#FFFFFF'):
        s = pres.Slides.Add(pres.Slides.Count + 1, 12)
        s.FollowMasterBackground = 0
        s.Background.Fill.Solid()
        s.Background.Fill.ForeColor.RGB = rgb(bg)
        return s

    # ============================================================
    # 1 封面
    # ============================================================
    s = new_slide(PRIMARY_DARK)
    # 装饰圆
    circle = s.Shapes.AddShape(9, 690, -110, 430, 430)
    circle.Fill.Visible = -1
    circle.Fill.Solid()
    circle.Fill.ForeColor.RGB = rgb('#2E6DA4')
    circle.Fill.Transparency = 0.55
    circle.Line.Visible = 0
    circle2 = s.Shapes.AddShape(9, 780, 330, 250, 250)
    circle2.Fill.Visible = -1
    circle2.Fill.Solid()
    circle2.Fill.ForeColor.RGB = rgb('#2E6DA4')
    circle2.Fill.Transparency = 0.65
    circle2.Line.Visible = 0

    add_text(s, '期中实践 · 智慧审计', 70, 62, 500, 26, 11, PRIMARY_LIGHT, bold=True)
    add_box(s, 70, 128, 8, 150, fill=ACCENT, radius=False)
    add_text(s, '智慧审计 × 机器学习', 96, 120, 720, 70, 40, WHITE, bold=True)
    add_text(s, 'Financial Audit Transactions 异常识别实践', 96, 198, 720, 36, 18, PRIMARY_LIGHT)
    add_text(s, '问题：7,000 笔交易人工逐笔看不完；规则只看单字段；异常样本少且标签不可靠。',
             96, 248, 800, 22, 10.5, '#F5CBA7', bold=True)
    add_text(s, '目标：用 2% 审计容量，从 7,000 笔交易中筛出 140 条最值得核查的样本',
             96, 282, 800, 30, 13, WHITE, bold=True)
    add_text(s, '方法：Benford 首位数 + Isolation Forest + Top-K + 2×2 + 人工归因',
             96, 316, 800, 26, 10.5, '#D6EAF8')
    add_text(s, '期中实践汇报  ·  2026 年 10 月', 96, 430, 500, 24, 11, PRIMARY_LIGHT)
    add_text(s, '全流程 CPU  ·  免费工具  ·  约 3 天', 640, 430, 250, 24, 10, PRIMARY_LIGHT, align=3)

    # ============================================================
    # 2 学了什么：流水线
    # ============================================================
    s = new_slide()
    add_title(s, '学了什么：一条可复现的审计数据流水线',
              '针对上面的问题，我把解决过程拆成四步；从原始 CSV 到 140 条名单，每步都有留痕。')

    boxes = [
        ('① 获取', PRIMARY, '7,000 × 59\n0 缺失\n955 条异常（13.64%）\n11 组字段（含标签）', '→ audit_raw.csv'),
        ('② 清洗', ACCENT, '59 → 115 维\n剔除 13 列\n79 个 One-Hot\n4 个金额 log1p+z', '→ audit_clean.csv'),
        ('③ 分析', GREEN, 'Benford MAD 0.00405\nIF 标记 140 条\nTop-140 精确率 92.1%\n2×2 规则—模型互补', '→ scores / top140 名单'),
        ('④ 汇报', PRIMARY_DARK, '6 张统一风格图\n10 页 PPT（+3 页附录）\n3 条人工归因案例\n答辩问答备忘', '→ figures / 本 PPT'),
    ]
    xs = [40, 270, 500, 730]
    for x, (head, color, body, trace) in zip(xs, boxes):
        add_box(s, x, 142, 205, 222, fill=WHITE, line=BORDER, radius=True)
        add_box(s, x, 142, 205, 7, fill=color, radius=False)
        add_text(s, head, x + 16, 162, 175, 30, 16, color, bold=True)
        add_text(s, body, x + 16, 204, 175, 100, 10.5, TEXT, spacing=1.28)
        add_box(s, x + 16, 306, 173, 0.8, fill=BORDER, radius=False)
        add_text(s, trace, x + 16, 316, 175, 18, 8.5, GRAY)
    for x in [248, 478, 708]:
        arrow = s.Shapes.AddShape(33, x, 244, 18, 26)
        arrow.Fill.Visible = -1
        arrow.Fill.Solid()
        arrow.Fill.ForeColor.RGB = rgb('#BFC9CA')
        arrow.Line.Visible = 0

    add_box(s, 40, 392, 880, 58, fill=LIGHT_GRAY, radius=True)
    add_text(s, '全程 CPU、免费工具、约 3 天完成；不需要 GPU、深度学习框架、数据库或 SHAP。',
             60, 404, 840, 22, 12, TEXT, bold=True)
    add_text(s, '每个关键中间结果都保存为 CSV / 图 / 日志，方便复查和答辩。',
             60, 426, 840, 18, 9.5, GRAY)

    # ============================================================
    # 3 为何选审计
    # ============================================================
    s = new_slide()
    add_title(s, '为何选审计：三个条件刚好都满足')
    add_text(s, '审计现场：人力有限、异常稀少、规则只看单字段 → 需要无监督 + 可解释的 Top-K 筛查。',
             56, 68, 860, 22, 12.5, TEXT, bold=True)
    add_text(s, '不是“为了用机器学习而用”，而是这个场景同时具备可验证、可解释、可落地的条件。',
             56, 92, 860, 22, 12, GRAY)
    cards = [
        ('01', '有答案可验证', PRIMARY,
         ['955 条异常标签，占 13.64%',
          '8 类异常，可做 Top-K 评估',
          '不只看准确率，看命中率与召回率'],
         '验证：Top-K 精确率 / 召回率'),
        ('02', '有业务可解释', ACCENT,
         ['金额、审批、发票、供应商、关系网络',
          '59 个字段能落到具体审计动作',
          '异常不只是一个分数，而是一条线索'],
         '解释：金额 · 审批 · 发票 · 供应商'),
        ('03', '有规模可落地', GREEN,
         ['7,000 行 × 59 列，约 3.7 MB',
          'Isolation Forest 训练秒级',
          '个人 3 天可完成，不依赖 GPU'],
         '落地：Top-K 名单 + 人工核查'),
    ]
    xs = [40, 340, 640]
    for x, (num, head, color, items, foot) in zip(xs, cards):
        add_box(s, x, 132, 280, 262, fill=WHITE, line=BORDER, radius=True)
        add_box(s, x, 132, 280, 7, fill=color, radius=False)
        add_text(s, num, x + 18, 150, 60, 34, 24, color, bold=True)
        add_text(s, head, x + 18, 192, 244, 32, 16, DARK, bold=True)
        add_bullets(s, items, x + 18, 238, 244, 90, size=10.5, spacing=1.28)
        add_box(s, x + 18, 326, 244, 0.8, fill=BORDER, radius=False)
        add_text(s, foot, x + 18, 336, 244, 18, 9.5, GRAY)

    add_box(s, 40, 418, 880, 62, fill='#EBF5FB', radius=True)
    add_text(s, '审计不是追求 100% 准确，而是用有限人力优先核查高风险样本。',
             62, 430, 830, 26, 15, PRIMARY_DARK, bold=True)
    add_text(s, '这正是 Top-K 召回要解决的问题：先保证名单质量，再决定查多少。',
             62, 454, 830, 20, 10, GRAY)

    # ============================================================
    # 4 获取数据
    # ============================================================
    s = new_slide()
    add_title(s, '获取数据：先看清家底，再决定怎么洗',
              '探查后修正：数据零缺失，无需填充；Abnormal_Type / Risk_Class 是标签/标签派生列，建模剔除。')
    add_stat_strip(s, '7,000 行  ·  59 列  ·  955 条异常（13.64%）  ·  0 缺失  ·  11 组字段（含标签）')
    add_picture_fit(s, FIG / 'fig1_data_profile.png', 40, 150, 880, 360)
    add_footer(s, 4)

    # ============================================================
    # 5 清洗（一）
    # ============================================================
    s = new_slide()
    add_title(s, '清洗（一）：不删一行，把 59 列整理成 115 维',
              '常量列、泄漏列、ID/名称列剔除；低/中基数类别 One-Hot；城市不编码，保留省份。')
    add_stat_strip(s, '0 行删除  ·  13 列剔除  ·  79 个 One-Hot  ·  115 维 = 111 特征 + 2 标识 + 2 标签')
    add_picture_fit(s, FIG / 'fig2_cleaning.png', 40, 150, 880, 360)
    add_footer(s, 5)

    # ============================================================
    # 6 清洗（二）
    # ============================================================
    s = new_slide()
    add_title(s, '清洗（二）：金额长尾不处理，异常分数会被极端值带偏')
    add_text(s, '先 log1p 压缩量级，再 z-score 标准化；4 个金额/比率字段统一处理。',
             56, 66, 860, 22, 12, GRAY)
    add_text(s, '关键发现：131 笔金额恰好 500,000 元，全部为金额突增异常——log1p 后形成 13.1 处尖峰。',
             56, 90, 860, 22, 12, ACCENT, bold=True)
    add_stat_strip(s, '偏度 10.70 → 0.52  ·  4 个字段  ·  9 项质量检查通过  ·  7,000 行不变')
    add_picture_fit(s, FIG / 'fig3_amount_transform.png', 40, 150, 880, 360)
    add_footer(s, 6)

    # ============================================================
    # 7 Benford
    # ============================================================
    s = new_slide()
    add_title(s, '分析（一）：Benford 基线——整体正常，异常组露馅',
              '整体 MAD 0.00405（高度符合）；正常组 0.00325；异常组 0.03633（明显不符合）。')
    add_picture_fit(s, FIG / 'fig4_benford.png', 40, 142, 880, 365)
    add_footer(s, 7)

    # ============================================================
    # 8 Top-K
    # ============================================================
    s = new_slide()
    add_title(s, '分析（二）：模型把异常排在前面，Top-K 决定查多少',
              'contamination=0.02 对应 140 条审计容量；分数排序不变，K 可按人力调整。')
    add_picture_fit(s, FIG / 'fig5_topk.png', 40, 142, 880, 365)
    add_footer(s, 8)

    # ============================================================
    # 9 规则 × 模型
    # ============================================================
    s = new_slide()
    add_title(s, '理解与修正：规则和模型不是二选一，而是互补',
              '仅模型 122 条（异常率 91.8%）；同时命中 18 条（94.4%）→ 低重合、高互补\r'
              '案例信号：审批延迟 339 min + 发票匹配 0.44  ·  供应商风险 0.86  ·  金额/阈值 97.4%')
    add_picture_fit(s, FIG / 'fig6_rule_model.png', 40, 138, 880, 330)
    add_box(s, 40, 478, 880, 32, fill=ACCENT_LIGHT, radius=True)
    add_text(s, '建议流程：Benford 初筛（宽口径）→ 模型排序（高精度）→ 人工核查（18 条双命中优先）',
             58, 484, 844, 20, 10.5, '#935116', bold=True)
    add_footer(s, 9)

    # ============================================================
    # 10 收束
    # ============================================================
    s = new_slide()
    add_title(s, '收束：2% 审计容量下，Top-140 精确率 92.1%',
              '覆盖 7/8 类异常；Benford 与模型低重合、高互补；跨省不匹配是短板。')
    cols = [
        ('数据', PRIMARY, ['7,000 × 59，0 缺失', '59 → 115 维，不删一行', '955 条异常，6.33 : 1 不平衡'], '→ audit_clean.csv'),
        ('方法', ACCENT, ['Benford 看整体分布', 'Isolation Forest 看多维组合', '规则 × 模型 2×2 交叉'], '→ scores / flags'),
        ('结果', GREEN, ['Top-140 精确率 92.1%', '覆盖 7 / 8 类异常', '3 条可解释案例，假阳性 7.9%'], '→ top140_audit_list.csv'),
    ]
    xs = [40, 340, 640]
    for x, (head, color, items, trace) in zip(xs, cols):
        add_box(s, x, 126, 280, 210, fill=WHITE, line=BORDER, radius=True)
        add_box(s, x, 126, 280, 7, fill=color, radius=False)
        add_text(s, head, x + 18, 146, 244, 32, 17, color, bold=True)
        add_bullets(s, items, x + 18, 194, 244, 80, size=10.5, spacing=1.35)
        add_box(s, x + 18, 288, 244, 0.8, fill=BORDER, radius=False)
        add_text(s, trace, x + 18, 298, 244, 18, 8.8, GRAY)

    add_box(s, 40, 354, 880, 108, fill=LIGHT_GRAY, radius=True)
    add_text(s, '局限与下一步', 60, 366, 300, 24, 12, DARK, bold=True)
    add_bullets(s, [
        '无监督：标签只用于评估；contamination=0.02 是审计容量假设，不是真实异常率。',
        '短板：Cross_Province_Mismatch 115 条未进入 Top-140；可补图特征或有监督模型。',
        '下一步：接入真实凭证/流水，把 Top-140 名单做成审计工作台。',
    ], 60, 392, 840, 62, size=9.8, spacing=1.18)

    add_box(s, 40, 476, 880, 30, fill='#EBF5FB', radius=True)
    add_text(s, '结论：无监督筛查能有效缩小核查范围，但最终判断仍需人工凭证核查。',
             56, 482, 848, 20, 10.5, PRIMARY_DARK, bold=True, align=2)
    add_footer(s, 10)

    # ============================================================
    # 附录 A-1：115 维构成
    # ============================================================
    s = new_slide()
    add_title(s, '附录 A-1（备用）：115 维到底由什么组成',
              '原始 59 列 → 剔除 13 列 → 46 列 → One-Hot 后 115 列；0 行删除、0 缺失。')
    add_table(
        s,
        ['组成部分', '列数', '具体内容'],
        [
            ('One-Hot 类别', '79', '银行 9+9、省份 12+12、账户类型 6+6、支付渠道 7、发票类型 6、审批级别 4、边类型 8'),
            ('数值特征', '28', '时间、账龄、评分、计数等'),
            ('金额 / 比率 log1p+z', '4', '交易金额、历史均值、交易/均值比、阈值比'),
            ('标识列', '2', 'Transaction_ID、Transaction_DateTime（不进入模型）'),
            ('标签列', '2', 'Abnormal_Label、Abnormal_Type（只用于评估/对照）'),
            ('合计', '115', '其中 111 个特征参与建模'),
        ],
        40, 146, [170, 80, 630],
        row_h=38, header_h=34,
        aligns=[1, 2, 1],
        sizes=[10.5, 11, 9.3],
        colors=[DARK, ACCENT, TEXT],
        bolds=[True, True, False],
        highlight_rows=[5],
    )
    add_box(s, 40, 430, 880, 32, fill='#EBF5FB', radius=True)
    add_text(s, '115 = 79 个 One-Hot + 28 个数值 + 4 个金额/比率 + 2 个标识 + 2 个标签；建模只用 111 个特征。',
             58, 436, 844, 20, 10.5, PRIMARY_DARK, bold=True)
    add_text(s, '被剔除的 13 列：常量 Currency、泄漏 Risk_Class、节点/实体 ID 与名称、Invoice_ID、两个城市列。',
             56, 474, 860, 18, 9.5, GRAY)
    add_appendix_footer(s, 1)

    # ============================================================
    # 附录 A-2：Top-K 全表
    # ============================================================
    s = new_slide()
    add_title(s, '附录 A-2（备用）：Top-K 全表——K 由审计人力决定',
              '分数排序不变；K 越大，召回越高、精确率越低。K=140 对应本项目 2% 审计容量。')
    add_table(
        s,
        ['K', '命中', '精确率', '召回率', 'lift', '适用场景'],
        [
            ('50', '49', '98.0%', '5.1%', '7.18×', '只能查半天：先抓最稳的'),
            ('100', '96', '96.0%', '10.1%', '7.04×', '1 人 1 天'),
            ('140', '129', '92.1%', '13.5%', '6.75×', '本项目：2% 审计容量'),
            ('300', '259', '86.3%', '27.1%', '6.33×', '3 人 2 天'),
            ('500', '380', '76.0%', '39.8%', '5.57×', '扩大核查范围'),
            ('1000', '578', '57.8%', '60.5%', '4.24×', '接近“全量初筛”'),
        ],
        40, 146, [80, 90, 110, 110, 100, 390],
        row_h=38, header_h=34,
        aligns=[2, 2, 2, 2, 2, 1],
        sizes=[10.5, 10, 10.5, 10.5, 10.5, 9.5],
        colors=[ACCENT, TEXT, GREEN, TEXT, PRIMARY_DARK, TEXT],
        bolds=[True, False, True, False, True, False],
        highlight_rows=[2],
    )
    add_box(s, 40, 430, 880, 32, fill='#EBF5FB', radius=True)
    add_text(s, '精确率 = 命中 / K；召回率 = 命中 / 955；lift = 精确率 / 13.64%。',
             58, 436, 844, 20, 10.5, PRIMARY_DARK, bold=True)
    add_text(s, 'K 不是模型参数，是审计人力参数；分数排序不变，改 K 即可。',
             56, 474, 860, 18, 9.5, GRAY)
    add_appendix_footer(s, 2)

    # ============================================================
    # 附录 A-3：一条案例怎么读
    # ============================================================
    s = new_slide()
    add_title(s, '附录 A-3（备用）：一条案例怎么读——从原始字段到模型分数',
              '模型排名 1 / 7,000 的重复发票案例；Benford 未标记，模型 Top-140 命中。')

    add_box(s, 40, 136, 420, 300, fill=WHITE, line=BORDER, radius=True)
    add_text(s, '原始字段（节选）', 58, 148, 380, 24, 12, PRIMARY_DARK, bold=True)
    fields = [
        ('交易 ID', 'CN-TXN-0005220'),
        ('交易金额', '¥14,774.56'),
        ('历史均值', '¥25,215.21'),
        ('审批延迟', '338.55 min'),
        ('发票匹配', '0.4411'),
        ('规则违反', '4 条'),
        ('关系异常', '0.4931'),
        ('模型分数', '0.5099（排名 1 / 7,000）'),
    ]
    y = 180
    for i, (k, v) in enumerate(fields):
        fill = '#F8F9F9' if i % 2 == 0 else WHITE
        add_box(s, 56, y, 160, 28, fill=fill, line=BORDER, radius=False)
        add_text(s, k, 66, y + 4, 140, 20, 9.5, GRAY, anchor=3)
        add_box(s, 216, y, 228, 28, fill=fill, line=BORDER, radius=False)
        signal = k in ('审批延迟', '发票匹配', '规则违反', '关系异常')
        add_text(s, v, 226, y + 4, 208, 20, 9.5, ACCENT if signal else DARK,
                 bold=signal, anchor=3)
        y += 28

    add_text(s, '为什么模型把它排在第 1 名', 480, 148, 440, 24, 12, ACCENT, bold=True)
    cards = [
        ('338.55', '审批延迟 min', '中位数 59 min · 99.8 分位', PRIMARY),
        ('0.4411', '发票匹配', '中位数 0.915 · 最低 1.1%', ACCENT),
        ('4 条', '规则违反', '中位数 1 条', GREEN),
        ('0.4931', '关系异常', '中位数 0.269 · 99.6 分位', PRIMARY_DARK),
    ]
    positions = [(480, 178), (700, 178), (480, 296), (700, 296)]
    for (x, top), (num, label, note, color) in zip(positions, cards):
        add_box(s, x, top, 210, 106, fill=WHITE, line=BORDER, radius=True)
        add_box(s, x, top, 210, 5, fill=color, radius=False)
        add_text(s, num, x + 14, top + 12, 182, 30, 20, color, bold=True)
        add_text(s, label, x + 14, top + 44, 182, 18, 10, DARK, bold=True)
        add_text(s, note, x + 14, top + 66, 182, 30, 8.3, GRAY, spacing=1.0)

    add_text(s, '金额只有 1.48 万，低于历史均值 2.52 万；模型不是单看金额，而是看审批、发票、规则和关系网络的组合。',
             56, 448, 850, 18, 9.8, TEXT)
    add_box(s, 40, 474, 880, 30, fill=ACCENT_LIGHT, radius=True)
    add_text(s, '人工核查动作：调发票号、合同、审批链，确认是否重复报销。案例是线索，不是定罪。',
             58, 480, 844, 18, 10.5, '#935116', bold=True)
    add_appendix_footer(s, 3)

    # ---------------- 保存与导出 ----------------
    pres.SaveAs(str(PPT_PATH))
    print('saved pptx:', PPT_PATH, os.path.exists(PPT_PATH))

    for i in range(1, pres.Slides.Count + 1):
        try:
            app.ActiveWindow.View.GotoSlide(i)
        except Exception:
            pass
        png = PREVIEW / f'slide{i:02d}.png'
        pres.Slides(i).Export(str(png), 'PNG', 1600, 900)
        print('exported', png, os.path.exists(png), os.path.getsize(png) if os.path.exists(png) else 0)

    print('slides:', pres.Slides.Count)
finally:
    try:
        if pres is not None:
            pres.Saved = -1
            pres.Close()
    except Exception as e:
        print('close pres error:', e)
    try:
        app.Quit()
    except Exception as e:
        print('quit error:', e)
    time.sleep(1)
