# -*- coding: utf-8 -*-
"""生成《函数的极值》教学课件 PPTX（16:9，14页）"""
from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE
from pptx.oxml.ns import qn

OUT = "/workspace/试讲-函数的极值/教学课件-函数的极值.pptx"
FIG1 = "/workspace/试讲-函数的极值/assets/fig_main.png"
FIG2 = "/workspace/试讲-函数的极值/assets/fig_counter.png"

NAVY = RGBColor(0x1F, 0x38, 0x64)
NAVY_MID = RGBColor(0x2E, 0x53, 0x95)
ORANGE = RGBColor(0xE8, 0x77, 0x22)
TEXT = RGBColor(0x33, 0x3A, 0x45)
GRAY = RGBColor(0x6B, 0x73, 0x80)
LIGHT = RGBColor(0xF3, 0xF5, 0xF9)
LINE = RGBColor(0xD9, 0xDE, 0xE7)
WHITE = RGBColor(0xFF, 0xFF, 0xFF)

CN = "微软雅黑"
EN = "Times New Roman"

prs = Presentation()
prs.slide_width = Inches(13.333)
prs.slide_height = Inches(7.5)
BLANK = prs.slide_layouts[6]

TOTAL = 14
_page_no = [0]


def add_text(slide, l, t, w, h, runs, size=20, bold=False, color=TEXT,
             align=PP_ALIGN.LEFT, anchor=MSO_ANCHOR.TOP, line=1.35,
             cn=CN, en=EN, space_before=0):
    """runs 可为字符串，或 [(text, dict), ...] 的段落列表"""
    box = slide.shapes.add_textbox(Inches(l), Inches(t), Inches(w), Inches(h))
    tf = box.text_frame
    tf.word_wrap = True
    tf.margin_left = 0
    tf.margin_right = 0
    tf.margin_top = 0
    tf.margin_bottom = 0
    tf.vertical_anchor = anchor

    if isinstance(runs, str):
        paras = [(runs, {})]
    else:
        paras = runs

    for i, (txt, opt) in enumerate(paras):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.alignment = opt.get("align", align)
        p.line_spacing = opt.get("line", line)
        p.space_before = Pt(opt.get("space_before", space_before if i > 0 else 0))
        p.space_after = Pt(opt.get("space_after", 0))
        r = p.add_run()
        r.text = txt
        r.font.size = Pt(opt.get("size", size))
        r.font.bold = opt.get("bold", bold)
        r.font.color.rgb = opt.get("color", color)
        r.font.name = opt.get("en", en)
        # 中文字体
        rPr = r._r.get_or_add_rPr()
        ea = rPr.find(qn("a:ea"))
        if ea is None:
            ea = rPr.makeelement(qn("a:ea"), {})
            rPr.append(ea)
        ea.set("typeface", opt.get("cn", cn))
    return box


def _strip_style(shp):
    """去掉自选图形自带的主题样式（阴影等），避免渲染出灰影"""
    sp = shp._element
    style = sp.find(qn("p:style"))
    if style is not None:
        sp.remove(style)


def card(slide, l, t, w, h, fill=LIGHT, line_color=None, radius=0.04):
    shp = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE,
                                 Inches(l), Inches(t), Inches(w), Inches(h))
    shp.adjustments[0] = radius
    shp.fill.solid()
    shp.fill.fore_color.rgb = fill
    if line_color is None:
        shp.line.fill.background()
    else:
        shp.line.color.rgb = line_color
        shp.line.width = Pt(1)
    shp.shadow.inherit = False
    _strip_style(shp)
    shp.text_frame.text = ""
    return shp


def bar(slide, l, t, w, h, color):
    shp = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE,
                                 Inches(l), Inches(t), Inches(w), Inches(h))
    shp.fill.solid()
    shp.fill.fore_color.rgb = color
    shp.line.fill.background()
    shp.shadow.inherit = False
    _strip_style(shp)
    return shp


def new_slide(title, tag=None, page=True):
    s = prs.slides.add_slide(BLANK)
    # 顶部标题
    add_text(s, 0.72, 0.42, 10.6, 0.62, title, size=28, bold=True, color=NAVY)
    # 标题下短橙线
    bar(s, 0.74, 1.06, 0.62, 0.055, ORANGE)
    bar(s, 1.36, 1.067, 11.6, 0.012, LINE)
    if tag:
        add_text(s, 10.4, 0.5, 2.6, 0.4, tag, size=13, color=GRAY,
                 align=PP_ALIGN.RIGHT)
    if page:
        _page_no[0] = len(prs.slides._sldIdLst)
        add_text(s, 12.2, 6.92, 0.8, 0.3, f"{_page_no[0]}", size=11,
                 color=GRAY, align=PP_ALIGN.RIGHT)
        add_text(s, 0.72, 6.92, 6.0, 0.3, "§6.2 函数的极值", size=11,
                 color=GRAY)
    return s


# ============ 1 封面 ============
s = prs.slides.add_slide(BLANK)
bar(s, 0, 0, 13.333, 7.5, NAVY)
# 装饰：右侧浅色斜块
dec = s.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(8.55), Inches(0),
                         Inches(4.8), Inches(7.5))
dec.fill.solid()
dec.fill.fore_color.rgb = RGBColor(0x27, 0x44, 0x74)
dec.line.fill.background()
dec.shadow.inherit = False
_strip_style(dec)

bar(s, 1.0, 1.62, 0.9, 0.07, ORANGE)
add_text(s, 1.0, 1.95, 7.4, 0.5, "北师大版（2019）选择性必修第二册　第二章 导数及其应用",
         size=16, color=RGBColor(0xC5, 0xD2, 0xE8))
add_text(s, 1.0, 2.55, 8.2, 1.1, "函数的极值", size=56, bold=True, color=WHITE)
add_text(s, 1.0, 3.75, 8.2, 0.5, "§6.2　函数的极值　（第 1 课时）", size=20,
         color=RGBColor(0xC5, 0xD2, 0xE8))
bar(s, 1.0, 4.55, 5.2, 0.012, RGBColor(0x4A, 0x63, 0x8E))
add_text(s, 1.0, 4.85, 4.0, 0.4, "授课人：　　　　　", size=15,
         color=RGBColor(0xB8, 0xC6, 0xDE))
add_text(s, 1.0, 5.35, 4.0, 0.4, "学　校：　　　　　", size=15,
         color=RGBColor(0xB8, 0xC6, 0xDE))
# 右侧大字装饰
add_text(s, 9.15, 2.2, 3.8, 3.0,
         [("f′(x)", {"size": 54, "bold": True, "color": ORANGE, "align": PP_ALIGN.CENTER,
                     "line": 1.1}),
          ("极值", {"size": 30, "color": RGBColor(0xC5, 0xD2, 0xE8),
                    "align": PP_ALIGN.CENTER, "space_before": 10})],
         ),

# ============ 2 学习目标 ============
s = new_slide("学习目标", tag="第 1 课时")
items = [
    ("理解概念", "理解极大值、极小值与极值点的含义，能说明极值的“局部性”。"),
    ("掌握方法", "掌握用导数判定极值的方法，能按规范步骤求函数的极值。"),
    ("辨析关系", "辨析 f′(x₀)=0 与“x₀ 为极值点”的关系，能举反例说明。"),
]
for i, (h, d) in enumerate(items):
    top = 1.55 + i * 1.55
    card(s, 0.95, top, 11.4, 1.28)
    bar(s, 0.95, top, 0.08, 1.28, NAVY if i != 2 else ORANGE)
    add_text(s, 1.35, top + 0.16, 2.0, 0.4, h, size=20, bold=True, color=NAVY)
    add_text(s, 3.35, top + 0.24, 8.8, 0.85, d, size=17, color=TEXT)
add_text(s, 0.95, 6.35, 11.4, 0.4,
         "重点：极值的概念与求极值的方法　　难点：“局部性”的理解与“两侧变号”的必要性",
         size=14, color=GRAY)

# ============ 3 复习回顾 ============
s = new_slide("复习回顾：导数与单调性", tag="引入")
card(s, 0.95, 1.55, 6.0, 2.5, fill=WHITE, line_color=LINE)
add_text(s, 1.25, 1.72, 5.4, 0.4, "已学结论", size=15, bold=True, color=ORANGE)
add_text(s, 1.25, 2.15, 5.4, 1.8,
         [("在区间 (a, b) 内：", {"size": 17, "color": GRAY}),
          ("f′(x) > 0  ⟹  f(x) 单调递增", {"size": 20, "color": NAVY, "bold": True,
                                            "space_before": 10}),
          ("f′(x) < 0  ⟹  f(x) 单调递减", {"size": 20, "color": NAVY, "bold": True,
                                            "space_before": 8})])

card(s, 7.15, 1.55, 5.2, 2.5, fill=WHITE, line_color=LINE)
add_text(s, 7.45, 1.72, 4.6, 0.4, "热身：f(x) = x³-3x", size=15, bold=True,
         color=ORANGE)
add_text(s, 7.45, 2.15, 4.6, 1.8,
         [("f′(x) = 3x² - 3 = 3(x+1)(x-1)", {"size": 19, "color": NAVY, "bold": True}),
          ("增区间：(-∞,-1)，(1,+∞)", {"size": 17, "color": TEXT, "space_before": 10}),
          ("减区间：(-1, 1)", {"size": 17, "color": TEXT, "space_before": 6})])

card(s, 0.95, 4.35, 11.4, 1.55, fill=RGBColor(0xFD, 0xF3, 0xEA))
add_text(s, 1.25, 4.55, 11.0, 0.4, "想一想", size=15, bold=True, color=ORANGE)
add_text(s, 1.25, 4.95, 10.8, 0.8,
         "函数在 x = -1 处由“增”转“减”，在 x = 1 处由“减”转“增”。"
         "这两个“转折点”附近的函数值，有什么特殊之处？",
         size=19, color=TEXT)

# ============ 4 情境探究 ============
s = new_slide("观察图象：寻找“增减转换点”", tag="引入")
s.shapes.add_picture(FIG1, Inches(0.85), Inches(1.45), width=Inches(6.3))
qs = [
    ("问题1", "求 f′(x)，说出 f(x) 的单调区间。"),
    ("问题2", "图象在 x = -1、x = 1 处发生了什么变化？"),
    ("问题3", "这两个点附近的函数值，处于什么地位？"),
]
for i, (q, d) in enumerate(qs):
    top = 1.55 + i * 1.62
    card(s, 7.5, top, 5.0, 1.35)
    add_text(s, 7.78, top + 0.14, 1.2, 0.35, q, size=14, bold=True, color=ORANGE)
    add_text(s, 7.78, top + 0.55, 4.4, 0.75, d, size=16, color=TEXT)
add_text(s, 7.5, 6.45, 5.0, 0.4, "提示：用“附近”“局部”来描述", size=13, color=GRAY)

# ============ 5 概念生成 ============
s = new_slide("极大值与极小值", tag="概念")
add_text(s, 0.95, 1.42, 11.4, 0.45,
         "设函数 y = f(x) 在点 x₀ 附近有定义：", size=19, color=TEXT)

card(s, 0.95, 2.0, 5.6, 2.45, fill=WHITE, line_color=LINE)
bar(s, 0.95, 2.0, 5.6, 0.08, ORANGE)
add_text(s, 1.25, 2.2, 5.0, 0.4, "极大值", size=20, bold=True, color=ORANGE)
add_text(s, 1.25, 2.62, 5.0, 1.5,
         [("若对 x₀ 附近的所有 x，", {"size": 17}),
          ("都有 f(x) < f(x₀)，", {"size": 17, "space_before": 5}),
          ("则称 f(x₀) 为极大值，", {"size": 17, "space_before": 5}),
          ("x₀ 为极大值点。", {"size": 17, "space_before": 5})])

card(s, 6.75, 2.0, 5.6, 2.45, fill=WHITE, line_color=LINE)
bar(s, 6.75, 2.0, 5.6, 0.08, NAVY_MID)
add_text(s, 7.05, 2.2, 5.0, 0.4, "极小值", size=20, bold=True, color=NAVY_MID)
add_text(s, 7.05, 2.62, 5.0, 1.5,
         [("若对 x₀ 附近的所有 x，", {"size": 17}),
          ("都有 f(x) > f(x₀)，", {"size": 17, "space_before": 5}),
          ("则称 f(x₀) 为极小值，", {"size": 17, "space_before": 5}),
          ("x₀ 为极小值点。", {"size": 17, "space_before": 5})])

card(s, 0.95, 4.6, 11.4, 1.15, fill=RGBColor(0xFD, 0xF3, 0xEA))
add_text(s, 1.25, 4.82, 10.8, 0.7,
         "极大值与极小值统称为极值；极大值点与极小值点统称为极值点。",
         size=19, bold=True, color=NAVY)
add_text(s, 0.95, 5.95, 11.4, 0.5,
         "本质：在 x₀ 的“附近”作比较 —— 极值刻画的是局部特征。",
         size=16, color=GRAY)

# ============ 6 概念辨析 ============
s = new_slide("辨析三个关键词", tag="概念")
keys = [
    ("“附近”", "极值是局部性质，只与 x₀ 附近的函数值比较，不与整个定义域比较。",
     "极大值 ≠ 最大值"),
    ("极值点 / 极值", "极值点 x₀ 是自变量的取值；极值 f(x₀) 是函数值。二者不是同一个对象。",
     "x₀  ↔  f(x₀)"),
    ("内部点", "极值点必须在定义域内部取得，区间端点处不谈极值。",
     "端点不算极值点"),
]
for i, (k, d, tip) in enumerate(keys):
    left = 0.95 + (i % 3) * 3.86
    card(s, left, 1.55, 3.62, 3.4, fill=WHITE, line_color=LINE)
    add_text(s, left + 0.28, 1.75, 3.1, 0.5, k, size=21, bold=True, color=NAVY)
    bar(s, left + 0.3, 2.35, 0.5, 0.05, ORANGE)
    add_text(s, left + 0.28, 2.55, 3.1, 1.5, d, size=15, color=TEXT, line=1.45)
    add_text(s, left + 0.28, 4.28, 3.1, 0.4, tip, size=13, bold=True, color=ORANGE)

add_text(s, 0.95, 5.35, 11.4, 1.0,
         [("追问：极大值一定是函数的最大值吗？", {"size": 19, "bold": True, "color": NAVY}),
          ("例：f(x)=x³-3x 的极大值 f(-1)=2，但当 x→+∞ 时 f(x)→+∞，函数没有最大值。",
           {"size": 16, "color": GRAY, "space_before": 8})])

# ============ 7 判定方法 ============
s = new_slide("怎样判定极值？", tag="判定")
add_text(s, 0.95, 1.42, 11.4, 0.45,
         "前提：f(x) 在 x₀ 处可导，且 f′(x₀) = 0。观察 f′(x) 在 x₀ 两侧的符号：",
         size=18, color=TEXT)

rows = [
    ("左正右负", "x₀ 左侧 f′(x) > 0，右侧 f′(x) < 0", "f(x₀) 为极大值", ORANGE),
    ("左负右正", "x₀ 左侧 f′(x) < 0，右侧 f′(x) > 0", "f(x₀) 为极小值", NAVY_MID),
    ("两侧同号", "x₀ 两侧 f′(x) 符号相同", "f(x₀) 不是极值", GRAY),
]
for i, (tag, cond, res, color) in enumerate(rows):
    top = 1.98 + i * 1.36
    card(s, 0.95, top, 11.4, 1.15, fill=WHITE, line_color=LINE)
    bar(s, 0.95, top, 0.09, 1.15, color)
    add_text(s, 1.35, top + 0.28, 2.0, 0.5, tag, size=20, bold=True, color=color)
    add_text(s, 3.55, top + 0.32, 5.3, 0.5, cond, size=16, color=TEXT)
    add_text(s, 9.0, top + 0.28, 3.2, 0.5, res, size=19, bold=True, color=color,
             align=PP_ALIGN.RIGHT)

card(s, 0.95, 6.0, 11.4, 0.68, fill=RGBColor(0xEE, 0xF2, 0xFA))
add_text(s, 1.2, 6.12, 11.0, 0.4,
         "口诀：左正右负取极大，左负右正取极小；符号不变，极值不现。",
         size=16, bold=True, color=NAVY)

# ============ 8 反例辨析 ============
s = new_slide("辨析：f′(x₀)=0 就一定是极值点吗？", tag="辨析")
s.shapes.add_picture(FIG2, Inches(0.95), Inches(1.5), width=Inches(5.4))
card(s, 6.85, 1.55, 5.5, 2.0, fill=RGBColor(0xFD, 0xF3, 0xEA))
add_text(s, 7.15, 1.72, 5.0, 0.4, "反例", size=15, bold=True, color=ORANGE)
add_text(s, 7.15, 2.12, 5.0, 1.3,
         [("f(x) = x³，f′(x) = 3x²", {"size": 19, "color": NAVY, "bold": True}),
          ("f′(0) = 0，但 f(x) 在 R 上单调递增，", {"size": 16, "space_before": 8}),
          ("所以 x = 0 不是极值点。", {"size": 16, "space_before": 4})])

card(s, 6.85, 3.72, 5.5, 1.45, fill=WHITE, line_color=LINE)
add_text(s, 7.15, 3.9, 5.0, 0.35, "结论", size=14, bold=True, color=NAVY)
add_text(s, 7.15, 4.3, 5.0, 0.6,
         "f′(x₀)=0 是极值存在的必要条件，不是充分条件。",
         size=17, bold=True, color=NAVY)

add_text(s, 6.85, 5.45, 5.5, 1.2,
         "因此，求极值时不能只令 f′(x)=0 就下结论，必须检验两侧的符号是否变号。",
         size=16, color=GRAY, line=1.5)

# ============ 9 例1 分析 ============
s = new_slide("例1　求函数的极值", tag="例题")
card(s, 0.95, 1.42, 11.4, 0.85, fill=RGBColor(0xEE, 0xF2, 0xFA))
add_text(s, 1.25, 1.55, 10.9, 0.55, "求函数 f(x) = x³ - 3x 的极值。",
         size=24, bold=True, color=NAVY)

steps = [
    ("① 分析题意", "定义域为 R。求极值即找 f′(x)=0 的点，并判断其两侧 f′(x) 的符号。"),
    ("② 拟定计划", "求导 → 解方程 f′(x)=0 → 列表分析符号 → 按判定写出极值。"),
]
for i, (h, d) in enumerate(steps):
    top = 2.65 + i * 1.35
    card(s, 0.95, top, 11.4, 1.12, fill=WHITE, line_color=LINE)
    add_text(s, 1.25, top + 0.2, 2.3, 0.5, h, size=18, bold=True, color=ORANGE)
    add_text(s, 3.5, top + 0.22, 8.6, 0.75, d, size=16, color=TEXT)

add_text(s, 0.95, 5.55, 11.4, 1.2,
         [("提示：把 f′(x) 因式分解，符号由各因式决定 ——", {"size": 15, "color": GRAY}),
          ("f′(x) = 3x² - 3 = 3(x + 1)(x - 1)", {"size": 20, "color": NAVY,
                                                 "bold": True, "space_before": 6})])

# ============ 10 例1 实施（符号表） ============
s = new_slide("例1　实施计划", tag="例题")
add_text(s, 0.95, 1.4, 11.4, 0.5,
         "f′(x) = 3x² - 3 = 3(x + 1)(x - 1)，令 f′(x) = 0，得 x₁ = -1，x₂ = 1。",
         size=17, color=TEXT)

cols, rows_n = 6, 3
tbl_shape = s.shapes.add_table(rows_n, cols, Inches(0.95), Inches(2.05),
                               Inches(11.4), Inches(2.4))
table = tbl_shape.table
headers = ["x", "(-∞, -1)", "-1", "(-1, 1)", "1", "(1, +∞)"]
data = [
    ["f′(x)", "+", "0", "-", "0", "+"],
    ["f(x)", "↗", "极大值", "↘", "极小值", "↗"],
]
widths = [1.55, 2.05, 1.45, 2.0, 1.45, 2.05]
for j, w in enumerate(widths):
    table.columns[j].width = Inches(w)

for i in range(rows_n):
    table.rows[i].height = Inches(0.8)

cell_data = [headers] + data
for i, rowvals in enumerate(cell_data):
    for j, val in enumerate(rowvals):
        cell = table.cell(i, j)
        cell.text = ""
        cell.vertical_anchor = MSO_ANCHOR.MIDDLE
        cell.margin_left = Inches(0.05)
        cell.margin_right = Inches(0.05)
        cell.margin_top = Inches(0.02)
        cell.margin_bottom = Inches(0.02)
        p = cell.text_frame.paragraphs[0]
        p.alignment = PP_ALIGN.CENTER
        r = p.add_run()
        r.text = val
        r.font.size = Pt(19 if i == 0 else 20)
        r.font.bold = True
        r.font.name = EN
        rPr = r._r.get_or_add_rPr()
        ea = rPr.makeelement(qn("a:ea"), {})
        ea.set("typeface", CN)
        rPr.append(ea)
        if i == 0:
            r.font.color.rgb = WHITE
            cell.fill.solid()
            cell.fill.fore_color.rgb = NAVY
        else:
            cell.fill.solid()
            cell.fill.fore_color.rgb = WHITE if i % 2 else LIGHT
            if (i, j) in [(1, 2), (1, 4)]:
                r.font.color.rgb = ORANGE
            elif val in ("极大值", "极小值"):
                r.font.color.rgb = ORANGE if val == "极大值" else NAVY_MID
                r.font.size = Pt(16)
            elif val in ("↗", "↘", "+", "-"):
                r.font.color.rgb = NAVY
            else:
                r.font.color.rgb = TEXT

card(s, 0.95, 4.58, 5.5, 1.5, fill=WHITE, line_color=LINE)
add_text(s, 1.25, 4.75, 5.0, 1.2,
         [("极大值", {"size": 15, "bold": True, "color": ORANGE}),
          ("f(-1) = (-1)³ - 3×(-1) = 2", {"size": 19, "bold": True,
                                          "space_before": 6})])
card(s, 6.85, 4.58, 5.5, 1.5, fill=WHITE, line_color=LINE)
add_text(s, 7.15, 4.75, 5.0, 1.2,
         [("极小值", {"size": 15, "bold": True, "color": NAVY_MID}),
          ("f(1) = 1³ - 3×1 = -2", {"size": 19, "bold": True,
                                    "space_before": 6})])
add_text(s, 0.95, 6.18, 11.4, 0.4,
         "∴ f(x) 的极大值为 2，极小值为 -2。", size=17, bold=True, color=NAVY)

# ============ 11 例1 回顾反思 ============
s = new_slide("例1　回顾反思", tag="例题")
items = [
    ("极大值 2 是最大值吗？",
     "不是。x→+∞ 时 f(x)→+∞，函数无最大值 —— 极值只是“局部最大”。"),
    ("结果可信吗？",
     "符号表显示“先增后减再增”，与图象形状完全一致，数形互相印证。"),
    ("能省略表格吗？",
     "不能。只令 f′(x)=0 就可能把 f(x)=x³ 中的 x=0 误判为极值点。"),
]
for i, (q, a) in enumerate(items):
    top = 1.5 + i * 1.6
    card(s, 0.95, top, 11.4, 1.35, fill=WHITE, line_color=LINE)
    bar(s, 0.95, top, 0.09, 1.35, ORANGE if i == 2 else NAVY_MID)
    add_text(s, 1.35, top + 0.16, 10.5, 0.4, q, size=18, bold=True, color=NAVY)
    add_text(s, 1.35, top + 0.6, 10.5, 0.7, a, size=15, color=GRAY, line=1.4)

card(s, 0.95, 6.2, 11.4, 0.58, fill=RGBColor(0xFD, 0xF3, 0xEA))
add_text(s, 1.2, 6.3, 11.0, 0.4,
         "求极值四步：分析题意 → 拟定计划 → 实施计划 → 回顾反思",
         size=15, bold=True, color=NAVY)

# ============ 12 课堂小结 ============
s = new_slide("课堂小结", tag="小结")
steps = [
    ("图象特征", "增减转换\n转折点"),
    ("极值概念", "局部最大\n局部最小"),
    ("判定方法", "f′(x₀)=0\n且两侧变号"),
    ("求极值四步", "求导·列表\n判号·写值"),
]
for i, (h, d) in enumerate(steps):
    left = 0.85 + i * 3.2
    card(s, left, 1.75, 2.55, 2.05, fill=WHITE, line_color=NAVY if i < 3 else ORANGE)
    add_text(s, left, 1.95, 2.55, 0.45, h, size=18, bold=True, color=NAVY,
             align=PP_ALIGN.CENTER)
    add_text(s, left, 2.5, 2.55, 1.2,
             [(ln, {"size": 14, "color": GRAY, "align": PP_ALIGN.CENTER,
                    "line": 1.5}) for ln in d.split("\n")])
    if i < 3:
        add_text(s, left + 2.6, 2.62, 0.55, 0.5, "→", size=24, bold=True,
                 color=ORANGE, align=PP_ALIGN.CENTER)

card(s, 0.95, 4.35, 5.6, 1.65, fill=RGBColor(0xFD, 0xF3, 0xEA))
add_text(s, 1.2, 4.52, 5.1, 0.4, "两个提醒", size=15, bold=True, color=ORANGE)
add_text(s, 1.2, 4.92, 5.1, 1.0,
         [("① 极值是局部概念，不是最大值；", {"size": 16}),
          ("② f′(x₀)=0 未必是极值点。", {"size": 16, "space_before": 5})])

card(s, 6.85, 4.35, 5.5, 1.65, fill=RGBColor(0xEE, 0xF2, 0xFA))
add_text(s, 7.1, 4.52, 5.0, 0.4, "研究方法", size=15, bold=True, color=NAVY)
add_text(s, 7.1, 4.92, 5.0, 1.0,
         "由形到数、由特殊到一般、用反例检验结论的边界。", size=16, color=TEXT,
         line=1.5)

# ============ 13 作业 ============
s = new_slide("作业布置", tag="作业")
hw = [
    ("基础", "教材 §6.2 练习第 1、2 题。"),
    ("巩固", "求函数 f(x) = x³ - 3x² 的极值，仿照例1写出完整过程。"),
    ("拓展", "思考：若 f′(x₀)=0，且在 x₀ 两侧 f′(x) 同号，函数图象在 x₀ 处有什么特点？"),
]
for i, (t, d) in enumerate(hw):
    top = 1.65 + i * 1.55
    card(s, 0.95, top, 11.4, 1.28, fill=WHITE, line_color=LINE)
    bar(s, 0.95, top, 0.09, 1.28, ORANGE if i == 2 else NAVY_MID)
    add_text(s, 1.35, top + 0.4, 1.3, 0.45, t, size=18, bold=True, color=NAVY)
    add_text(s, 2.85, top + 0.3, 9.2, 0.8, d, size=16, color=TEXT, line=1.4)

add_text(s, 0.95, 6.35, 11.4, 0.6,
         "下节课：函数的最大（小）值 —— 从“局部”走向“整体”。",
         size=16, bold=True, color=GRAY)

prs.save(OUT)
print("saved:", OUT, "slides:", len(prs.slides._sldIdLst))
