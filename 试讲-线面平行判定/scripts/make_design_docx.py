# -*- coding: utf-8 -*-
"""生成《直线与平面平行的判定》教学设计（严格按学院模板结构）"""
from docx import Document
from docx.shared import Pt, Cm, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_LINE_SPACING
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

OUT = "/workspace/试讲-线面平行判定/教学设计-直线与平面平行的判定.docx"

doc = Document()

# ---------- 全局 ----------
sec = doc.sections[0]
sec.page_width = Cm(21.0)
sec.page_height = Cm(29.7)
sec.left_margin = Cm(2.4)
sec.right_margin = Cm(2.4)
sec.top_margin = Cm(2.2)
sec.bottom_margin = Cm(2.2)

style = doc.styles["Normal"]
style.font.name = "Times New Roman"
style.font.size = Pt(11)
style.element.rPr.rFonts.set(qn("w:eastAsia"), "宋体")
style.paragraph_format.line_spacing_rule = WD_LINE_SPACING.ONE_POINT_FIVE


def set_font(run, name_cn="宋体", name_en="Times New Roman", size=11,
             bold=False, color=None):
    run.font.name = name_en
    run.font.size = Pt(size)
    run.bold = bold
    run._element.rPr.rFonts.set(qn("w:eastAsia"), name_cn)
    if color:
        run.font.color.rgb = RGBColor(*color)


def para(text, size=11, bold=False, cn="宋体", en="Times New Roman",
         align=WD_ALIGN_PARAGRAPH.JUSTIFY, indent_first=0, space_before=0,
         space_after=0, line=1.5):
    p = doc.add_paragraph()
    p.alignment = align
    pf = p.paragraph_format
    pf.first_line_indent = Pt(indent_first)
    pf.space_before = Pt(space_before)
    pf.space_after = Pt(space_after)
    pf.line_spacing = line
    r = p.add_run(text)
    set_font(r, cn, en, size, bold)
    return p


def heading(num_text, size=12):
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(8)
    p.paragraph_format.space_after = Pt(3)
    p.paragraph_format.line_spacing = 1.3
    r = p.add_run(num_text)
    set_font(r, "黑体", "Times New Roman", size, True)
    return p


def sub_heading(text, size=11):
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(5)
    p.paragraph_format.space_after = Pt(2)
    p.paragraph_format.line_spacing = 1.4
    r = p.add_run(text)
    set_font(r, "黑体", "Times New Roman", size, True)
    return p


def shade(cell, hexcolor):
    tcPr = cell._tc.get_or_add_tcPr()
    shd = OxmlElement("w:shd")
    shd.set(qn("w:val"), "clear")
    shd.set(qn("w:fill"), hexcolor)
    tcPr.append(shd)


def cell_text(cell, text, size=10.5, bold=False, cn="宋体",
              align=WD_ALIGN_PARAGRAPH.JUSTIFY, line=1.35):
    cell.text = ""
    p = cell.paragraphs[0]
    p.alignment = align
    p.paragraph_format.line_spacing = line
    p.paragraph_format.space_before = Pt(1)
    p.paragraph_format.space_after = Pt(1)
    for i, seg in enumerate(text.split("\n")):
        if i > 0:
            p = cell.add_paragraph()
            p.alignment = align
            p.paragraph_format.line_spacing = line
            p.paragraph_format.space_before = Pt(0)
            p.paragraph_format.space_after = Pt(1)
        r = p.add_run(seg)
        set_font(r, cn, "Times New Roman", size, bold)


def set_col_widths(table, widths_cm):
    """真正固定列宽：设置 tblW、tblLayout=fixed、tblGrid 与每个 tcW"""
    table.autofit = False
    tbl = table._tbl
    tblPr = tbl.tblPr

    # 1) tblW
    for old in tblPr.findall(qn("w:tblW")):
        tblPr.remove(old)
    tblW = OxmlElement("w:tblW")
    tblW.set(qn("w:type"), "dxa")
    tblW.set(qn("w:w"), str(int(sum(widths_cm) * 567)))
    tblPr.append(tblW)

    # 2) tblLayout = fixed
    for old in tblPr.findall(qn("w:tblLayout")):
        tblPr.remove(old)
    layout = OxmlElement("w:tblLayout")
    layout.set(qn("w:type"), "fixed")
    tblPr.append(layout)

    # 3) tblGrid
    grid = tbl.find(qn("w:tblGrid"))
    if grid is not None:
        tbl.remove(grid)
    grid = OxmlElement("w:tblGrid")
    for w in widths_cm:
        gc = OxmlElement("w:gridCol")
        gc.set(qn("w:w"), str(int(w * 567)))
        grid.append(gc)
    tbl.insert(list(tbl).index(tblPr) + 1, grid)

    # 4) 每个单元格宽度（按合并跨度累计，跳过重复的合并单元）
    for row in table.rows:
        idx = 0
        seen = set()
        for cell in row.cells:
            tc_id = id(cell._tc)
            if tc_id in seen:
                continue
            seen.add(tc_id)
            span = cell.grid_span if hasattr(cell, "grid_span") else 1
            w = sum(widths_cm[idx: idx + span])
            if w:
                cell.width = Cm(w)
            idx += span

    # 5) 行不跨页拆分
    for row in table.rows:
        trPr = row._tr.get_or_add_trPr()
        if trPr.find(qn("w:cantSplit")) is None:
            el = OxmlElement("w:cantSplit")
            trPr.append(el)


def repeat_header(table):
    """表头跨页重复"""
    tr = table.rows[0]._tr
    trPr = tr.get_or_add_trPr()
    el = OxmlElement("w:tblHeader")
    el.set(qn("w:val"), "true")
    trPr.append(el)


def set_cell_margins(table, left=0.05, right=0.05, top=0.02, bottom=0.02):
    """设置整表单元格内边距（cm）"""
    tblPr = table._tbl.tblPr
    for old in tblPr.findall(qn("w:tblCellMar")):
        tblPr.remove(old)
    mar = OxmlElement("w:tblCellMar")
    for tag, v in (("w:top", top), ("w:left", left),
                   ("w:bottom", bottom), ("w:right", right)):
        e = OxmlElement(tag)
        e.set(qn("w:w"), str(int(v * 567)))
        e.set(qn("w:type"), "dxa")
        mar.append(e)
    tblPr.append(mar)


# ---------- 封面信息表 ----------
title = doc.add_paragraph()
title.alignment = WD_ALIGN_PARAGRAPH.CENTER
title.paragraph_format.space_after = Pt(10)
r = title.add_run("2024级数学专业师范生技能训练教学设计")
set_font(r, "黑体", "Times New Roman", 16, True)

info = doc.add_table(rows=3, cols=6)
info.style = "Table Grid"
info.alignment = WD_TABLE_ALIGNMENT.CENTER
rows_data = [
    ["姓名", "", "学段", "高中", "授课班级", ""],
    ["职称", "", "学科", "数学", "授课时间", ""],
]
for i, row in enumerate(rows_data):
    for j, val in enumerate(row):
        cell_text(info.rows[i].cells[j], val, size=10.5,
                  bold=(j % 2 == 0), align=WD_ALIGN_PARAGRAPH.CENTER)
# 课题行：先设列宽，再合并第 2~6 格
set_col_widths(info, [2.0, 3.0, 1.8, 2.6, 2.6, 4.2])
row3 = info.rows[2]
cell_text(row3.cells[0], "课题", size=10.5, bold=True,
          align=WD_ALIGN_PARAGRAPH.CENTER)
merged = row3.cells[1]
merged = merged.merge(row3.cells[5])
cell_text(merged,
          "§8.5.2 直线与平面平行的判定（人教A版必修第二册 第八章 立体几何初步）",
          size=10.5, align=WD_ALIGN_PARAGRAPH.LEFT)
merged.width = Cm(14.2)

doc.add_paragraph().paragraph_format.space_after = Pt(0)

# ---------- 1. 教材分析 ----------
heading("1.教材分析")
sub_heading("1.1 教材地位分析")
para("本节课选自人教A版（2019）高中数学必修第二册第八章《立体几何初步》"
     "§8.5.2 直线与平面平行（第1课时：直线与平面平行的判定，教材第134—136页）。"
     "从知识序列看，"
     "学生已在本章前面学习了基本立体图形、直观图，并在 §8.4 中借助长方体认识了空间中"
     "直线与平面的三种位置关系（直线在平面内、直线与平面相交、直线与平面平行），"
     "在 §8.5.1 中又学习了基本事实4（平行线的传递性）与等角定理，具备证明线线平行的工具；"
     "本节之后，教材将学习直线与平面平行的性质定理，以及平面与平面平行的判定与性质。"
     "因此，本节既是线线平行研究的延伸，又是面面平行研究的基础，在“线线—线面—面面”的"
     "平行关系体系中处于承上启下的枢纽位置。", indent_first=22)
para("从育人价值看，线面平行的定义要求“直线与平面没有公共点”，而直线无限延伸、平面无限延展，"
     "这一条件在操作上无法直接验证。判定定理的建立过程，本质上是寻找一个可操作的等价条件，"
     "把空间问题转化为平面问题、把线面关系转化为线线关系，是渗透转化与化归思想、"
     "发展直观想象与逻辑推理素养的典型素材。", indent_first=22)

sub_heading("1.2 教材内容分析")
para("本节教材的编排路径是：先追问“怎样判定直线与平面平行”，指出依定义需证明“没有公共点”"
     "却难以操作；再给出两个现实情境——门扇绕一边转动、矩形硬纸板绕一边在桌面上转动，"
     "引导学生直观感知“平面外的一条直线与平面内的一条直线平行”这一共同特征；"
     "在此基础上归纳出直线与平面平行的判定定理，并用文字语言、图形语言、符号语言三种形式"
     "加以表述；最后通过例2（空间四边形相邻两边中点连线平行于经过另外两边的平面）"
     "示范定理的应用。教材以长方体为基本载体，体现了“直观感知—操作确认—推理论证”的"
     "立体几何研究方法。", indent_first=22)
para("教学重点：直线与平面平行的判定定理及其应用。", indent_first=22)
para("教学难点：判定定理中“平面外”“平面内”两个条件的必要性；由“线线平行”推出"
     "“线面平行”这一转化意识的建立。", indent_first=22)

# ---------- 2. 教学目标 ----------
heading("2.教学目标")
sub_heading("2.1 基础知识与基本技能目标")
para("① 理解并掌握直线与平面平行的判定定理，能用文字语言、图形语言、符号语言三种形式"
     "准确表述该定理；", indent_first=22)
para("② 能在简单几何体（空间四边形、正方体等）中运用判定定理证明直线与平面平行，"
     "书写规范、条件完整；", indent_first=22)
para("③ 能说明定理中“平面外”“平面内”两个条件缺一不可，并举出反例加以辨析。", indent_first=22)

sub_heading("2.2 数学思想目标")
para("通过“情境观察—归纳猜想—确认定理”的过程，渗透由特殊到一般、由直观到严谨的认识方法；"
     "通过把“线面平行”转化为“线线平行”，渗透转化与化归思想，体会空间问题平面化的基本策略。",
     indent_first=22)

sub_heading("2.3 数学活动经验目标")
para("经历“定义难以操作—寻找可操作判据—在情境中发现特征—归纳定理—辨析条件—应用定理”"
     "的完整命题学习过程，积累观察、猜想、说理与规范表达的活动经验，提升空间想象与"
     "交流表达能力。", indent_first=22)

sub_heading("2.4 核心素养目标")
para("① 直观想象：能从门扇、纸板等实物与立体图形中识别线面位置关系，把实物特征抽象为"
     "几何图形；", indent_first=22)
para("② 逻辑推理：理解判定定理的条件结构，能按“找线—证平行—写条件—得结论”的逻辑链"
     "完成证明；", indent_first=22)
para("③ 数学抽象：实现文字、图形、符号三种语言的相互转译，把握“线线平行⇒线面平行”"
     "的本质。", indent_first=22)

# ---------- 3. 学情分析 ----------
heading("3.学情分析")
para("从认知基础看，学生已在 §8.4 中掌握空间中直线与平面的三种位置关系，知道线面平行的"
     "定义是“直线与平面没有公共点”；在 §8.5.1 中学习了基本事实4 和等角定理，"
     "并在初中平面几何中积累了用三角形中位线、平行四边形证明线线平行的经验。"
     "这些都为判定定理的归纳与应用提供了支撑。", indent_first=22)
para("可能存在的困难有三点：一是缺乏“换一个可操作标准”的意识，面对“证明没有公共点”"
     "往往无从下手，需要教师明确指出定义在操作上的局限；二是空间想象能力尚在发展，"
     "在立体图中不易发现“平面内的那条线”，尤其不会主动添加辅助线（如连接 BD）"
     "把隐藏的平行关系显现出来；三是书写时容易漏写“a 不在平面 α 内”或“b 在平面 α 内”，"
     "只凭 a∥b 就下结论，对定理条件的完备性缺乏体会。教学中应通过情境对比、反例辨析"
     "与板书示范逐一突破。", indent_first=22)

# ---------- 4. 教学策略分析 ----------
heading("4.教学策略分析")
para("遵循“以学生为主体，教师是数学课堂活动的组织者、引导者与合作者”的教学原则，"
     "采用情境教学与启发式讲授相结合的方法，以问题链贯穿全课：先暴露定义的不可操作性，"
     "产生寻找判据的内在需求；再用两个生活情境提供共同特征的直观基础；然后归纳、确认定理"
     "并辨析条件；最后在例题中把定理用起来。学生主要采用观察发现、合情推理与合作交流的"
     "学习方式，在“最近发展区”内完成命题的自主建构。", indent_first=22)
para("教学手段上，用矩形硬纸板现场演示转动过程，用教室门扇作实物参照，配合课件呈现立体图形，"
     "增强直观性；同时用板书完整呈现定理的三种语言表述与例题的规范解题过程，"
     "做到“实物与图形辅助直观、板书承载严谨”。", indent_first=22)

# ---------- 5. 教学过程 ----------
p5_head = heading("5.教学过程")
p5_head.paragraph_format.keep_with_next = True

proc = doc.add_table(rows=1, cols=3)
proc.style = "Table Grid"
proc.alignment = WD_TABLE_ALIGNMENT.CENTER
hdr = proc.rows[0].cells
for i, t in enumerate(["教学内容与问题设计", "教师行为", "学生活动"]):
    cell_text(hdr[i], t, size=10.5, bold=True, cn="黑体",
              align=WD_ALIGN_PARAGRAPH.CENTER)
    shade(hdr[i], "DCE6F1")


def add_teaching_row(content, teacher, student):
    row = proc.add_row()
    cell_text(row.cells[0], content, size=10, align=WD_ALIGN_PARAGRAPH.LEFT)
    cell_text(row.cells[1], teacher, size=10, align=WD_ALIGN_PARAGRAPH.LEFT)
    cell_text(row.cells[2], student, size=10, align=WD_ALIGN_PARAGRAPH.LEFT)
    return row


def add_intent_row(text):
    row = proc.add_row()
    merged = row.cells[0].merge(row.cells[1]).merge(row.cells[2])
    cell_text(merged, text, size=10, cn="楷体", align=WD_ALIGN_PARAGRAPH.LEFT)
    shade(merged, "F2F2F2")


add_teaching_row(
    "5.1 创设情境（约3分钟）\n"
    "复习：空间中直线与平面有哪几种位置关系？\n"
    "问题1：怎样判定一条直线与一个平面平行？按定义需要证明什么？\n"
    "问题2：直线无限延伸、平面无限延展，“没有公共点”能直接验证吗？\n"
    "情境1：教室的门扇绕固定的一边转动，转动的这一边与墙面有公共点吗？它与墙面平行吗？\n"
    "情境2：把矩形硬纸板 ABCD 平放在桌面上，绕边 DC 转动（AB 离开桌面），"
    "AB 与桌面有公共点吗？AB 与 DC 是什么关系？DC 在哪里？",
    "先请学生口答三种位置关系，板书“平行＝无公共点”；"
    "用问题2把学生逼到“定义不可操作”的困境，指出需要寻找新的判据；"
    "现场转动门扇与硬纸板，引导学生把注意力放在“转动的边”与“固定的边”的关系上。",
    "回答：直线在平面内、直线与平面相交、直线与平面平行；"
    "意识到“没有公共点”无法逐一验证；"
    "观察演示后回答：转动的边与固定的边平行，固定的边在墙面（桌面）内，"
    "转动的边不在墙面（桌面）内，且与墙面（桌面）平行。")
add_intent_row(
    "设计意图：以定义的不可操作性制造认知冲突，使判定定理的产生具有内在需求，而不是"
    "被动接受一个现成结论；两个情境都指向“平面外一条直线与平面内一条直线平行”这一"
    "共同结构，为随后的归纳提供充足的直观素材；现场演示有助于发展直观想象素养，"
    "并让学生自己说出“在平面内”“不在平面内”这两个关键条件，为条件辨析埋下伏笔。")

add_teaching_row(
    "5.2 新知探究（约6分钟）\n"
    "（1）归纳猜想\n"
    "两个情境的共同点：平面外的一条直线，与平面内的一条直线平行。\n"
    "（2）确认定理\n"
    "定理：平面外一条直线与此平面内的一条直线平行，则该直线与此平面平行。\n"
    "符号语言：a⊄α，b⊂α，且 a∥b ⇒ a∥α（a 在 α 外，b 在 α 内）。\n"
    "图形语言：画平面 α 与直线 a、b，标出平行记号。\n"
    "（3）条件辨析\n"
    "问题3：去掉“a⊄α”，结论还成立吗？\n"
    "问题4：去掉“b⊂α”，b 还能起到“桥梁”作用吗？\n"
    "（4）转化价值\n"
    "线面平行（难验证） ← 线线平行（可用中位线、平行四边形证明）",
    "引导学生用自己的话概括共同特征，再逐字对照定理表述，强调“平面外”“平面内”"
    "两个限定；"
    "板书文字、符号、图形三种语言，并标注三者的一一对应；"
    "用反例回应问题3：若 a 也在平面 α 内，a∥b 时 a 与 α 并不平行；"
    "结合情境追问，明确三个条件缺一不可。",
    "归纳并表述共同特征；"
    "对照板书复述定理，完成三种语言的转译；"
    "回答问题3：不成立，此时直线在平面内；"
    "回答问题4：b 若不在平面内，就无法说明 a 与“这个平面”的关系。")
add_intent_row(
    "设计意图：命题课的核心不是记住结论，而是经历“猜想—确认—辨析”的完整过程。"
    "先由情境归纳，使定理具有来处；再用三种语言表述，落实数学抽象与规范表达；"
    "最后用反例辨析两个限定条件，直接指向学生书写中漏写条件的典型错误，"
    "使“条件完备”成为自觉要求。转化关系的板书，为下一环节的“找线”策略提供方法支撑。")

add_teaching_row(
    "5.3 例题讲解（约5分钟）\n"
    "例2 如图，空间四边形 ABCD 中，E，F 分别是 AB，AD 的中点。\n"
    "求证：EF∥平面 BCD。\n"
    "① 分析题意：要证线面平行，只需在平面 BCD 内找一条直线与 EF 平行；"
    "EF 是 △ABD 的中位线，自然想到 BD。\n"
    "② 拟定计划：连接 BD（把平面内的线“画出来”）→ 由中位线定理得 EF∥BD "
    "→ 依次写出三个条件 → 得结论。\n"
    "③ 实施计划：\n"
    "连接 BD。\n"
    "∵ E，F 分别是 AB，AD 的中点，∴ EF∥BD。\n"
    "又 EF⊄平面 BCD，BD⊂平面 BCD，\n"
    "∴ EF∥平面 BCD。\n"
    "④ 回顾反思：关键在于“找线”，常借中位线或平行四边形；"
    "书写时三个条件缺一不可；想一想还有没有别的线可用。",
    "按“分析—计划—实施—回顾”四步示范，边讲边完整板书；"
    "在“分析”处追问“在哪一个平面内找线、找哪一条线”，"
    "把思维过程显性化；"
    "在“回顾”处强调书写规范，并追问若取 BC 中点能否同样奏效。",
    "跟随分析，说出“要在平面 BCD 内找与 EF 平行的线”；"
    "独立写出中位线推理；"
    "补全“EF⊄平面 BCD，BD⊂平面 BCD”两个条件；"
    "讨论其他可能的辅助线。")
add_intent_row(
    "设计意图：以四步示范使解题过程结构化、可迁移，符合试讲对例题讲解“分析题意、"
    "拟定计划、实施计划、回顾反思”的要求；把“找线”这一隐性策略显性化，"
    "使学生获得可操作的方法而不是被动模仿；板书完整呈现三条件，"
    "强化“条件完备才能得结论”的严谨意识。")

add_teaching_row(
    "5.4 课堂小结（约1分钟）\n"
    "以框图梳理研究路径：\n"
    "定义（无公共点，难操作）→ 情境观察 → 归纳猜想 → 判定定理（三个条件）"
    "→ 应用（在平面内找平行线）。\n"
    "一句话概括：证线面平行，就在平面内找一条平行线。\n"
    "强调：判定定理的三个条件缺一不可。",
    "提出问题“这节课我们是怎样得到一个判定定理的”，"
    "与学生共同完成框图，板书关键节点。",
    "回顾学习过程，用一句话概括方法；"
    "补充易错点：漏写“平面外”“平面内”。")
add_intent_row(
    "设计意图：用框图呈现“定义—情境—猜想—定理—应用”的命题学习路径，"
    "帮助学生形成结构化的认识，并把“空间问题平面化”的转化策略提升为可迁移的数学活动经验。")

add_teaching_row(
    "5.5 作业设计\n"
    "① 教材 §8.5.2 练习第1、2题；\n"
    "② 正方体 ABCD-A₁B₁C₁D₁ 中，E，F 分别是 AB，AD 的中点，"
    "求证：EF∥平面 A₁B₁C₁D₁；\n"
    "③ 思考：若把判定定理中的“平面外”这一条件去掉，结论还成立吗？请举反例说明。",
    "布置作业并说明第②题的要求（辅助线要写清、三个条件要齐全）；"
    "第③题不作统一要求，供学有余力的学生思考。",
    "记录作业，明确书写规范。")
add_intent_row(
    "设计意图：作业分基础、巩固、拓展三个层次，兼顾不同水平的学生；第③题回扣条件辨析，"
    "使“条件完备”的认识在课后得到再一次强化。")

set_col_widths(proc, [8.2, 3.9, 4.1])
repeat_header(proc)


def allow_split(table):
    """教学内容行允许跨页，避免整行下移在页底留下大片空白"""
    for row in table.rows:
        cells = row.cells
        if cells[0]._tc is cells[1]._tc:      # 合并行（设计意图）不拆
            continue
        trPr = row._tr.find(qn("w:trPr"))
        if trPr is None:
            continue
        el = trPr.find(qn("w:cantSplit"))
        if el is not None:
            trPr.remove(el)


allow_split(proc)

# ---------- 6. 板书设计 ----------
h6 = heading("6.板书设计")
h6.paragraph_format.keep_with_next = True
p6_head = para("（黑板分三个区域，左侧定理与条件辨析，中间例题完整解答，右侧图形区）",
               size=10, cn="楷体", indent_first=22)
p6_head.paragraph_format.keep_with_next = True
board = doc.add_table(rows=1, cols=3)
board.style = "Table Grid"
board.alignment = WD_TABLE_ALIGNMENT.CENTER
b = board.rows[0].cells
cell_text(b[0],
          "直线与平面平行的判定\n"
          "一、定理\n"
          "平面外一条直线与此平面\n"
          "内的一条直线平行，则该\n"
          "直线与此平面平行。\n"
          "符号：a⊄α，b⊂α，a∥b\n"
          "　　　⇒ a∥α\n"
          "二、三个条件（缺一不可）\n"
          "① a 在平面 α 外\n"
          "② b 在平面 α 内\n"
          "③ a∥b\n"
          "三、转化\n"
          "线面平行 ← 线线平行\n"
          "（找中位线、平行四边形）", size=9,
          align=WD_ALIGN_PARAGRAPH.LEFT, line=1.2)

# 中间：例题完整解答
cell_text(b[1],
          "例2 空间四边形 ABCD 中，E，F 分别是\n"
          "AB，AD 的中点。求证：EF∥平面 BCD。\n\n"
          "证明：连接 BD。\n"
          "∵ E，F 分别是 AB，AD 的中点，\n"
          "∴ EF∥BD（三角形中位线定理）。\n"
          "又 EF⊄平面 BCD，\n"
          "　 BD⊂平面 BCD，\n"
          "∴ EF∥平面 BCD。\n\n"
          "（关键：在平面 BCD 内找一条\n"
          "　与 EF 平行的直线）",
          size=10, bold=False, align=WD_ALIGN_PARAGRAPH.LEFT, line=1.25)

cell_text(b[2],
          "（图形区）\n\n"
          "1. 画空间四边形 ABCD，\n"
          "   标出 E、F 与辅助线 BD\n\n"
          "2. 画平面 α，α 内直线 b，\n"
          "   α 外直线 a，a∥b\n\n"
          "3. 门扇、硬纸板转动简图", size=10,
          align=WD_ALIGN_PARAGRAPH.LEFT, line=1.25)
set_col_widths(board, [4.6, 7.2, 4.4])

# ---------- 7. 教学反思 ----------
h7 = heading("7.教学反思")
h7.paragraph_format.space_before = Pt(14)
h7.paragraph_format.keep_with_next = True
sub_heading("三个亮点")
para("① 冲突起于定义：不从定理出发，而是先让学生体会“证明没有公共点”在操作上不可行，"
     "使判定定理成为解决问题的需要，命题的引入有内在动力。", indent_first=22)
para("② 情境提供结构：门扇与硬纸板两个情境都指向“平面外一条直线与平面内一条直线平行”"
     "这一共同结构，学生在观察中自己说出两个关键限定，为条件辨析准备了感性基础。",
     indent_first=22)
para("③ 方法显性化：例题按四步展开，并把“找线”这一隐性策略明确点出，"
     "板书三条件齐全，学生获得的是可迁移的操作路径，而不是对一道题的模仿。",
     indent_first=22)

sub_heading("两个培育")
para("培育直观想象素养：借助实物演示与立体图形，在“看得见”与“想得到”之间建立联系，"
     "使学生能在空间中识别线面位置关系。", indent_first=22)
para("培育逻辑推理素养：通过条件辨析与规范书写，使学生体会定理成立的条件边界，"
     "形成言必有据、步骤完整的表达习惯。", indent_first=22)

sub_heading("一个思考")
para("受 15 分钟的限制，判定定理只能借助情境直观“确认”，未能给出严格证明——事实上，"
     "该定理在教材体系中作为可接受的事实使用，其严格论证要等到后续学习空间向量之后才更自然。"
     "课堂上如何把握“直观确认”与“逻辑证明”的分寸，仍需斟酌。此外，学生对“a⊄α”这一符号"
     "的书写仍不熟练，常在证明中漏写，需要在作业批改中持续强化。若再上一次，"
     "可把情境2 的演示改为学生动手操作，以争取更多的参与时间。", indent_first=22)

doc.save(OUT)
print("saved:", OUT)
