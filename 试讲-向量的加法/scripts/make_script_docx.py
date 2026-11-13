# -*- coding: utf-8 -*-
"""生成《向量的加法》8 分钟试讲逐字稿"""
from docx import Document
from docx.shared import Pt, Cm, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

OUT = "/workspace/试讲-向量的加法/逐字稿-向量的加法.docx"

doc = Document()
sec = doc.sections[0]
sec.page_width = Cm(21.0)
sec.page_height = Cm(29.7)
sec.left_margin = Cm(2.6)
sec.right_margin = Cm(2.6)
sec.top_margin = Cm(2.4)
sec.bottom_margin = Cm(2.4)

style = doc.styles["Normal"]
style.font.name = "Times New Roman"
style.font.size = Pt(12)
style.element.rPr.rFonts.set(qn("w:eastAsia"), "宋体")
style.paragraph_format.line_spacing = 1.35
style.paragraph_format.space_after = Pt(0)

NAVY = (0x1F, 0x38, 0x64)
BLUE = (0x2E, 0x53, 0x95)
GRAY = (0x6B, 0x73, 0x80)


def set_font(run, name_cn="宋体", name_en="Times New Roman", size=12,
             bold=False, color=None):
    run.font.name = name_en
    run.font.size = Pt(size)
    run.font.bold = bold
    if color is not None:
        run.font.color.rgb = RGBColor(*color)
    rPr = run._element.get_or_add_rPr()
    rFonts = rPr.find(qn("w:rFonts"))
    if rFonts is None:
        rFonts = OxmlElement("w:rFonts")
        rPr.append(rFonts)
    rFonts.set(qn("w:eastAsia"), name_cn)
    rFonts.set(qn("w:ascii"), name_en)
    rFonts.set(qn("w:hAnsi"), name_en)


def para(runs, size=12, line=1.35, before=0, after=0, indent_char=0,
         align=WD_ALIGN_PARAGRAPH.JUSTIFY):
    p = doc.add_paragraph()
    p.alignment = align
    pf = p.paragraph_format
    pf.line_spacing = line
    pf.space_before = Pt(before)
    pf.space_after = Pt(after)
    if indent_char:
        pPr = p._p.get_or_add_pPr()
        ind = pPr.find(qn("w:ind"))
        if ind is None:
            ind = OxmlElement("w:ind")
            pPr.append(ind)
        ind.set(qn("w:firstLineChars"), str(int(indent_char * 100)))
    if isinstance(runs, str):
        runs = [(runs, {})]
    for text, opt in runs:
        r = p.add_run(text)
        set_font(r, opt.get("cn", "宋体"), opt.get("en", "Times New Roman"),
                 opt.get("size", size), opt.get("bold", False),
                 opt.get("color"))
    return p


def title(text):
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.space_after = Pt(6)
    r = p.add_run(text)
    set_font(r, "黑体", "Times New Roman", 18, True, NAVY)
    return p


def h1(text):
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(8)
    p.paragraph_format.space_after = Pt(4)
    p.paragraph_format.keep_with_next = True
    r = p.add_run(text)
    set_font(r, "黑体", "Times New Roman", 13, True, NAVY)
    return p


def say(text, keep=False):
    """教师讲授语"""
    p = para([("【师】", {"bold": True, "cn": "黑体"}),
              (text, {})], indent_char=2)
    if keep:
        p.paragraph_format.keep_with_next = True
    return p


def stu(text):
    """预设学生回答"""
    p = para([("【生】", {"bold": True, "cn": "黑体", "color": BLUE}),
              (text, {"color": BLUE})], indent_char=2)
    p.paragraph_format.keep_with_next = True
    return p


def act(text):
    """板书、课件与动作提示"""
    p = para([("〔" + text + "〕", {"cn": "楷体", "size": 10.5, "color": GRAY})],
             indent_char=2, before=2, after=2)
    p.paragraph_format.keep_with_next = True
    return p


# ================= 标题与说明 =================
title("《向量的加法》试讲逐字稿")
para([("北师大版（2019）必修第二册　§2.1 向量的加法",
       {"size": 10.5, "color": GRAY})],
     align=WD_ALIGN_PARAGRAPH.CENTER, after=2)
para([("课型：新授课试讲　|　课件：智趣通关·向量的加法　|　总时长 8 分钟",
       {"size": 10.5, "color": GRAY})],
     align=WD_ALIGN_PARAGRAPH.CENTER, after=8)

para([("使用说明：", {"bold": True, "cn": "黑体", "size": 10.5}),
      ("【师】为教师讲授语，建议按原话讲；【生】为预设学生回答，"
       "学生答出即接下一句，答不出则按括号内的提示降阶追问；"
       "〔  〕为板书、课件操作与动作提示，不念出来。全稿正文约 1200 字，"
       "按每分钟 200 字计约 6 分钟，余下的时间留给课件演示、板书书写与必要停顿。",
       {"size": 10.5})],
     after=10)

# ================= 时间分配表 =================
h1("时间分配一览")
tb = doc.add_table(rows=1, cols=4)
tb.style = "Table Grid"
tb.alignment = WD_TABLE_ALIGNMENT.CENTER
hdr = tb.rows[0].cells
for i, t in enumerate(["环节", "时间", "时长", "教与学的要点"]):
    cell = hdr[i]
    cell.text = ""
    p = cell.paragraphs[0]
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.space_before = Pt(1)
    p.paragraph_format.space_after = Pt(1)
    r = p.add_run(t)
    set_font(r, "黑体", "Times New Roman", 10.5, True)
    tcPr = cell._tc.get_or_add_tcPr()
    shd = OxmlElement("w:shd")
    shd.set(qn("w:val"), "clear")
    shd.set(qn("w:fill"), "E8EDF6")
    tcPr.append(shd)

rows = [
    ("复习导入", "0:00—0:40", "40″", "由“数量能相加”引出“向量能否相加”"),
    ("情境探究", "0:40—2:10", "1′30″", "天车合位移，破除“3+4=7”"),
    ("形成法则", "2:10—3:50", "1′40″", "三角形法则、平行四边形法则、交换律"),
    ("辨析", "3:50—5:20", "1′30″", "共线情形与 |a+b| 的范围"),
    ("应用", "5:20—6:40", "1′20″", "二力合成、小船渡河"),
    ("通关小结", "6:40—7:40", "1′00″", "回扣课件三个问号"),
    ("布置作业", "7:40—8:00", "20″", "智变训练 + 预习例2"),
]
for r0, r1, r2, r3 in rows:
    cells = tb.add_row().cells
    for i, txt in enumerate((r0, r1, r2, r3)):
        cells[i].text = ""
        p = cells[i].paragraphs[0]
        p.alignment = (WD_ALIGN_PARAGRAPH.CENTER if i < 3
                       else WD_ALIGN_PARAGRAPH.LEFT)
        p.paragraph_format.space_before = Pt(1)
        p.paragraph_format.space_after = Pt(1)
        p.paragraph_format.line_spacing = 1.25
        rr = p.add_run(txt)
        set_font(rr, "宋体", "Times New Roman", 10)

widths = [3.2, 3.4, 2.4, 6.8]
tb.autofit = False
tblPr = tb._tbl.tblPr
for old in tblPr.findall(qn("w:tblW")):
    tblPr.remove(old)
tblW = OxmlElement("w:tblW")
tblW.set(qn("w:type"), "dxa")
tblW.set(qn("w:w"), str(int(sum(widths) * 567)))
tblPr.append(tblW)
for old in tblPr.findall(qn("w:tblLayout")):
    tblPr.remove(old)
lay = OxmlElement("w:tblLayout")
lay.set(qn("w:type"), "fixed")
tblPr.append(lay)
grid = tb._tbl.find(qn("w:tblGrid"))
if grid is not None:
    tb._tbl.remove(grid)
grid = OxmlElement("w:tblGrid")
for w in widths:
    gc = OxmlElement("w:gridCol")
    gc.set(qn("w:w"), str(int(w * 567)))
    grid.append(gc)
tb._tbl.insert(list(tb._tbl).index(tblPr) + 1, grid)
for row in tb.rows:
    for i, cell in enumerate(row.cells):
        cell.width = Cm(widths[i])
    trPr = row._tr.get_or_add_trPr()
    trPr.append(OxmlElement("w:cantSplit"))

# ================= 一、复习导入 =================
h1("一、复习导入，提出新运算（0:00—0:40）")
say("上课！同学们好，请坐。")
say("前面我们学习了向量。谁能用一句话说说，向量和数量最大的不同是什么？")
stu("向量既有大小，又有方向。")
say("非常好。数量可以相加，那向量能不能相加？如果能，是直接把长度相加吗？")
say("今天我们就带着这个问题，进入《向量的加法》。")
act("课件：切到“复习回顾—情境舞台”")
say("请看屏幕上的天车吊物。天车把货物吊起 3 米，再向右移动 4 米。"
    "大家先估一估：货物最后的合位移是多少？")
stu("5 米？")
stu("7 米？")
say("有同学说 7 米。3 加 4 等于 7，听起来很顺。但真的是 7 米吗？我们让天车走一遍。")
act("课件：3D 天车演示上升 3 米、右移 4 米")

# ================= 二、情境探究 =================
h1("二、情境探究：为什么不是 7 米（0:40—2:10）")
say("看终点。货物并没有沿着“先竖直 3 米、再水平 4 米”的折线长度走，"
    "它最终的合位移是从起点指向终点。")
say("谁能解释，为什么不能直接 3 加 4？")
stu("因为位移有方向，不是同一条直线上的长度相加。")
say("抓住关键了。位移有方向，长度不能直接相加。")
act("板书：位移有方向，不能只加长度")
say("我们把上升这段位移记作向量 AB，右移这段记作向量 BC。"
    "那么从起点 A 到终点 C 的合位移，就是向量 AC。")
say("也就是说，向量 AB 加向量 BC 等于向量 AC。")
say("大家看，这两个向量是怎么放的？")
stu("首尾相接。")
say("对，第一个向量的终点，接着第二个向量的起点。和向量从哪里指向哪里？")
stu("从第一个起点指向最后一个终点。")
say("这就是向量加法的三角形法则。")
act("板书：三角形法则——首尾相接，起点指向终点")

# ================= 三、形成法则 =================
h1("三、形成法则：三角形与平行四边形（2:10—3:50）")
say("如果两个向量不是首尾相接，而是从同一点出发，比如两个力同时作用在一个物体上，"
    "又该怎么合成呢？")
act("课件：切到“平行四边形法则”示意")
say("以同一点 O 为起点作向量 a、向量 b，以它们为邻边补成平行四边形。"
    "大家猜，和向量是哪一条？")
stu("对角线。")
say("对，从公共起点出发的那条对角线，就是向量 a 加向量 b。这叫平行四边形法则。")
act("板书：平行四边形法则——同起点，对角线")
say("那问题来了：三角形法则和平行四边形法则，得到的是同一个和吗？")
stu("是。")
say("为什么？", keep=True)
stu("把平行四边形的一条边平移过去，就变成首尾相接了。")
say("很好。两个法则本质一致，只是摆放方式不同：首尾相接用三角形，同起点用平行四边形。")
say("再看一个细节：先加向量 a 再加向量 b，和先加向量 b 再加向量 a，终点会变吗？")
act("课件：交换次序演示")
stu("不变。")
say("所以向量加法满足交换律：向量 a 加向量 b 等于向量 b 加向量 a。"
    "多个向量首尾相接，还可以得到多边形法则。")

# ================= 四、辨析 =================
h1("四、辨析共线向量与模的范围（3:50—5:20）")
say("刚才天车上升 3 米、右移 4 米，合位移是 5 米。"
    "那是不是所有向量相加，结果都比直接加长度小？")
say("我们看共线的情况。如果两个向量同向，向量 a 长度 4，向量 b 长度 3，和向量多长？")
stu("7。")
say("如果反向呢？")
stu("1。")
say("对。所以和向量的模的大小，会在一个范围内变化。")
act("课件：拖动夹角，显示范围")
p = para([("一般地，", {}), ("｜|a|−|b|｜ ≤ |a+b| ≤ |a|+|b|", {"bold": True}),
          ("。", {})], indent_char=2, before=2, after=2)
p.alignment = WD_ALIGN_PARAGRAPH.CENTER
say("什么时候取最大？")
stu("同向。")
say("什么时候取最小？")
stu("反向。")
say("注意，这个不等式告诉我们：向量和的模，不一定等于模的和。"
    "方向，才是向量加法里最关键的变量。")
act("板书：同向最大，反向最小；模不等式")

# ================= 五、应用 =================
h1("五、应用：合力与航速（5:20—6:40）")
say("我们回到生活。课件里有两个 450 牛的力，夹角 60 度。直接相加是 900 牛，对不对？")
stu("不对。")
say("为什么？")
stu("方向不同，要用平行四边形法则。")
say("对，合力约是 450 乘根号 3，也就是约 779 牛。方向不同，合力不是简单相加。")
say("再看小船渡河。船在静水中速度 4 米每秒，船头垂直河岸，水流速度可调。船实际会怎样？")
stu("会被水冲偏。")
say("实际航速是船速和水速的合速度。所以解决位移、力、速度问题，"
    "都不能只看大小，还要看方向。")
act("课件：切到“应用”或“趣味数学”页面")

# ================= 六、小结 =================
h1("六、通关小结（6:40—7:40）")
say("我们回到课件的三个问号。第一，天车吊物为什么不是 7 米？")
stu("位移有方向，不能直接加长度。")
say("第二，两个法则为什么得到同一个和？")
stu("本质都是首尾相接，平行四边形可以转化为三角形。")
say("第三，交换次序终点会变吗？")
stu("不会，向量加法满足交换律。")
say("好，今天我们学习了向量加法：三角形法则、平行四边形法则、共线向量、"
    "模的范围和运算律。核心一句话：")
act("板书：向量加法 = 大小 + 方向")
say("向量加法，加的不只是长度，更是方向。")

# ================= 七、作业 =================
h1("七、布置作业（7:40—8:00）")
say("今天的作业：用课件完成“智变训练”前 5 题，重点体会“首尾相接”和"
    "“同起点对角线”两种作图；并预习例2，思考多个向量相加怎样用多边形法则。下课！")

doc.save(OUT)
print("saved:", OUT)
