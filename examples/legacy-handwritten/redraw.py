#!/usr/bin/env python3
"""
dryrun-01 第三步：把第 6 页「左边的流程图」重画出来，出两份：
  slide6-left-redraw.pptx   可编辑的形状和文字
  slide6-left-redraw.svg    矢量图

规矩：
  - 元素不增不减。18 个方框、10 条关系（原图 11 个线条对象，C10+C11 经用户确认为一条折线）、
    5 个分区，一个不多一个不少。
  - 原文照抄，不修错。原图的编号重复和「步骤 6 写成开发环境」两处**原样保留**，改之前先问。
  - 配色由我定，但保持原图的颜色分组（同类的还是同色），不把颜色承载的含义拆散。
  - 坐标沿用原图（150dpi 像素 → 英寸）。
"""

from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE, MSO_CONNECTOR
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from lxml import etree
import copy

DPI = 150.0
W_PX, H_PX = 2000, 1125
def IN(px): return Inches(px / DPI)

# ---------- 配色（我定的，但保留原图的颜色分组） ----------
C_MAIN   = "0B84CE"   # 流程主节点：持续集成 / 制品库 / 应用部署
C_CMDB   = "76B73B"   # 外部系统：CMDB
C_ITSM   = "E01B79"   # 外部系统：ITSM
C_CAP    = "2460EF"   # 能力小方块 ×5
C_BAND   = "37A76F"   # 能力通道条
C_ENV    = "F2F2F2"   # 环境框
C_ENVTXT = "E01B79"   # 环境名
C_NOTE   = "F0F0F0"   # 步骤标注底
C_NOTETX = "595959"
C_LINE   = "1577C4"
C_ICON   = "1F3864"
FONT = "微软雅黑"

# ---------- 元素（坐标全部来自原图实测） ----------
BAND = (173, 580, 878, 94)
ENVS = [(39, 798, 275, 242, "开发环境"),
        (330, 798, 275, 242, "测试环境"),
        (622, 798, 275, 242, "预发布环境"),
        (914, 798, 275, 242, "生产环境")]

BOXES = [
    ("持续集成\n（CI）",  83, 183, 192, 82, C_MAIN),
    ("制品库",           498, 184, 192, 82, C_MAIN),
    ("CMDB",              79, 470, 192, 82, C_CMDB),
    ("应用部署\n（CD）",  498, 470, 192, 82, C_MAIN),
    ("ITSM",             930, 470, 192, 82, C_ITSM),
    ("工具库",           194, 625, 137, 42, C_CAP),
    ("流程库",           373, 623, 137, 42, C_CAP),
    ("文件分发",         545, 624, 137, 42, C_CAP),
    ("定时调度",         723, 625, 137, 42, C_CAP),
    ("批量任务",         899, 625, 137, 42, C_CAP),
]

ICONS = [(83, 893), (194, 893), (376, 893), (487, 893),
         (658, 893), (768, 893), (958, 893), (1069, 893)]
ICON_W, ICON_H = 77, 124

NOTES = [
    ("1.代码构建，上传制品",              308, 193, 153, 64),
    ("2.获取应用环境信息",                299, 479, 153, 64),
    ("3.获取应用制品信息，制定部署策略",   510, 326, 153, 84),
    ("5.部署验证后，进行制品晋级",         759, 322, 127, 84),
    ("通过部署工单自动发布",               746, 479, 153, 64),
    ("4.部署到开发环境",                   83, 702, 153, 64),
    ("6.部署到测试环境，制品晋级",         388, 702, 153, 64),   # ← 已按用户确认改为「测试环境」
    ("7.部署到预发布环境，制品晋级",       658, 702, 153, 64),
    ("8.部署到生产环境",                  970, 702, 153, 64),
]

# 连线：(路径点列表, 箭头在最后一点)
# (路径, 起点有没有箭头, 终点有没有箭头)
LINES = [
    ([(275, 224), (498, 224)], False, True),      # 持续集成 → 制品库
    ([(271, 511), (498, 511)], False, True),      # CMDB → 应用部署
    ([(931, 511), (691, 511)], False, True),      # ITSM → 应用部署
    ([(595, 266), (595, 470)], False, True),      # 制品库 → 应用部署
    ([(595, 552), (595, 594)], False, True),      # 应用部署 → 能力通道
    ([(176, 627), (176, 798)], False, True),      # 能力通道 → 开发环境
    ([(468, 657), (468, 798)], False, True),
    ([(760, 653), (760, 798)], False, True),
    ([(1051, 627), (1051, 798)], False, True),
    # ★ 双向关系：一头指向制品库，一头指向下（落在 ITSM→应用部署 那条线上）
    ([(722, 507), (722, 460), (822, 460), (822, 224), (697, 224)], True, True),
]

BAND_TITLE = ("自动化运维平台能力通道", 173, 582, 878, 36)


# =====================================================================
# PPTX
# =====================================================================
A = '{http://schemas.openxmlformats.org/drawingml/2006/main}'


def set_arrow(shape, head=False, tail=False):
    """两头分别控制。双向关系 = 两头都有箭头。"""
    ln = shape.line._get_or_add_ln()
    for tag in ("headEnd", "tailEnd"):
        for e in ln.findall(A + tag):
            ln.remove(e)
    for want, tag in ((head, "headEnd"), (tail, "tailEnd")):
        if want:
            t = etree.SubElement(ln, A + tag)
            t.set("type", "triangle"); t.set("w", "med"); t.set("len", "med")


def rrect(slide, x, y, w, h, fill, radius=0.10, line=None, lw=1.0):
    sh = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, IN(x), IN(y), IN(w), IN(h))
    try:
        sh.adjustments[0] = radius
    except Exception:
        pass
    sh.fill.solid(); sh.fill.fore_color.rgb = RGBColor.from_string(fill)
    if line:
        sh.line.color.rgb = RGBColor.from_string(line); sh.line.width = Pt(lw)
    else:
        sh.line.fill.background()
    sh.shadow.inherit = False
    return sh


def put_text(sh, text, size, color, bold=True):
    tf = sh.text_frame
    tf.word_wrap = True
    tf.margin_left = tf.margin_right = Emu(0)
    tf.margin_top = tf.margin_bottom = Emu(0)
    tf.vertical_anchor = MSO_ANCHOR.MIDDLE
    lines = text.split("\n")
    for i, ln in enumerate(lines):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.alignment = PP_ALIGN.CENTER
        r = p.add_run(); r.text = ln
        r.font.size = Pt(size); r.font.bold = bold
        r.font.color.rgb = RGBColor.from_string(color)
        r.font.name = FONT


def textbox(slide, x, y, w, h, text, size, color, bold=False, align=PP_ALIGN.CENTER):
    tb = slide.shapes.add_textbox(IN(x), IN(y), IN(w), IN(h))
    tf = tb.text_frame; tf.word_wrap = True
    tf.margin_left = tf.margin_right = Emu(0)
    tf.margin_top = tf.margin_bottom = Emu(0)
    tf.vertical_anchor = MSO_ANCHOR.MIDDLE
    for i, ln in enumerate(text.split("\n")):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.alignment = align
        r = p.add_run(); r.text = ln
        r.font.size = Pt(size); r.font.bold = bold
        r.font.color.rgb = RGBColor.from_string(color); r.font.name = FONT
    return tb


def draw_icon(slide, x, y):
    """简化版「虚拟机 Agent」图标：中央机架 + 两侧导轨 + 下方两行字。"""
    cx = x + ICON_W / 2
    rack_w, rack_x, rack_y = 34, cx - 17, y + 2
    rrect(slide, rack_x, rack_y, rack_w, 58, "FFFFFF", radius=0.04,
          line=C_ICON, lw=1.5)
    for i in range(1, 4):
        ln = slide.shapes.add_connector(MSO_CONNECTOR.STRAIGHT,
                                        IN(rack_x), IN(rack_y + i * 14.5),
                                        IN(rack_x + rack_w), IN(rack_y + i * 14.5))
        ln.line.color.rgb = RGBColor.from_string(C_ICON); ln.line.width = Pt(1.0)
    for sx in (cx - 28, cx + 20):
        rrect(slide, sx, y + 12, 8, 38, C_ICON, radius=0.0)
    textbox(slide, x - 10, y + 64, ICON_W + 20, 60, "虚拟机\nAgent",
            10.5, C_ICON, bold=True)


def build_pptx(path):
    prs = Presentation()
    prs.slide_width = IN(2000); prs.slide_height = IN(1125)
    slide = prs.slides.add_slide(prs.slide_layouts[6])

    for (x, y, w, h, _t) in ENVS:
        rrect(slide, x, y, w, h, C_ENV, radius=0.03)
    rrect(slide, *BAND[:4], fill=C_BAND, radius=0.10)

    for (x, y, w, h, name) in ENVS:
        textbox(slide, x + 30, y + 2, w - 60, 64, name, 20, C_ENVTXT, bold=True)

    for pts, has_head, has_tail in LINES:
        for a, b in zip(pts, pts[1:]):
            cn = slide.shapes.add_connector(MSO_CONNECTOR.STRAIGHT,
                                            IN(a[0]), IN(a[1]), IN(b[0]), IN(b[1]))
            cn.line.color.rgb = RGBColor.from_string(C_LINE)
            cn.line.width = Pt(1.75)
            set_arrow(cn,
                      head=(has_head and a == pts[0]),
                      tail=(has_tail and b == pts[-1]))

    for (txt, x, y, w, h, col) in BOXES:
        sh = rrect(slide, x, y, w, h, col, radius=0.14)
        put_text(sh, txt, 17 if h > 60 else 13, "FFFFFF")

    textbox(slide, BAND_TITLE[1], BAND_TITLE[2], BAND_TITLE[3], BAND_TITLE[4],
            BAND_TITLE[0], 15, "FFFFFF", bold=False)

    for (txt, x, y, w, h) in NOTES:
        rrect(slide, x, y, w, h, C_NOTE, radius=0.05)
        textbox(slide, x + 4, y, w - 8, h, txt, 9.5, C_NOTETX, bold=False)

    for (x, y) in ICONS:
        draw_icon(slide, x, y)

    prs.save(path)
    return path


# =====================================================================
# SVG
# =====================================================================
def esc(s):
    return s.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")


def svg_rect(x, y, w, h, fill, rx=None, stroke=None, sw=1.5):
    s = f'<rect x="{x}" y="{y}" width="{w}" height="{h}" fill="#{fill}"'
    if rx:
        s += f' rx="{rx}"'
    if stroke:
        s += f' stroke="#{stroke}" stroke-width="{sw}"'
    return s + "/>"


def svg_text(x, y, lines, size, fill, bold=False, anchor="middle"):
    w = ' font-weight="bold"' if bold else ""
    out = [f'<text x="{x}" y="{y}" font-family="{FONT}, PingFang SC, sans-serif" '
           f'font-size="{size}" fill="#{fill}" text-anchor="{anchor}"{w}>']
    for i, ln in enumerate(lines):
        dy = 0 if i == 0 else size * 1.15
        out.append(f'<tspan x="{x}" dy="{dy}">{esc(ln)}</tspan>')
    out.append("</text>")
    return "".join(out)


def build_svg(path):
    p = [f'<svg xmlns="http://www.w3.org/2000/svg" width="{W_PX}" height="{H_PX}" '
         f'viewBox="0 0 {W_PX} {H_PX}">',
         f'<rect width="{W_PX}" height="{H_PX}" fill="#FFFFFF"/>',
         '<defs><marker id="ah" viewBox="0 0 10 10" refX="9" refY="5" '
         'markerWidth="7" markerHeight="7" orient="auto-start-reverse">'
         f'<path d="M 0 0 L 10 5 L 0 10 z" fill="#{C_LINE}"/></marker></defs>']

    for (x, y, w, h, _t) in ENVS:
        p.append(svg_rect(x, y, w, h, C_ENV, rx=8))
    p.append(svg_rect(*BAND[:4], fill=C_BAND, rx=10))

    for (x, y, w, h, name) in ENVS:
        p.append(svg_text(x + w / 2, y + 42, [name], 22, C_ENVTXT, bold=True))

    for pts, has_head, has_tail in LINES:
        d = " ".join(f'{"M" if i == 0 else "L"} {a[0]} {a[1]}'
                     for i, a in enumerate(pts))
        m = (' marker-start="url(#ah)"' if has_head else "") + \
            (' marker-end="url(#ah)"' if has_tail else "")
        p.append(f'<path d="{d}" fill="none" stroke="#{C_LINE}" '
                 f'stroke-width="2"{m}/>')

    for (txt, x, y, w, h, col) in BOXES:
        p.append(svg_rect(x, y, w, h, col, rx=10))
        p.append(svg_text(x + w / 2, y + h / 2 - (len(txt.split(chr(10))) - 1) * 9 + 7,
                          txt.split("\n"), 19 if h > 60 else 15, "FFFFFF", bold=True))

    p.append(svg_text(BAND_TITLE[1] + BAND_TITLE[3] / 2, BAND_TITLE[2] + 24,
                      [BAND_TITLE[0]], 16, "FFFFFF"))

    for (txt, x, y, w, h) in NOTES:
        p.append(svg_rect(x, y, w, h, C_NOTE, rx=5))
        n = len(txt) // 9 + 1
        p.append(svg_text(x + w / 2, y + h / 2 - (n - 1) * 7 + 5,
                          [txt[i:i + 9] for i in range(0, len(txt), 9)],
                          11, C_NOTETX))

    for (x, y) in ICONS:
        cx = x + ICON_W / 2
        p.append(svg_rect(cx - 17, y + 2, 34, 58, "FFFFFF", rx=2,
                          stroke=C_ICON, sw=1.6))
        for i in range(1, 4):
            p.append(f'<line x1="{cx-17}" y1="{y+2+i*14.5}" x2="{cx+17}" '
                     f'y2="{y+2+i*14.5}" stroke="#{C_ICON}" stroke-width="1.2"/>')
        for sx in (cx - 28, cx + 20):
            p.append(svg_rect(sx, y + 12, 8, 38, C_ICON))
        p.append(svg_text(cx, y + 92, ["虚拟机", "Agent"], 13, C_ICON, bold=True))

    p.append("</svg>")
    open(path, "w").write("\n".join(p))
    return path


if __name__ == "__main__":
    out = "diagram-redraw/dryrun-01/"
    print("PPTX:", build_pptx(out + "slide6-left-redraw.pptx"))
    print("SVG :", build_svg(out + "slide6-left-redraw.svg"))
