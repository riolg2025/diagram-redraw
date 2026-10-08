#!/usr/bin/env python3
"""
build_deck.py —— 把 spec.json 画成 PPTX（可编辑）和/或 SVG（矢量）。

**只认 spec，不认 pptx。** 这是"读"和"画"分开的地方：
换一个输入，只要能产出同样格式的 spec，这一步就不用动。

用法：
  python3 build_deck.py spec.json --pptx out.pptx --svg out.svg

里面每一个细节都对应一个踩过的坑：

  * 描边 0.5pt —— 原图实测就是 0.5pt（1px@150dpi）。给粗了看起来像"文字压边框"。
  * latin + ea + cs 都设字体 —— 只设 latin，中文会回落到主题默认字体，
    字宽不一样，会多余换行。
  * 字体带 panose/pitchFamily/charset(GB2312) —— 不带这些，替代字体可能选得更宽。
  * word_wrap=False —— 原图 bodyPr 是 vertOverflow/horzOverflow=overflow，不强制换行。
  * 文本框保留 OOXML 默认内边距（左右 15px、上下 7.5px）—— 设成 0 文字会整体左上偏。
  * SVG 里的 pt -> px 要 ×150/72 —— SVG 坐标是像素，不是磅。
  * SVG 的 marker 必须先定义再引用 —— 漏了 `<defs>` 就是"16 个箭头全没有头"。
  * upDownArrow 要画成两个尖 + 一根杆 —— 画成"一个尖 + 倒钩"就成了单向箭头。
  * 括号是"横线 + 中间一个直角尖" —— 用曲线画会又宽又圆，不像。
"""

import argparse
import base64
import os
import sys

from lxml import etree
from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.dml import MSO_LINE_DASH_STYLE
from pptx.enum.shapes import MSO_CONNECTOR, MSO_SHAPE
from pptx.enum.text import MSO_ANCHOR, PP_ALIGN
from pptx.util import Emu, Inches, Pt
from PIL import Image

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from common import load_spec

A = '{http://schemas.openxmlformats.org/drawingml/2006/main}'
DPI = 150.0
PT2PX = DPI / 72.0

DASH = {'dash': MSO_LINE_DASH_STYLE.DASH,
        'sysDash': MSO_LINE_DASH_STYLE.DASH_DOT,
        'lgDash': MSO_LINE_DASH_STYLE.LONG_DASH,
        'dot': MSO_LINE_DASH_STYLE.ROUND_DOT,
        'sysDot': MSO_LINE_DASH_STYLE.ROUND_DOT}

SHAPE_MAP = {
    'rect': MSO_SHAPE.RECTANGLE,
    'container': MSO_SHAPE.RECTANGLE,
    'roundrect': MSO_SHAPE.ROUNDED_RECTANGLE,
    'diamond': MSO_SHAPE.DIAMOND,
    'ellipse': MSO_SHAPE.OVAL,
    'upArrow': MSO_SHAPE.UP_ARROW,
    'downArrow': MSO_SHAPE.DOWN_ARROW,
    'leftArrow': MSO_SHAPE.LEFT_ARROW,
    'rightArrow': MSO_SHAPE.RIGHT_ARROW,
    'upDownArrow': MSO_SHAPE.UP_DOWN_ARROW,
    'leftRightArrow': MSO_SHAPE.LEFT_RIGHT_ARROW,
}
FONT = "微软雅黑"
FONT_ATTRS = {"panose": "020B0604030504040204",
              "pitchFamily": "34", "charset": "-122"}
DEFAULT_LW_PT = 0.5          # 原图实测


def IN(px):
    return Inches(px / DPI)


def set_arrow(sh, head=False, tail=False):
    ln = sh.line._get_or_add_ln()
    for tag in ("headEnd", "tailEnd"):
        for e in ln.findall(A + tag):
            ln.remove(e)
    for want, tag in ((head, "headEnd"), (tail, "tailEnd")):
        if want:
            t = etree.SubElement(ln, A + tag)
            t.set("type", "triangle")
            t.set("w", "med")
            t.set("len", "med")


def style_text(tf, lines, size_pt, color, align, insets=None, anchor=None,
               face=FONT, wrap=True):
    tf.word_wrap = wrap
    l, t, r, b = insets or (15.0, 7.5, 15.0, 7.5)
    tf.margin_left, tf.margin_right = Emu(int(l / DPI * 914400)), Emu(int(r / DPI * 914400))
    tf.margin_top, tf.margin_bottom = Emu(int(t / DPI * 914400)), Emu(int(b / DPI * 914400))
    tf.vertical_anchor = {"ctr": MSO_ANCHOR.MIDDLE, "t": MSO_ANCHOR.TOP,
                          "b": MSO_ANCHOR.BOTTOM}.get(anchor or 'ctr',
                                                       MSO_ANCHOR.MIDDLE)
    for i, ln in enumerate(lines):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.alignment = align
        run = p.add_run()
        run.text = ln
        run.font.size = Pt(size_pt)
        run.font.color.rgb = RGBColor.from_string(color)
        run.font.name = face
        rPr = run._r.get_or_add_rPr()
        for tag in ("latin", "ea", "cs"):
            e = rPr.find(A + tag)
            if e is None:
                e = etree.SubElement(rPr, A + tag)
            e.set("typeface", face)
            for k, v in FONT_ATTRS.items():
                e.set(k, v)


def add_shape(sl, el):
    kind = el["kind"]
    b = el["box"]
    shp = SHAPE_MAP.get(kind, MSO_SHAPE.RECTANGLE)
    sh = sl.shapes.add_shape(shp, IN(b[0]), IN(b[1]), IN(b[2]), IN(b[3]))
    if kind == "roundrect":
        try:
            sh.adjustments[0] = 0.16
        except Exception:
            pass
    fill = el.get("fill")
    if fill is None:
        sh.fill.background()      # 没底色才透明；容器**不等于**透明
    else:
        sh.fill.solid()
        sh.fill.fore_color.rgb = RGBColor.from_string(fill)
    lc = el.get("line") or el.get("line_color")
    if lc:
        sh.line.color.rgb = RGBColor.from_string(lc)
        sh.line.width = Pt(el.get("line_pt", DEFAULT_LW_PT))
    else:
        sh.line.fill.background()
    sh.shadow.inherit = False
    txt = el.get("text") or ""
    if txt:
        align = PP_ALIGN.CENTER if kind != "text" else PP_ALIGN.LEFT
        style_text(sh.text_frame, txt.split("\n"), el.get("font_pt", 8.0),
                   el.get("fg", "333333"), align,
                   el.get("insets"), el.get("anchor"),
                   wrap=(el.get("wrap", "square") != "none"))
    return sh


def build_pptx(spec, path):
    cw, ch = spec["source"]["canvas"]
    prs = Presentation()
    prs.slide_width, prs.slide_height = IN(cw), IN(ch)
    sl = prs.slides.add_slide(prs.slide_layouts[6])

    E = spec["elements"]
    order = ["container", "rect", "roundrect", "diamond", "ellipse",
             "line", "brace", "arrow", "picture", "text"]
    pics = []
    for kind in order:
        for el in [e for e in E if e["kind"] == kind]:
            if kind in ("rect", "container", "roundrect", "diamond", "ellipse"):
                add_shape(sl, el)

            elif kind == "line":
                p1, p2 = el["p1"], el["p2"]
                cn = sl.shapes.add_connector(MSO_CONNECTOR.STRAIGHT,
                                             IN(p1[0]), IN(p1[1]),
                                             IN(p2[0]), IN(p2[1]))
                cn.line.color.rgb = RGBColor.from_string(el.get("color", "333333"))
                cn.line.width = Pt(max(el.get("width_pt", 0.75), 0.4))
                if el.get("dash", "solid") != "solid":
                    cn.line.dash_style = DASH.get(el["dash"], MSO_LINE_DASH_STYLE.DASH)
                if el.get("arrow") == "end":
                    set_arrow(cn, tail=True)
                elif el.get("arrow") == "both":
                    set_arrow(cn, head=True, tail=True)

            elif kind == "brace":
                b = el["box"]
                cx, cy = b[0] + b[2] / 2, b[1] + b[3] / 2
                sh = sl.shapes.add_shape(MSO_SHAPE.LEFT_BRACE,
                                         IN(cx - b[2] / 2), IN(cy - b[3] / 2),
                                         IN(b[2]), IN(b[3]))
                sh.rotation = el.get("rot", 0)
                sh.fill.background()
                sh.line.color.rgb = RGBColor.from_string(el.get("color", "404040"))
                sh.line.width = Pt(DEFAULT_LW_PT)
                sh.shadow.inherit = False

            elif kind == "arrow":
                b = el["box"]
                sh = sl.shapes.add_shape(SHAPE_MAP.get(el["prst"], MSO_SHAPE.RIGHT_ARROW),
                                         IN(b[0]), IN(b[1]), IN(b[2]), IN(b[3]))
                sh.fill.solid()
                sh.fill.fore_color.rgb = RGBColor.from_string(el.get("fill", "808080"))
                sh.line.fill.background()
                sh.shadow.inherit = False

            elif kind == "picture":
                b = el["box"]
                src = os.path.join(os.path.dirname(os.path.abspath(spec["_path"])),
                                   el["src"])
                sl.shapes.add_picture(src, IN(b[0]), IN(b[1]), IN(b[2]), IN(b[3]))

            elif kind == "text":
                b = el["box"]
                tb = sl.shapes.add_textbox(IN(b[0]), IN(b[1]), IN(b[2]), IN(b[3]))
                style_text(tb.text_frame, (el.get("text") or "").split("\n"),
                           el.get("font_pt", 10.0), el.get("color", "264180"),
                           PP_ALIGN.LEFT, el.get("insets"), el.get("anchor"),
                           wrap=(el.get("wrap", "square") != "none"))

    prs.save(path)
    return path


# --------------------------------------------------------------------------
# SVG
# --------------------------------------------------------------------------
def esc(s):
    return (s or "").replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")


def svg_text(x, y, lines, size_pt, fill, anchor="middle", face=FONT):
    size = size_pt * PT2PX
    out = [f'<text x="{x:.1f}" y="{y:.1f}" font-size="{size:.2f}" fill="#{fill}" '
           f'text-anchor="{anchor}" '
           f'font-family="{face}, PingFang SC, Microsoft YaHei, sans-serif">']
    for i, ln in enumerate(lines):
        out.append(f'<tspan x="{x:.1f}" dy="{0 if i == 0 else size * 1.15:.2f}">'
                   f'{esc(ln)}</tspan>')
    out.append("</text>")
    return "".join(out)


def build_svg(spec, path):
    cw, ch = spec["source"]["canvas"]
    o = [f'<svg xmlns="http://www.w3.org/2000/svg" width="{cw}" height="{ch}" '
         f'viewBox="0 0 {cw} {ch}">',
         f'<rect width="{cw}" height="{ch}" fill="#FFFFFF"/>',
         # ★ marker 必须先定义。漏了这句，下面 16 处 marker-end 全部失效 -> 箭头全没头
         '<defs><marker id="ah" viewBox="0 0 10 10" refX="9" refY="5" '
         'markerWidth="6" markerHeight="6" orient="auto-start-reverse">'
         '<path d="M 0 0 L 10 5 L 0 10 z" fill="context-stroke"/></marker></defs>']

    def rect_of(el, rx=0):
        b = el["box"]
        fill = el.get("fill") or "none"
        st = ""
        lc = el.get("line") or el.get("line_color")
        if lc:
            st = (f' stroke="#{lc}" stroke-width="'
                  f'{el.get("line_pt", DEFAULT_LW_PT) * PT2PX:.2f}"')
        return (f'<rect x="{b[0]:.1f}" y="{b[1]:.1f}" width="{b[2]:.1f}" '
                f'height="{b[3]:.1f}" rx="{rx}" fill="#{fill}"{st}/>')

    E = spec["elements"]
    for kind, rx in (("container", 3), ("rect", 3), ("roundrect", 10)):
        for el in [e for e in E if e["kind"] == kind]:
            fill = el.get("fill")
            if kind == "container" and not fill:
                el = dict(el, fill=None)
            o.append(rect_of(el, rx))
            if el.get("text"):
                b = el["box"]
                ins = el.get("insets") or [15, 7.5, 15, 7.5]
                lines = el["text"].split("\n")
                sz = el.get("font_pt", 8.0)
                y = b[1] + b[3] / 2 - (len(lines) - 1) * sz * PT2PX * 0.58 + sz * PT2PX * 0.35
                o.append(svg_text(b[0] + b[2] / 2, y, lines, sz,
                                  el.get("fg", "333333"), face=el.get("face") or FONT))

    for el in [e for e in E if e["kind"] == "diamond"]:
        b = el["box"]
        cx, cy = b[0] + b[2] / 2, b[1] + b[3] / 2
        st = ""
        if el.get("line"):
            st = f' stroke="#{el["line"]}" stroke-width="{DEFAULT_LW_PT * PT2PX:.2f}"'
        o.append(f'<polygon points="{cx:.1f},{b[1]:.1f} {b[0]+b[2]:.1f},{cy:.1f} '
                 f'{cx:.1f},{b[1]+b[3]:.1f} {b[0]:.1f},{cy:.1f}" '
                 f'fill="#{el.get("fill") or "none"}"{st}/>')

    for el in [e for e in E if e["kind"] == "ellipse"]:
        b = el["box"]
        cx, cy = b[0] + b[2] / 2, b[1] + b[3] / 2
        o.append(f'<ellipse cx="{cx:.1f}" cy="{cy:.1f}" rx="{b[2]/2:.1f}" '
                 f'ry="{b[3]/2:.1f}" fill="#{el.get("fill") or "none"}"/>')
        if el.get("text"):
            o.append(svg_text(cx, cy + el.get("font_pt", 8) * PT2PX * 0.35,
                              el["text"].split("\n"), el.get("font_pt", 8),
                              el.get("fg", "333333"), face=el.get("face") or FONT))

    for el in [e for e in E if e["kind"] == "line"]:
        p1, p2 = el["p1"], el["p2"]
        dash = (' stroke-dasharray="7,4"'
                if el.get("dash", "solid") != "solid" else "")
        mk = ""
        if el.get("arrow") == "end":
            mk = ' marker-end="url(#ah)"'
        elif el.get("arrow") == "both":
            mk = ' marker-start="url(#ah)" marker-end="url(#ah)"'
        o.append(f'<line x1="{p1[0]:.1f}" y1="{p1[1]:.1f}" x2="{p2[0]:.1f}" '
                 f'y2="{p2[1]:.1f}" stroke="#{el.get("color", "333333")}" '
                 f'stroke-width="{max(el.get("width_pt", 0.75), 0.4) * PT2PX:.2f}"'
                 f'{dash}{mk}/>')

    for el in [e for e in E if e["kind"] == "brace"]:
        b = el["box"]
        cx, cy = b[0] + b[2] / 2, b[1] + b[3] / 2
        half = b[3] / 2
        dep = (b[2] / 2) * 0.75
        sign = -1 if el.get("rot", 0) == 90 else 1
        t = max(dep, 6)
        d = (f'M {cx-half:.1f} {cy:.1f} L {cx-t:.1f} {cy:.1f} '
             f'L {cx:.1f} {cy+sign*dep:.1f} L {cx+t:.1f} {cy:.1f} '
             f'L {cx+half:.1f} {cy:.1f}')
        o.append(f'<path d="{d}" fill="none" stroke="#{el.get("color", "404040")}" '
                 f'stroke-width="{DEFAULT_LW_PT * PT2PX:.2f}" stroke-linejoin="miter"/>')

    for el in [e for e in E if e["kind"] == "arrow"]:
        b = el["box"]
        x, y, w, h = b
        f_ = el.get("fill", "808080")
        p = el["prst"]
        if p in ("upDownArrow", "downArrow", "upArrow"):
            hh = h * 0.34
            sw = max(w * 0.42, 2)
            sx0, sx1 = x + w / 2 - sw / 2, x + w / 2 + sw / 2
            if p == "upDownArrow":
                d = (f'M {x+w/2:.1f} {y:.1f} L {x+w:.1f} {y+hh:.1f} L {sx1:.1f} {y+hh:.1f} '
                     f'L {sx1:.1f} {y+h-hh:.1f} L {x+w:.1f} {y+h-hh:.1f} L {x+w/2:.1f} {y+h:.1f} '
                     f'L {x:.1f} {y+h-hh:.1f} L {sx0:.1f} {y+h-hh:.1f} L {sx0:.1f} {y+hh:.1f} '
                     f'L {x:.1f} {y+hh:.1f} Z')
            elif p == "downArrow":
                d = (f'M {sx0:.1f} {y:.1f} L {sx1:.1f} {y:.1f} L {sx1:.1f} {y+h-hh:.1f} '
                     f'L {x+w:.1f} {y+h-hh:.1f} L {x+w/2:.1f} {y+h:.1f} L {x:.1f} {y+h-hh:.1f} '
                     f'L {sx0:.1f} {y+h-hh:.1f} Z')
            else:
                d = (f'M {sx0:.1f} {y+h:.1f} L {x+w/2:.1f} {y:.1f} L {sx1:.1f} {y+h:.1f} '
                     f'L {sx1:.1f} {y+hh:.1f} L {x+w:.1f} {y+hh:.1f} L {x+w:.1f} {y+h-hh:.1f} '
                     f'L {x:.1f} {y+h-hh:.1f} L {x:.1f} {y+hh:.1f} L {sx0:.1f} {y+hh:.1f} Z')
        else:
            hh = w * 0.34
            sh_ = max(h * 0.42, 2)
            sy0, sy1 = y + h / 2 - sh_ / 2, y + h / 2 + sh_ / 2
            d = (f'M {x:.1f} {y+h/2:.1f} L {x+w-hh:.1f} {y:.1f} L {x+w-hh:.1f} {sy0:.1f} '
                 f'L {x+w:.1f} {sy0:.1f} L {x+w:.1f} {sy1:.1f} L {x+w-hh:.1f} {sy1:.1f} '
                 f'L {x+w-hh:.1f} {y+h:.1f} Z')
        o.append(f'<path d="{d}" fill="#{f_}"/>')

    for el in [e for e in E if e["kind"] == "picture"]:
        b = el["box"]
        src = os.path.join(os.path.dirname(os.path.abspath(spec["_path"])), el["src"])
        ext = os.path.splitext(src)[1].lstrip('.').lower()
        mime = {"png": "image/png", "jpg": "image/jpeg", "jpeg": "image/jpeg",
                "gif": "image/gif", "bmp": "image/bmp"}.get(ext, "image/png")
        b64 = base64.b64encode(open(src, "rb").read()).decode()
        o.append(f'<image x="{b[0]:.1f}" y="{b[1]:.1f}" width="{b[2]:.1f}" '
                 f'height="{b[3]:.1f}" preserveAspectRatio="none" '
                 f'xlink:href="data:{mime};base64,{b64}" '
                 f'xmlns:xlink="http://www.w3.org/1999/xlink"/>')

    for el in [e for e in E if e["kind"] == "text"]:
        b = el["box"]
        ins = el.get("insets") or [15, 7.5, 15, 7.5]
        lines = (el.get("text") or "").split("\n")
        sz = el.get("font_pt", 10.0)
        anchor = el.get("anchor") or 't'
        if anchor == 'ctr':
            y0 = b[1] + b[3] / 2 - (len(lines) - 1) * sz * PT2PX * 0.58
        else:
            y0 = b[1] + ins[1] + sz * PT2PX * 0.85
        for i, ln in enumerate(lines):
            o.append(f'<text x="{b[0]+ins[0]:.1f}" '
                     f'y="{y0 + i * sz * PT2PX * 1.16:.1f}" '
                     f'font-size="{sz * PT2PX:.2f}" fill="#{el.get("color", "264180")}" '
                     f'font-family="{el.get("face") or FONT}, PingFang SC, sans-serif">'
                     f'{esc(ln)}</text>')

    o.append("</svg>")
    open(path, "w", encoding="utf-8").write("\n".join(o))
    return path


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("spec")
    ap.add_argument("--pptx")
    ap.add_argument("--svg")
    args = ap.parse_args()
    spec = load_spec(args.spec)
    spec["_path"] = os.path.abspath(args.spec)
    if args.pptx:
        print("PPTX:", build_pptx(spec, args.pptx))
    if args.svg:
        print("SVG :", build_svg(spec, args.svg))
    if not (args.pptx or args.svg):
        print("没给 --pptx 或 --svg，什么都没做")


if __name__ == "__main__":
    main()
