#!/usr/bin/env python3
"""
build_output.py —— 清单 → 输出（PPTX 可编辑图形 / SVG）。

★ 只认清单，不认原件。 换一个输入，只要能产出同样的清单，这一步不用动。

    spec.layout  ──►  PPTX 可编辑形状 + SVG

为什么两条路各写各的、不做格式转换：
    PPTX 是**声明式**的（说个圆角矩形，位置和文字它帮你摆），
    SVG 是**命令式**的（每根线每行字的坐标都得自己算）。
    "先画 SVG 再转 PPT"实测走不通——转出来是一张死图（见 docs/实测与调研.md）。

★ 预设形状直接用清单里的 `prst` 名字写进 DrawingML，
  不经映射表 —— 这样 OOXML 那 187 个预设形状都能用，
  不用为每一个在代码里加一行。

用法：
    python3 build_output.py work/spec.json --pptx work/out.pptx
    python3 build_output.py work/spec.json --pptx work/out.pptx --svg work/out.svg
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

A = '{http://schemas.openxmlformats.org/drawingml/2006/main}'
DPI = 150.0
PT2PX = DPI / 72.0

# kind → 默认预设几何（清单里给了 prst 就用清单的）
KIND_PRST = {"rect": "rect", "container": "rect", "roundrect": "roundRect",
             "diamond": "diamond", "ellipse": "ellipse"}
BOX_KINDS = ("rect", "container", "roundrect", "diamond", "ellipse",
             "brace", "arrow", "text")

# ★ 中文字体要带这三个属性，否则替代字体可能选得更宽，多出换行
FONT_ATTRS = {"panose": "020B0604030504040204", "pitchFamily": "34",
              "charset": "-122"}
DASH = {"dash": MSO_LINE_DASH_STYLE.DASH,
        "sysDash": MSO_LINE_DASH_STYLE.DASH_DOT,
        "lgDash": MSO_LINE_DASH_STYLE.LONG_DASH,
        "dot": MSO_LINE_DASH_STYLE.ROUND_DOT,
        "sysDot": MSO_LINE_DASH_STYLE.ROUND_DOT}
ALIGN = {"ctr": PP_ALIGN.CENTER, "l": PP_ALIGN.LEFT, "r": PP_ALIGN.RIGHT}
ANCHOR = {"ctr": MSO_ANCHOR.MIDDLE, "t": MSO_ANCHOR.TOP, "b": MSO_ANCHOR.BOTTOM}


def IN(px):
    return Inches(float(px) / DPI)


def load(path):
    import json
    with open(path, encoding="utf-8") as f:
        return json.load(f)


# --------------------------------------------------------------------------
def set_prst(sh, name):
    """把形状的预设几何改成清单里说的那个（Python-pptx 只认枚举，这里直接改 XML）。"""
    if not name:
        return
    for e in sh._element.iter():
        if e.tag == A + 'prstGeom':
            e.set('prst', name)
            return


def style_text(tf, lines, t, face):
    tf.word_wrap = bool(t.get("wrap", True))
    ins = t.get("insets")
    if ins:
        tf.margin_left = Emu(int(ins[0] / DPI * 914400)) if len(ins) > 0 else None
        tf.margin_top = Emu(int(ins[1] / DPI * 914400)) if len(ins) > 1 else None
        tf.margin_right = Emu(int(ins[2] / DPI * 914400)) if len(ins) > 2 else None
        tf.margin_bottom = Emu(int(ins[3] / DPI * 914400)) if len(ins) > 3 else None
    tf.vertical_anchor = ANCHOR.get(t.get("anchor") or "ctr", MSO_ANCHOR.MIDDLE)
    for i, ln in enumerate(lines):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.alignment = ALIGN.get(t.get("align") or "ctr", PP_ALIGN.CENTER)
        run = p.add_run()
        run.text = ln
        if t.get("font_pt"):
            run.font.size = Pt(t["font_pt"])
        if t.get("color"):
            run.font.color.rgb = RGBColor.from_string(t["color"])
        run.font.name = face
        rPr = run._r.get_or_add_rPr()
        for tag in ("latin", "ea", "cs"):
            e = rPr.find(A + tag)
            if e is None:
                e = etree.SubElement(rPr, A + tag)
            e.set("typeface", face)
            for k, v in FONT_ATTRS.items():
                e.set(k, v)


def set_arrow(cn, arrow):
    if arrow not in ("end", "both"):
        return
    ln = cn.line._get_or_add_ln()
    for tag in ("headEnd", "tailEnd"):
        for e in ln.findall(A + tag):
            ln.remove(e)
    for want, tag in ((arrow == "both", "headEnd"), (True, "tailEnd")):
        if want:
            t = etree.SubElement(ln, A + tag)
            t.set("type", "triangle")
            t.set("w", "med")
            t.set("len", "med")


def apply_line(obj, ln):
    """通用描边设置：obj 要有 .line（shape / connector 都有）。"""
    if not ln or not ln.get("color"):
        obj.line.fill.background()
        return
    obj.line.color.rgb = RGBColor.from_string(ln["color"])
    obj.line.width = Pt(ln.get("pt", 0.75))
    if ln.get("dash") and ln["dash"] != "solid":
        obj.line.dash_style = DASH.get(ln["dash"], MSO_LINE_DASH_STYLE.DASH)


# --------------------------------------------------------------------------
def build_pptx(spec, path):
    cw, ch = spec["source"]["canvas"]
    prs = Presentation()
    prs.slide_width, prs.slide_height = IN(cw), IN(ch)
    sl = prs.slides.add_slide(prs.slide_layouts[6])
    spec_dir = os.path.dirname(os.path.abspath(spec["_path"]))

    items = sorted(spec.get("layout") or [], key=lambda s: s.get("z", 0))
    for it in items:
        kind = it["kind"]
        t = it.get("text") or {}
        lines = [str(x) for x in (t.get("lines") or [])]

        # ★ 把清单的 id 写进形状名字 —— 这是"生成时留痕"，
        #   让 Check A 能**逐个对象**对账，而不是只能比"文字集合"。
        def stamp(obj):
            try:
                obj.name = str(it["id"])
            except Exception:
                pass
            return obj

        # ---- 线（连接符）----
        if kind == "line":
            p1, p2 = it["p1"], it["p2"]
            cn = sl.shapes.add_connector(MSO_CONNECTOR.STRAIGHT,
                                         IN(p1[0]), IN(p1[1]), IN(p2[0]), IN(p2[1]))
            stamp(cn)
            apply_line(cn, it.get("line"))
            set_arrow(cn, it.get("arrow"))
            continue

        # ---- 图片 ----
        if kind == "picture":
            b = it["box"]
            src = os.path.join(spec_dir, it["src"])
            stamp(sl.shapes.add_picture(src, IN(b[0]), IN(b[1]), IN(b[2]), IN(b[3])))
            continue

        # ---- 其余：形状 / 文本框 ----
        b = it["box"]
        if kind == "text":
            sh = sl.shapes.add_textbox(IN(b[0]), IN(b[1]), IN(b[2]), IN(b[3]))
        else:
            sh = sl.shapes.add_shape(MSO_SHAPE.RECTANGLE,
                                     IN(b[0]), IN(b[1]), IN(b[2]), IN(b[3]))
            set_prst(sh, it.get("prst") or KIND_PRST.get(kind, "rect"))
            fill = it.get("fill")
            if fill:
                sh.fill.solid()
                sh.fill.fore_color.rgb = RGBColor.from_string(fill)
            else:
                sh.fill.background()
            apply_line(sh, it.get("line"))
            sh.shadow.inherit = False
        stamp(sh)

        if it.get("rot"):
            sh.rotation = it["rot"]
        if lines:
            face = t.get("face") or "微软雅黑"
            if kind == "text" and not t.get("align"):
                t = dict(t, align="l")
            style_text(sh.text_frame, lines, t, face)

    prs.save(path)
    return path


# --------------------------------------------------------------------------
# SVG
# --------------------------------------------------------------------------
def esc(s):
    return (str(s or "").replace("&", "&amp;").replace("<", "&lt;")
            .replace(">", "&gt;"))


def svg_text(x, y, lines, t, face, anchor="middle"):
    size = (t.get("font_pt") or 10) * PT2PX
    out = ['<text x="%.1f" y="%.1f" font-size="%.2f" fill="#%s" text-anchor="%s" '
           'font-family="%s, PingFang SC, Microsoft YaHei, sans-serif">'
           % (x, y, size, t.get("color") or "333333", anchor, face)]
    for i, ln in enumerate(lines):
        out.append('<tspan x="%.1f" dy="%.2f">%s</tspan>'
                   % (x, 0.0 if i == 0 else size * 1.15, esc(ln)))
    out.append("</text>")
    return "".join(out)


def build_svg(spec, path):
    cw, ch = spec["source"]["canvas"]
    spec_dir = os.path.dirname(os.path.abspath(spec["_path"]))
    o = ['<svg xmlns="http://www.w3.org/2000/svg" width="%s" height="%s" '
         'viewBox="0 0 %s %s">' % (cw, ch, cw, ch),
         '<rect width="%s" height="%s" fill="#FFFFFF"/>' % (cw, ch),
         # ★ marker 必须先定义。漏了这句，所有 marker-end 全部失效 → 箭头全没头
         '<defs><marker id="ah" viewBox="0 0 10 10" refX="9" refY="5" '
         'markerWidth="6" markerHeight="6" orient="auto-start-reverse">'
         '<path d="M 0 0 L 10 5 L 0 10 z" fill="context-stroke"/></marker></defs>']

    for it in sorted(spec.get("layout") or [], key=lambda s: s.get("z", 0)):
        kind = it["kind"]
        ln = it.get("line") or {}
        stroke = (' stroke="#%s" stroke-width="%.2f"'
                  % (ln["color"], ln.get("pt", 0.75) * PT2PX)) if ln.get("color") else ""
        if ln.get("dash") and ln["dash"] != "solid":
            stroke += ' stroke-dasharray="7,4"'
        fill = "#%s" % it["fill"] if it.get("fill") else "none"
        t = it.get("text") or {}
        face = t.get("face") or "微软雅黑"

        if kind == "line":
            p1, p2 = it["p1"], it["p2"]
            dash = (' stroke-dasharray="7,4"'
                    if ln.get("dash") and ln["dash"] != "solid" else "")
            mk = ""
            if it.get("arrow") == "end":
                mk = ' marker-end="url(#ah)"'
            elif it.get("arrow") == "both":
                mk = ' marker-start="url(#ah)" marker-end="url(#ah)"'
            o.append('<line x1="%.1f" y1="%.1f" x2="%.1f" y2="%.1f" stroke="#%s" '
                     'stroke-width="%.2f"%s%s/>'
                     % (p1[0], p1[1], p2[0], p2[1], ln.get("color") or "333333",
                        ln.get("pt", 0.75) * PT2PX, dash, mk))
            continue

        if kind == "picture":
            b = it["box"]
            src = os.path.join(spec_dir, it["src"])
            ext = os.path.splitext(src)[1].lstrip('.').lower()
            mime = {"png": "image/png", "jpg": "image/jpeg", "jpeg": "image/jpeg",
                    "gif": "image/gif", "bmp": "image/bmp"}.get(ext, "image/png")
            b64 = base64.b64encode(open(src, "rb").read()).decode()
            o.append('<image x="%.1f" y="%.1f" width="%.1f" height="%.1f" '
                     'preserveAspectRatio="none" xlink:href="data:%s;base64,%s" '
                     'xmlns:xlink="http://www.w3.org/1999/xlink"/>'
                     % (b[0], b[1], b[2], b[3], mime, b64))
            continue

        b = it["box"]
        x, y, w, h = b
        tr = ' transform="rotate(%s %.1f %.1f)"' % (it["rot"], x + w / 2, y + h / 2) \
            if it.get("rot") else ""
        if kind in ("rect", "container"):
            o.append('<rect x="%.1f" y="%.1f" width="%.1f" height="%.1f" fill="%s"%s%s/>'
                     % (x, y, w, h, fill, stroke, tr))
        elif kind == "roundrect":
            o.append('<rect x="%.1f" y="%.1f" width="%.1f" height="%.1f" rx="10" '
                     'fill="%s"%s%s/>' % (x, y, w, h, fill, stroke, tr))
        elif kind == "ellipse":
            o.append('<ellipse cx="%.1f" cy="%.1f" rx="%.1f" ry="%.1f" fill="%s"%s%s/>'
                     % (x + w / 2, y + h / 2, w / 2, h / 2, fill, stroke, tr))
        elif kind == "diamond":
            o.append('<polygon points="%.1f,%.1f %.1f,%.1f %.1f,%.1f %.1f,%.1f" '
                     'fill="%s"%s%s/>'
                     % (x + w / 2, y, x + w, y + h / 2, x + w / 2, y + h, x, y + h / 2,
                        fill, stroke, tr))
        else:
            # brace / arrow / 其它：先按矩形占位，等接上 187 个预设几何的表再补
            o.append('<rect x="%.1f" y="%.1f" width="%.1f" height="%.1f" fill="%s"%s%s/>'
                     % (x, y, w, h, fill, stroke, tr))

        if t.get("lines"):
            lines = [str(v) for v in t["lines"]]
            sz = (t.get("font_pt") or 10) * PT2PX
            al = {"ctr": "middle", "l": "start", "r": "end"}.get(t.get("align") or "ctr",
                                                                 "middle")
            cx = x + w / 2 if al == "middle" else (x + 6 if al == "start" else x + w - 6)
            y0 = y + h / 2 - (len(lines) - 1) * sz * 0.58 + sz * 0.35
            o.append(svg_text(cx, y0, lines, t, face, al))

    o.append("</svg>")
    with open(path, "w", encoding="utf-8") as f:
        f.write("\n".join(o))
    return path


# --------------------------------------------------------------------------
def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("spec")
    ap.add_argument("--pptx")
    ap.add_argument("--svg")
    args = ap.parse_args()
    spec = load(args.spec)
    spec["_path"] = os.path.abspath(args.spec)
    if args.pptx:
        print("PPTX :", build_pptx(spec, args.pptx))
    if args.svg:
        print("SVG  :", build_svg(spec, args.svg))
    if not (args.pptx or args.svg):
        print("没给 --pptx 或 --svg，什么都没做")


if __name__ == "__main__":
    main()
