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
import math
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


def EMU(px):
    return int(float(px) / DPI * 914400)


def _local(el, name):
    """按元素名找（不管命名空间）—— p: 和 a: 都有 cNvPr / xfrm 这类同名元素。"""
    for e in el.iter():
        if e.tag.rsplit('}', 1)[-1] == name:
            return e
    return None


def add_verbatim(sl, it, spec_dir):
    """**原样搬运**：把原件的形状 XML 直接搬进来，一个像素都不改。

    ★ 为什么要有这条：表达不了 ≠ 没法要。
      FREEFORM（图标一类）是 <p:sp> + <a:custGeom>，搬过来**还是原生可编辑形状**，
      完全符合"要可编辑"这条主线。我一开始把它列进"画不出来"，
      那是**拿自己管线的能力当原件的边界**。

    代价要说清楚：自由曲线点开是**一堆节点**，改起来不如预设形状顺手。
    所以清单里它是 kind=verbatim，交付时要告诉用户"这几个是搬过来的"。
    """
    src = os.path.join(spec_dir, it["src"])
    el = etree.parse(src).getroot()

    # 1) 删掉**悬空的关系引用** —— 新文档里没有这些 rId，
    #    留着的话 PowerPoint 打开会报"需要修复"。
    for tag in ('custDataLst', 'hlinkClick', 'hlinkHover', 'audioFile',
                'videoFile', 'media', 'oleObject'):
        for e in list(el.iter()):
            if e.tag.rsplit('}', 1)[-1] == tag:
                e.getparent().remove(e)

    # 2) 换一个不撞的 id，并把名字设成清单的 id（生成时留痕，和别的形状一致）
    spTree = sl.shapes._spTree
    used = [int(e.get("id")) for e in spTree.iter()
            if e.tag.rsplit('}', 1)[-1] == 'cNvPr' and (e.get("id") or "").isdigit()]
    cNvPr = _local(el, 'cNvPr')
    if cNvPr is not None:
        cNvPr.set("id", str((max(used) + 1) if used else 2))
        cNvPr.set("name", str(it["id"]))

    # 3) 坐标换成绝对坐标（组变换在观察那一步已经算好了）
    xfrm = _local(el, 'xfrm')
    if xfrm is not None:
        b = it["box"]
        off, ext = _local(xfrm, 'off'), _local(xfrm, 'ext')
        if off is not None:
            off.set("x", str(EMU(b[0])))
            off.set("y", str(EMU(b[1])))
        if ext is not None:
            ext.set("cx", str(EMU(b[2])))
            ext.set("cy", str(EMU(b[3])))

    spTree.append(el)
    return el


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
    tf.vertical_anchor = ANCHOR.get(t.get("anchor") or "t", MSO_ANCHOR.TOP)
    for i, ln in enumerate(lines):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.alignment = ALIGN.get(t.get("align") or "ctr", PP_ALIGN.CENTER)
        # 行距 / 段前距：原图是 120% 行距，不写就按 100% 渲染，整块文字会挪位
        pPr = p._pPr if p._pPr is not None else p._p.get_or_add_pPr()
        if t.get("lnspc"):
            e = etree.SubElement(pPr, A + 'lnSpc')
            etree.SubElement(e, A + 'spcPct').set('val', str(int(t["lnspc"])))
        if t.get("spc_before") is not None:
            e = etree.SubElement(pPr, A + 'spcBef')
            etree.SubElement(e, A + 'spcPts').set('val', str(int(t["spc_before"])))
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


def set_arrow(cn, arrow, atype="arrow"):
    """★ 箭头类型要用**原件的**。

    硬编码 type="triangle" 实测是个坑：原图是 type="arrow"（细长箭头），
    改成 triangle（实心三角）后**头部明显大一圈**。
    而且原图没写 w / len（用默认），我们显式写 med 也会让它更大。
    """
    if arrow not in ("end", "both"):
        return
    ln = cn.line._get_or_add_ln()
    for tag in ("headEnd", "tailEnd"):
        for e in ln.findall(A + tag):
            ln.remove(e)
    t = (atype or "arrow")
    if t and t != "none":
        if arrow == "both":
            etree.SubElement(ln, A + 'headEnd').set("type", t)
        etree.SubElement(ln, A + 'tailEnd').set("type", t)


def apply_line(obj, ln):
    """通用描边设置：obj 要有 .line（shape / connector 都有）。"""
    if not ln or not ln.get("color"):
        obj.line.fill.background()
        return
    obj.line.color.rgb = RGBColor.from_string(ln["color"])
    # 透明度：原图是 alpha=91%，不写就是 100%（实测线会更"实"）
    a = ln.get("alpha")
    if a is not None and int(a) < 100000:
        clr = obj.line.color._xFill.find(A + 'srgbClr')
        if clr is not None:
            etree.SubElement(clr, A + 'alpha').set('val', str(int(a)))
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
            # ★ 折线要保留原件的预设几何（bentConnector2/3…），
            #   否则一律画成两点直线，拐弯全丢。
            set_prst(cn, it.get("prst"))
            stamp(cn)
            apply_line(cn, it.get("line"))
            set_arrow(cn, it.get("arrow"), it.get("arrow_type") or "arrow")
            # ★ 连接符的旋转。漏了它，折线会拐反、两端也错。
            if it.get("rot"):
                cn.rotation = it["rot"]
            continue

        # ---- 原样搬运 ----
        if kind == "verbatim":
            add_verbatim(sl, it, spec_dir)
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

        # ★ 填充和描边要**对所有 kind** 生效，包括 text。
        #   实测栽过：原来给 text 单开一个分支、把这两样整个跳过，
        #   结果 9 个说明标签的浅灰底（F0F0F0）在 PPTX 里全丢了 ——
        #   而 SVG 那边是无条件画的，于是成了"**两个出口不一致**"：
        #   SVG 有、PPTX 没有。用户一眼就看出来了。
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


def tr_of(it):
    b = it["box"]
    return (' transform="rotate(%s %.1f %.1f)"' % (it["rot"], b[0] + b[2] / 2,
                                                   b[1] + b[3] / 2)
            if it.get("rot") else "")


def svg_text(x, y, lines, t, face, anchor="middle", ident=None):
    size = (t.get("font_pt") or 10) * PT2PX
    out = ['<text%s x="%.1f" y="%.1f" font-size="%.2f" fill="#%s" text-anchor="%s" '
           'font-family="%s, PingFang SC, Microsoft YaHei, sans-serif">'
           % (' id="%s"' % ident if ident else "", x, y, size,
              t.get("color") or "333333", anchor, face)]
    for i, ln in enumerate(lines):
        out.append('<tspan x="%.1f" dy="%.2f">%s</tspan>'
                   % (x, 0.0 if i == 0 else size * 1.15, esc(ln)))
    out.append("</text>")
    return "".join(out)


# ---- SVG：连接符的路径 ----------------------------------------------------
# 定义抄自 OOXML 的权威数据（refs/ppt-master/psd.xml），不是我编的：
#   straightConnector1  (l,t) → (r,b)
#   bentConnector2      (l,t) → (r,t) → (r,b)
#   bentConnector3      (l,t) → (x1,t) → (x1,b) → (r,b)    x1 = w·adj1/100000
CONN_PTS = {
    "straightConnector1": lambda w, h, a: [(0, 0), (w, h)],
    "bentConnector2": lambda w, h, a: [(0, 0), (w, 0), (w, h)],
    "bentConnector3": lambda w, h, a: [(0, 0), (w * a, 0), (w * a, h), (w, h)],
}


def svg_map_pts(pts, x, y, w, h, flip_h, flip_v, rot):
    """把局部路径点套上翻转和旋转，落到画布坐标。

    ★ rot 必须算。实测栽过：原件那条折线是 rot=270°，
      不算的话折线拐反，**两端也落在错的地方** —— 因为旋转过的连接符，
      盒子四角根本不是端点。我一开始就是拿盒子四角当端点，所以"两端"也错了。
    """
    cx, cy = x + w / 2.0, y + h / 2.0
    r = math.radians(rot or 0)
    co, si = math.cos(r), math.sin(r)
    out = []
    for px, py in pts:
        if flip_h:
            px = w - px
        if flip_v:
            py = h - py
        px, py = x + px, y + py
        if rot:
            dx, dy = px - cx, py - cy
            px, py = cx + dx * co - dy * si, cy + dx * si + dy * co
        out.append((px, py))
    return out


def conn_box(p1, p2):
    """从两端推出盒子 + 翻转 —— 和 python-pptx 的 add_connector 同一套算法。"""
    x, y = min(p1[0], p2[0]), min(p1[1], p2[1])
    w, h = abs(p2[0] - p1[0]), abs(p2[1] - p1[1])
    return x, y, w, h, p1[0] > p2[0], p1[1] > p2[1]


def drawingml_path_to_svg(el, box):
    """`<a:custGeom>` 的 pathLst → SVG path 的 d。

    这是「原样搬运」在 SVG 侧要的东西。**认不出的命令要报出来，不静默丢。**
    """
    A_ = '{http://schemas.openxmlformats.org/drawingml/2006/main}'
    pl = next((e for e in el.iter() if e.tag == A_ + 'pathLst'), None)
    if pl is None:
        return None, ["没有 custGeom 路径"]
    path = next((e for e in pl.iter() if e.tag == A_ + 'path'), None)
    if path is None:
        return None, ["pathLst 里没有 path"]
    pw = float(path.get('w') or 0) or 1.0
    ph = float(path.get('h') or 0) or 1.0
    bx, by, bw, bh = box
    sx, sy = bw / pw, bh / ph
    d, warns = [], []

    def pt(e):
        return (bx + float(e.get('x') or 0) * sx, by + float(e.get('y') or 0) * sy)

    for cmd in path:
        tag = cmd.tag.rsplit('}', 1)[-1]
        ps = [pt(c) for c in cmd if c.tag.rsplit('}', 1)[-1] == 'pt']
        if tag == 'moveTo' and ps:
            d.append("M %.1f %.1f" % ps[0])
        elif tag == 'lnTo' and ps:
            d.append("L %.1f %.1f" % ps[0])
        elif tag == 'cubicBezTo' and len(ps) == 3:
            d.append("C " + " ".join("%.1f %.1f" % p for p in ps))
        elif tag == 'quadBezTo' and len(ps) == 2:
            d.append("Q " + " ".join("%.1f %.1f" % p for p in ps))
        elif tag == 'close':
            d.append("Z")
        else:
            warns.append("路径命令 %s 还没实现" % tag)
    return (" ".join(d) if d else None), warns


def svg_fill_of(el):
    """搬运过来的形状在 SVG 里用什么填充（颜色在观察那步已解析成绝对色）。"""
    A_ = '{http://schemas.openxmlformats.org/drawingml/2006/main}'
    sf = next((e for e in el.iter() if e.tag == A_ + 'solidFill'), None)
    if sf is None:
        return "none"
    c = next((e for e in sf.iter() if e.tag == A_ + 'srgbClr'), None)
    if c is None:
        return "none"
    a = next((e for e in c if e.tag == A_ + 'alpha'), None)
    if a is not None:
        return "#%s" % c.get('val')
    return "#%s" % c.get('val')


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

    svg_warn = []
    for it in sorted(spec.get("layout") or [], key=lambda s: s.get("z", 0)):
        kind = it["kind"]
        iid = it.get("id") or ""
        ln = it.get("line") or {}
        stroke = (' stroke="#%s" stroke-width="%.2f"'
                  % (ln["color"], ln.get("pt", 0.75) * PT2PX)) if ln.get("color") else ""
        if ln.get("dash") and ln["dash"] != "solid":
            stroke += ' stroke-dasharray="7,4"'
        fill = "#%s" % it["fill"] if it.get("fill") else "none"
        t = it.get("text") or {}
        face = t.get("face") or "微软雅黑"

        if kind == "line":
            # ★ 折线要按**预设几何 + 翻转 + 旋转**画成真路径。
            #   以前这里一律写 <line>，两条折线全被画成了斜线 —— 而 Check A
            #   当时只看 PPTX，SVG 错了没人知道。
            p1, p2 = it["p1"], it["p2"]
            bx, by, bw, bh, fh, fv = conn_box(p1, p2)
            prst = it.get("prst") or "straightConnector1"
            gen = CONN_PTS.get(prst)
            if gen is None:
                svg_warn.append("%s：连接符预设几何 %s 没实现，先按直线画"
                                % (iid, prst))
                gen = CONN_PTS["straightConnector1"]
            pts = svg_map_pts(gen(bw, bh, 0.5), bx, by, bw, bh, fh, fv,
                              it.get("rot"))
            d = "M " + " L ".join("%.1f %.1f" % p for p in pts)
            mk = ""
            if it.get("arrow") == "end":
                mk = ' marker-end="url(#ah)"'
            elif it.get("arrow") == "both":
                mk = ' marker-start="url(#ah)" marker-end="url(#ah)"'
            # 注意：stroke 属性已经在这条里拼好了，**不能再把 stroke 变量塞进来**
            # （那样会出现两个 stroke 属性 → SVG 不合法 → 浏览器整页报错）
            dash_attr = (' stroke-dasharray="7,4"'
                         if ln.get("dash") and ln["dash"] != "solid" else "")
            o.append('<path id="%s" d="%s" fill="none" stroke="#%s" '
                     'stroke-width="%.2f"%s%s/>'
                     % (iid, d, ln.get("color") or "333333",
                        ln.get("pt", 0.75) * PT2PX, dash_attr, mk))
            continue

        if kind == "verbatim":
            # ★ 搬运过来的形状，SVG 侧也要搬 —— 以前没有这个分支，
            #   8 个图标落进了 else，被画成了矩形。
            el = etree.parse(os.path.join(spec_dir, it["src"])).getroot()
            d, warns = drawingml_path_to_svg(el, it["box"])
            svg_warn += ["%s：%s" % (iid, w) for w in warns]
            if d:
                o.append('<path id="%s" d="%s" fill="%s" stroke="none"%s/>'
                         % (iid, d, svg_fill_of(el), tr_of(it)))
            else:
                svg_warn.append("%s：搬运项在 SVG 里没画出来" % iid)
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
        tr = tr_of(it)
        if kind in ("rect", "container"):
            o.append('<rect id="%s" x="%.1f" y="%.1f" width="%.1f" height="%.1f" '
                     'fill="%s"%s%s/>' % (iid, x, y, w, h, fill, stroke, tr))
        elif kind == "roundrect":
            o.append('<rect id="%s" x="%.1f" y="%.1f" width="%.1f" height="%.1f" '
                     'rx="10" fill="%s"%s%s/>'
                     % (iid, x, y, w, h, fill, stroke, tr))
        elif kind == "ellipse":
            o.append('<ellipse id="%s" cx="%.1f" cy="%.1f" rx="%.1f" ry="%.1f" '
                     'fill="%s"%s%s/>'
                     % (iid, x + w / 2, y + h / 2, w / 2, h / 2, fill, stroke, tr))
        elif kind == "diamond":
            o.append('<polygon id="%s" points="%.1f,%.1f %.1f,%.1f %.1f,%.1f %.1f,%.1f" '
                     'fill="%s"%s%s/>'
                     % (iid, x + w / 2, y, x + w, y + h / 2, x + w / 2, y + h, x,
                        y + h / 2, fill, stroke, tr))
        elif kind != "text":
            # brace / arrow / 其它预设几何：还没接上 187 个那个表 —— **报出来，不静默**
            # （text 不走这里报警：文本框本来就是"透明矩形 + 单独的 <text>"）
            svg_warn.append("%s：kind=%s 在 SVG 里按矩形占位" % (iid, kind))
        if True:
            o.append('<rect id="%s" x="%.1f" y="%.1f" width="%.1f" height="%.1f" '
                     'fill="%s"%s%s/>' % (iid, x, y, w, h, fill, stroke, tr))

        if t.get("lines"):
            lines = [str(v) for v in t["lines"]]
            sz = (t.get("font_pt") or 10) * PT2PX
            al = {"ctr": "middle", "l": "start", "r": "end"}.get(t.get("align") or "ctr",
                                                                 "middle")
            cx = x + w / 2 if al == "middle" else (x + 6 if al == "start" else x + w - 6)
            y0 = y + h / 2 - (len(lines) - 1) * sz * 0.58 + sz * 0.35
            o.append(svg_text(cx, y0, lines, t, face, al, iid))

    o.append("</svg>")
    with open(path, "w", encoding="utf-8") as f:
        f.write("\n".join(o))
    # ★ 不静默：SVG 侧画不了 / 降级的地方全部报出来。
    #   实测栽过：8 个图标在 SVG 里变成了矩形，两个检查都看不见。
    for w_ in svg_warn:
        print("  ⚠ SVG 降级：%s" % w_, file=sys.stderr)
    return path, svg_warn


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
        p_, warns_ = build_svg(spec, args.svg)
        print("SVG  :", p_)
        print("       SVG 降级 %d 处" % len(warns_))
    if not (args.pptx or args.svg):
        print("没给 --pptx 或 --svg，什么都没做")


if __name__ == "__main__":
    main()
