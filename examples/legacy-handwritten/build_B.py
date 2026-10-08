#!/usr/bin/env python3
"""
dryrun-02：重画 B 区（第 5 页右侧的四层架构）。

规矩：
  - 元素不增不减。8 个空框（6 roundRect + 2 diamond）**原样留空，一个字不填**——
    已确认那是作者故意的，意思是"这里放你自己定义的步骤"。
  - 6 张图片原样贴回。
  - 颜色从原图渲染里**采样**，不手写：这样主题强调色（黄 #F1B700）自动跟着走。
  - 坐标用修正后的组变换（`abs = off + (child - chOff) * scale`）。

输出：slide5-B-redraw.pptx / slide5-B-redraw.svg
"""
import base64
import io
import json
from PIL import Image, ImageDraw, ImageFont
from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE, MSO_CONNECTOR
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.dml import MSO_LINE_DASH_STYLE
from lxml import etree

DPI = 150.0
W, H = 2000, 1125
SRC_PPTX = "/Users/rio/Desktop/测试架构图.pptx"
SS = 4                      # 第 5 页（0-based）
RENDER = "diagram-redraw/dryrun-02/orig/page-0003.png"
OUT = "diagram-redraw/dryrun-02/"
A = '{http://schemas.openxmlformats.org/drawingml/2006/main}'
P = '{http://schemas.openxmlformats.org/presentationml/2006/main}'


def IN(px):
    return Inches(px / DPI)


# ---------- 1. 原图取色 ----------
ren = Image.open(RENDER).convert("RGB")


def s(x, y):
    p = ren.getpixel((int(x), int(y)))
    return "%02X%02X%02X" % p


def darken(h, f=0.72):
    r, g, b = int(h[0:2], 16), int(h[2:4], 16), int(h[4:6], 16)
    return "%02X%02X%02X" % (int(r * f), int(g * f), int(b * f))


from collections import Counter


def mode_fill(b):
    """框内出现最多的颜色 = 填充色（文字占少数，会被众数压掉）。
    b = [x, y, w, h]"""
    x0, y0 = b[0], b[1]
    x1, y1 = b[0] + b[2], b[1] + b[3]
    c = Counter()
    for yy in range(int(y0) + 4, int(y1) - 4, 2):
        for xx in range(int(x0) + 4, int(x1) - 4, 2):
            if 0 <= xx < W and 0 <= yy < H:
                c[ren.getpixel((xx, yy))] += 1
    if not c:
        return "CCCCCC"
    p = c.most_common(1)[0][0]
    return "%02X%02X%02X" % p


def darkest(x0, y0, x1, y1):
    """区域里最深的像素——用来取文字颜色，比单点采样稳。"""
    best, bv = None, 999
    for yy in range(int(y0), int(y1), 2):
        for xx in range(int(x0), int(x1), 2):
            p = ren.getpixel((xx, yy))
            v = sum(p)
            if v < bv:
                bv, best = v, p
    return "%02X%02X%02X" % best


COL = {"pink": "FEC1CE", "blue": "B6CAFA", "ellipse": "B6CAFA"}
COL["title"] = darkest(753, 304, 1060, 350)
COL["pinkbd"] = darken(COL["pink"], 0.82)
COL["bluebd"] = darken(COL["blue"], 0.80)
COL["accent"] = s(1105, 721)      # 黄色双向箭头
COL["navy"] = "08266E"
COL["txt"] = "333333"
print("采样配色:", COL)

# ---------- 2. 走一遍 pptx，拿绝对框（修正后的变换） ----------
pr = Presentation(SRC_PPTX)
grp = [sh for sh in pr.slides[SS].shapes if sh.shape_type == 6][1]


def xfrm(el):
    for e in el.iter():
        if e.tag == A + 'xfrm':
            o, ex = e.find(A + 'off'), e.find(A + 'ext')
            co, ce = e.find(A + 'chOff'), e.find(A + 'chExt')
            if o is None or ex is None:
                continue
            d = {'off': (int(o.get('x')), int(o.get('y'))),
                 'ext': (int(ex.get('cx')), int(ex.get('cy')))}
            if co is not None and ce is not None:
                d['ch'] = (int(co.get('x')), int(co.get('y')),
                           int(ce.get('cx')), int(ce.get('cy')))
            return d
    return None


g = xfrm(grp._element)
SX = g['ext'][0] / g['ch'][2]
SY = g['ext'][1] / g['ch'][3]
GO = g['off']
CH = g['ch'][:2]
# 目标位置：把 B 居中放到空白页上
BX0, BY0 = GO[0] / 6096, GO[1] / 6096
BW, BH = g['ext'][0] / 6096, g['ext'][1] / 6096
OX, OY = (W - BW) / 2 - BX0, (H - BH) / 2 - BY0
print(f"B 原始框 {BX0:.0f},{BY0:.0f} {BW:.0f}x{BH:.0f} -> 目标偏移 ({OX:+.0f},{OY:+.0f})")


def absbox(sh, parent_off=None, psx=None, psy=None, pch=None):
    x = xfrm(sh._element)
    if x is None:
        return None
    if parent_off is None:
        parent_off, psx, psy, pch = GO, SX, SY, CH
    X = parent_off[0] + (x['off'][0] - pch[0]) * psx
    Y = parent_off[1] + (x['off'][1] - pch[1]) * psy
    w = x['ext'][0] * psx
    h = x['ext'][1] * psy
    return [X / 6096 + OX, Y / 6096 + OY, w / 6096, h / 6096]


def prst(sh):
    for e in sh._element.iter():
        if e.tag.endswith('}prstGeom'):
            return e.get('prst') or ''
    return ''


def runs(sh):
    tb = sh._element.find(P + 'txBody')
    return 0 if tb is None else len(tb.findall('.//' + A + 'r'))


def lnprops(sh):
    ln = None
    for e in sh._element.iter():
        if e.tag == A + 'ln':
            ln = e
            break
    if ln is None:
        return (9525, 'solid', None)
    w = int(ln.get('w')) if ln.get('w') else 9525
    pd = ln.find(A + 'prstDash')
    dash = pd.get('val') if pd is not None else 'solid'
    sf = ln.find(A + 'solidFill')
    col = None
    if sf is not None:
        c = sf.find(A + 'srgbClr')
        if c is not None:
            col = c.get('val')
    return (w, dash, col)


def fillcol(sh):
    spPr = None
    for e in sh._element.iter():
        if e.tag.endswith('}spPr'):
            spPr = e
            break
    if spPr is None:
        return None
    sf = spPr.find(A + 'solidFill')
    if sf is None:
        return None
    c = sf.find(A + 'srgbClr')
    if c is not None:
        return c.get('val')
    sc = sf.find(A + 'schemeClr')
    if sc is not None:
        v = sc.get('val')
        return {"accent3": COL["accent"], "tx1": "000000"}.get(v, "808080")
    return None


ELEMS = {"rects": [], "blanks": [], "boxes": [], "ellipses": [], "diamonds": [],
         "lines": [], "braces": [], "arrows": [], "pics": [], "texts": []}

imgcache = []
for sh in grp.shapes:
    st = str(sh.shape_type).split(' ')[0]
    b = absbox(sh)
    if b is None:
        continue
    p = prst(sh)
    if st == "PICTURE":
        blob = sh.image.blob
        ext = sh.image.ext
        imgcache.append((b, blob, ext))
        ELEMS["pics"].append(b)
    elif st == "LINE":
        w, dash, col = lnprops(sh)
        x = xfrm(sh._element)
        e_ = sh._element
        fh = fv = False
        for el in e_.iter():
            if el.tag == A + 'xfrm':
                fh = el.get('flipH') in ('1', 'true')
                fv = el.get('flipV') in ('1', 'true')
                break
        ELEMS["lines"].append({"box": b, "w": w, "dash": dash, "col": col,
                               "xml_col": col, "flipH": fh, "flipV": fv,
                               "name": sh.name})
    elif st == "TEXT_BOX":
        ELEMS["texts"].append({"box": b, "text": sh.text_frame.text.strip(),
                               "name": sh.name})
    elif st == "AUTO_SHAPE":
        t = sh.text_frame.text.strip() if sh.has_text_frame else ""
        if p in ("leftBrace", "rightBrace"):
            ELEMS["braces"].append({"box": b, "rot": sh.rotation})
        elif "Arrow" in p:
            ELEMS["arrows"].append({"box": b, "prst": p, "fill": fillcol(sh)})
        elif p == "rect":
            ELEMS["rects"].append({"box": b, "name": sh.name})
        elif p == "roundRect":
            if t:
                ELEMS["boxes"].append({"box": b, "text": t, "fill": fillcol(sh)})
            else:
                ELEMS["blanks"].append({"box": b, "prst": p})
        elif p == "diamond":
            ELEMS["blanks"].append({"box": b, "prst": p})
        elif p == "ellipse":
            ELEMS["ellipses"].append({"box": b, "text": t, "fill": fillcol(sh)})

# ---- 用渲染图给每个元素采真实颜色 ----
for e in ELEMS["boxes"]:
    b = e["box"]
    ox, oy = b[0] - OX, b[1] - OY
    e["fill"] = mode_fill([ox, oy, b[2], b[3]])
    e["fg"] = darkest(ox, oy, ox + b[2], oy + b[3])
for e in ELEMS["ellipses"]:
    b = e["box"]
    ox, oy = b[0] - OX, b[1] - OY
    e["fill"] = mode_fill([ox, oy, b[2], b[3]])
    e["fg"] = darkest(ox, oy, ox + b[2], oy + b[3])
COL["blue"] = mode_fill([ELEMS["blanks"][0]["box"][0] - OX, ELEMS["blanks"][0]["box"][1] - OY,
                         ELEMS["blanks"][0]["box"][2], ELEMS["blanks"][0]["box"][3]])
for br in ELEMS["braces"]:
    b = br["box"]
    # 括号旋转过：渲染后的中心就是框中心，长度沿水平方向 = b[3]
    bcx = b[0] - OX + b[2] / 2
    bcy = b[1] - OY + b[3] / 2
    sx = bcx + (b[3] / 2) * 0.5          # 取在横臂上
    br["col"] = darkest(sx - 4, bcy - 4, sx + 5, bcy + 5)
for ar in ELEMS["arrows"]:
    b = ar["box"]
    ar["fill"] = mode_fill([b[0] - OX, b[1] - OY, b[2], b[3]])
for L in ELEMS["lines"]:
    b = L["box"]
    ox, oy = b[0] - OX, b[1] - OY
    # 沿线取几个点（0.3/0.5/0.7），每个点附近找最深，再取其中最深的。
    # 单点会踩到抗锯齿或空白；整个 bbox 又会混进别的东西。
    cands = []
    for f in [i / 10 for i in range(1, 10)]:
        px_, py_ = ox + b[2] * f, oy + b[3] * f
        c = darkest(px_ - 3, py_ - 3, px_ + 4, py_ + 4)
        if c:
            cands.append(c)
    # 原则：原图 XML 显式写了颜色就用它（采样会被穿过的元素污染）；
    # 只有没写（继承主题）时才去渲染图里采。
    L["col"] = L.get("xml_col") or (
        min(cands, key=lambda h: sum(int(h[i:i+2], 16) for i in (0, 2, 4))) if cands else None)
print("采样后：方块填充", [e["fill"] for e in ELEMS["boxes"][:3]],
      "箭头", [a["fill"] for a in ELEMS["arrows"]], "空框", COL["blue"])

print(f"读入：矩形底 {len(ELEMS['rects'])} / 空框 {len(ELEMS['blanks'])} / "
      f"有字方块 {len(ELEMS['boxes'])} / 椭圆 {len(ELEMS['ellipses'])} / "
      f"线 {len(ELEMS['lines'])} / 括号 {len(ELEMS['braces'])} / "
      f"箭头 {len(ELEMS['arrows'])} / 图 {len(ELEMS['pics'])} / 字 {len(ELEMS['texts'])}")
json.dump({k: v for k, v in ELEMS.items()},
          open(OUT + "redraw-B-elements.json", "w"), ensure_ascii=False, indent=1)

# ---------- 3. 自检：空框应该落在浅蓝上，有字方块落在粉上 ----------
print("\n自检（用原图坐标采样）：")
for _bk in ELEMS["blanks"][:3]:
    b = _bk["box"]
    ox, oy = b[0] - OX, b[1] - OY
    print(f"   空框 @原图({ox:.0f},{oy:.0f}) 颜色 #{s(ox + 8, oy + 6)}  (应为 {COL['blue']})")
for e in ELEMS["boxes"][:2]:
    ox, oy = e["box"][0] - OX, e["box"][1] - OY
    print(f"   {e['text'][:10]:12s} @原图({ox:.0f},{oy:.0f}) 颜色 #{s(ox + 8, oy + 6)}  (应为 {COL['pink']})")

# ---------- 4. 出 PPTX ----------
def set_line(sh, color, width_pt, dash=None, head=False, tail=False):
    sh.line.color.rgb = RGBColor.from_string(color)
    sh.line.width = Pt(width_pt)
    if dash:
        sh.line.dash_style = dash
    if head or tail:
        ln = sh.line._get_or_add_ln()
        for tag in ("headEnd", "tailEnd"):
            for e in ln.findall(A + tag):
                ln.remove(e)
        for want, tag in ((head, "headEnd"), (tail, "tailEnd")):
            if want:
                t = etree.SubElement(ln, A + tag)
                t.set("type", "triangle"); t.set("w", "med"); t.set("len", "med")


def add_rect(sl, b, fill=None, line=None, lw=1.0, shape=MSO_SHAPE.RECTANGLE, radius=None):
    sh = sl.shapes.add_shape(shape, IN(b[0]), IN(b[1]), IN(b[2]), IN(b[3]))
    if radius is not None:
        try:
            sh.adjustments[0] = radius
        except Exception:
            pass
    if fill:
        sh.fill.solid(); sh.fill.fore_color.rgb = RGBColor.from_string(fill)
    else:
        sh.fill.background()
    if line:
        sh.line.color.rgb = RGBColor.from_string(line); sh.line.width = Pt(lw)
    else:
        sh.line.fill.background()
    sh.shadow.inherit = False
    return sh


def put_text(sh, text, size, color, align=PP_ALIGN.CENTER, bold=False):
    tf = sh.text_frame
    # 原图是 vertOverflow/horzOverflow = overflow：不强制换行，允许溢出
    tf.word_wrap = False
    # 原图用的是 OOXML 默认内边距：左右 0.1in(15px)，上下 0.05in(7.5px)
    tf.margin_left = tf.margin_right = Emu(91440)
    tf.margin_top = tf.margin_bottom = Emu(45720)
    tf.vertical_anchor = MSO_ANCHOR.MIDDLE
    for i, ln in enumerate(text.split("\n")):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.alignment = align
        r = p.add_run(); r.text = ln
        r.font.size = Pt(size); r.font.bold = bold
        r.font.color.rgb = RGBColor.from_string(color)
        # ★ 中文必须显式设 ea（东亚）字体。只设 latin 的话，中文会回落到主题默认字体，
        #   字宽不一样 -> 偏大、多余换行。原图是 latin + ea 都设微软雅黑。
        r.font.name = "微软雅黑"
        rPr = r._r.get_or_add_rPr()
        # 原图的字体元素带 panose / pitchFamily / charset(GB2312)。
        # 不带这些，中文可能回落到另一个更宽的字体。
        FONT_ATTRS = {"panose": "020B0604030504040204",
                      "pitchFamily": "34", "charset": "-122"}
        for tag in ("latin", "ea", "cs"):
            e = rPr.find(A + tag)
            if e is None:
                e = etree.SubElement(rPr, A + tag)
            e.set("typeface", "微软雅黑")
            for k, v in FONT_ATTRS.items():
                e.set(k, v)


def build_pptx(path):
    prs = Presentation()
    prs.slide_width = IN(W); prs.slide_height = IN(H)
    sl = prs.slides.add_slide(prs.slide_layouts[6])

    # 层底
    for r in ELEMS["rects"]:
        add_rect(sl, r["box"], fill=None, line=COL["navy"], lw=0.5)
    # 空框（保持空白！）
    for bk in ELEMS["blanks"]:
        b = bk["box"]
        if bk["prst"] == "diamond":
            add_rect(sl, b, fill=COL["blue"], line=COL["bluebd"], lw=0.5,
                     shape=MSO_SHAPE.DIAMOND)
        else:
            add_rect(sl, b, fill=COL["blue"], line=COL["bluebd"], lw=0.5,
                     shape=MSO_SHAPE.ROUNDED_RECTANGLE, radius=0.16)
    # 椭圆 开始/结束
    for e in ELEMS["ellipses"]:
        sh = add_rect(sl, e["box"], fill=e["fill"] or COL["ellipse"],
                      shape=MSO_SHAPE.OVAL)
        put_text(sh, e["text"], 8, e["fg"], bold=True)
    # 有字方块
    for e in ELEMS["boxes"]:
        sh = add_rect(sl, e["box"], fill=e["fill"],
                      line=COL["pinkbd"], lw=0.5,
                      shape=MSO_SHAPE.ROUNDED_RECTANGLE, radius=0.22)
        put_text(sh, e["text"], 8, e["fg"])
    # 线
    DASH = {'dash': MSO_LINE_DASH_STYLE.DASH, 'sysDash': MSO_LINE_DASH_STYLE.DASH_DOT,
            'lgDash': MSO_LINE_DASH_STYLE.LONG_DASH}
    for L in ELEMS["lines"]:
        x, y, w, h = L["box"]
        x1, y1, x2, y2 = x, y, x + w, y + h
        if L.get("flipH"):
            x1, x2 = x2, x1
        if L.get("flipV"):
            y1, y2 = y2, y1
        cn = sl.shapes.add_connector(MSO_CONNECTOR.STRAIGHT, IN(x1), IN(y1),
                                     IN(x2), IN(y2))
        L["_p"] = (x1, y1, x2, y2)
        col = L["col"] or COL["navy"]
        set_line(cn, col, max(L["w"] / 12700.0, 0.5), DASH.get(L["dash"]))
        # 箭头（原图大多数带 tailEnd，这里按宽度判断是否流程图连线）
        if 0.7 < L["w"] / 12700.0 < 1.0 and L["dash"] == "solid" and col == COL["navy"]:
            pass
        elif L["dash"] != "solid":
            set_line(cn, col, max(L["w"] / 12700.0, 0.5),
                     DASH.get(L["dash"]), tail=True)
        else:
            set_line(cn, col, max(L["w"] / 12700.0, 0.5), tail=True)
    # 括号（旋转 90/270 -> 横向）
    for br in ELEMS["braces"]:
        b = br["box"]
        cx, cy = b[0] + b[2] / 2, b[1] + b[3] / 2
        sh = sl.shapes.add_shape(MSO_SHAPE.LEFT_BRACE, IN(cx - b[2] / 2),
                                 IN(cy - b[3] / 2), IN(b[2]), IN(b[3]))
        sh.rotation = br["rot"]
        sh.fill.background()
        sh.line.color.rgb = RGBColor.from_string(br.get("col") or COL["navy"]); sh.line.width = Pt(0.5)
        sh.shadow.inherit = False
    # 箭头
    SH = {"upDownArrow": MSO_SHAPE.UP_DOWN_ARROW, "downArrow": MSO_SHAPE.DOWN_ARROW,
          "rightArrow": MSO_SHAPE.RIGHT_ARROW}
    for ar in ELEMS["arrows"]:
        sh = add_rect(sl, ar["box"], fill=ar["fill"] or "000000",
                      shape=SH.get(ar["prst"], MSO_SHAPE.RIGHT_ARROW))
    # 图片
    import tempfile, os
    for i, (b, blob, ext) in enumerate(imgcache):
        p = os.path.join(tempfile.gettempdir(), f"b_img_{i}.{ext}")
        open(p, "wb").write(blob)
        sl.shapes.add_picture(p, IN(b[0]), IN(b[1]), IN(b[2]), IN(b[3]))
    # 文字
    for t in ELEMS["texts"]:
        tb = sl.shapes.add_textbox(IN(t["box"][0]), IN(t["box"][1]),
                                   IN(t["box"][2]), IN(t["box"][3]))
        put_text(tb, t["text"], 10, COL["title"], align=PP_ALIGN.LEFT)
    prs.save(path)
    return path


# ---------- 5. 出 SVG ----------
def esc(s_):
    return s_.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")


def svg_round(x, y, w, h, fill, rx, stroke=None, sw=1.0):
    o = f'<rect x="{x:.0f}" y="{y:.0f}" width="{w:.0f}" height="{h:.0f}" rx="{rx}" fill="#{fill}"'
    if stroke:
        o += f' stroke="#{stroke}" stroke-width="{sw}"'
    return o + "/>"


def build_svg(path):
    out = [f'<svg xmlns="http://www.w3.org/2000/svg" width="{W}" height="{H}" '
           f'viewBox="0 0 {W} {H}">',
           f'<rect width="{W}" height="{H}" fill="#FFFFFF"/>',
           # ★ 上一版漏了这段：16 处 marker-end 全指向不存在的标记 -> 箭头全没了
           '<defs><marker id="ah" viewBox="0 0 10 10" refX="9" refY="5" '
           'markerWidth="6" markerHeight="6" orient="auto-start-reverse">'
           '<path d="M 0 0 L 10 5 L 0 10 z" fill="#1a1a1a"/></marker></defs>']
    for r in ELEMS["rects"]:
        b = r["box"]
        out.append(svg_round(b[0], b[1], b[2], b[3], "FFFFFF", 3, COL["navy"], 1.04))
    for bk in ELEMS["blanks"]:
        b = bk["box"]
        if bk["prst"] == "diamond":
            cx, cy = b[0] + b[2] / 2, b[1] + b[3] / 2
            out.append(f'<polygon points="{cx:.0f},{b[1]:.0f} {b[0]+b[2]:.0f},{cy:.0f} '
                       f'{cx:.0f},{b[1]+b[3]:.0f} {b[0]:.0f},{cy:.0f}" fill="#{COL["blue"]}" '
                       f'stroke="#{COL["bluebd"]}" stroke-width="1"/>')
        else:
            out.append(svg_round(b[0], b[1], b[2], b[3], COL["blue"], 10, COL["bluebd"], 1.04))
    for e in ELEMS["ellipses"]:
        b = e["box"]
        out.append(f'<ellipse cx="{b[0]+b[2]/2:.0f}" cy="{b[1]+b[3]/2:.0f}" '
                   f'rx="{b[2]/2:.0f}" ry="{b[3]/2:.0f}" fill="#{e["fill"] or COL["ellipse"]}"/>')
        out.append(f'<text x="{b[0]+b[2]/2:.0f}" y="{b[1]+b[3]/2+6:.0f}" font-size="16.67" '
                   f'fill="#1F3864" text-anchor="middle" font-family="微软雅黑,PingFang SC,sans-serif">{esc(e["text"])}</text>')
    for e in ELEMS["boxes"]:
        b = e["box"]
        out.append(svg_round(b[0], b[1], b[2], b[3], e["fill"] or COL["pink"], 9,
                             COL["pinkbd"], 1.04))
        out.append(f'<text x="{b[0]+b[2]/2:.0f}" y="{b[1]+b[3]/2+6:.0f}" font-size="16.67" '
                   f'fill="#{COL["txt"]}" text-anchor="middle" font-family="微软雅黑,PingFang SC,sans-serif">{esc(e["text"])}</text>')
    for L in ELEMS["lines"]:
        x1, y1, x2, y2 = L.get("_p", (L["box"][0], L["box"][1],
                                      L["box"][0] + L["box"][2], L["box"][1] + L["box"][3]))
        col = L["col"] or "333333"
        dash = ' stroke-dasharray="6,4"' if L["dash"] != "solid" else ""
        mk = ' marker-end="url(#ah)"' if L["dash"] != "solid" or L["w"] / 12700.0 < 2.0 else ""
        out.append(f'<line x1="{x1:.0f}" y1="{y1:.0f}" x2="{x2:.0f}" y2="{y2:.0f}" '
                   f'stroke="#{col}" stroke-width="{max(L["w"]/12700.0,0.4)*150/72:.2f}"{dash}{mk}/>')
    for br in ELEMS["braces"]:
        b = br["box"]
        # 括号被旋转了 90/270 度：长度是 b[3]（未旋转时的高），不是 b[2]
        cx, cy = b[0] + b[2] / 2, b[1] + b[3] / 2
        brc = br.get("col") or COL["navy"]
        half = b[3] / 2
        x0, x1 = cx - half, cx + half
        dep = (b[2] / 2) * 0.75                       # 尖的深度
        sign = -1 if br["rot"] == 90 else 1           # rot=90 尖朝上；rot=270 尖朝下
        t = max(dep, 6)
        # 原图是又窄又尖的直角 V，不是圆弧
        d = (f'M {x0:.1f} {cy:.1f} L {cx-t:.1f} {cy:.1f} '
             f'L {cx:.1f} {cy+sign*dep:.1f} L {cx+t:.1f} {cy:.1f} L {x1:.1f} {cy:.1f}')
        out.append(f'<path d="{d}" fill="none" stroke="#{brc}" '
                   f'stroke-width="1.04" stroke-linejoin="miter"/>')
    for ar in ELEMS["arrows"]:
        b = ar["box"]
        f_ = ar["fill"] or "000000"
        x, y, w, h = b
        if ar["prst"] == "upDownArrow":
            # 上下双箭头：两个箭头尖 + 中间一根杆。上一版画成了单向箭头。
            cx = x + w / 2
            hh = h * 0.34          # 箭头尖高度
            sw = max(w * 0.42, 2)  # 杆宽
            sx0, sx1 = cx - sw / 2, cx + sw / 2
            out.append(
                f'<path d="M {cx:.1f} {y:.1f} L {x+w:.1f} {y+hh:.1f} L {sx1:.1f} {y+hh:.1f} '
                f'L {sx1:.1f} {y+h-hh:.1f} L {x+w:.1f} {y+h-hh:.1f} L {cx:.1f} {y+h:.1f} '
                f'L {x:.1f} {y+h-hh:.1f} L {sx0:.1f} {y+h-hh:.1f} L {sx0:.1f} {y+hh:.1f} '
                f'L {x:.1f} {y+hh:.1f} Z" fill="#{f_}"/>')
        elif ar["prst"] == "downArrow":
            out.append(f'<path d="M {x+w*0.3:.0f} {y:.0f} L {x+w*0.7:.0f} {y:.0f} L {x+w*0.7:.0f} {y+h*0.6:.0f} '
                       f'L {x+w:.0f} {y+h*0.6:.0f} L {x+w/2:.0f} {y+h:.0f} L {x:.0f} {y+h*0.6:.0f} '
                       f'L {x+w*0.3:.0f} {y+h*0.6:.0f} Z" fill="#{f_}"/>')
        else:
            out.append(f'<path d="M {x:.0f} {y+h*0.3:.0f} L {x+w*0.6:.0f} {y+h*0.3:.0f} L {x+w*0.6:.0f} {y:.0f} '
                       f'L {x+w:.0f} {y+h/2:.0f} L {x+w*0.6:.0f} {y+h:.0f} L {x+w*0.6:.0f} {y+h*0.7:.0f} '
                       f'L {x:.0f} {y+h*0.7:.0f} Z" fill="#{f_}"/>')
    for b, blob, ext in imgcache:
        mime = "image/png" if ext == "png" else f"image/{ext}"
        out.append(f'<image x="{b[0]:.0f}" y="{b[1]:.0f}" width="{b[2]:.0f}" height="{b[3]:.0f}" '
                   f'preserveAspectRatio="none" xlink:href="data:{mime};base64,'
                   f'{base64.b64encode(blob).decode()}" '
                   f'xmlns:xlink="http://www.w3.org/1999/xlink"/>')
    for t in ELEMS["texts"]:
        b = t["box"]
        lines = t["text"].split("\n")
        for i, ln in enumerate(lines):
            out.append(f'<text x="{b[0]+15:.0f}" y="{b[1]+7.5+(i+1)*20:.0f}" font-size="20.83" '
                       f'fill="#{COL["title"]}" font-family="微软雅黑,PingFang SC,sans-serif">{esc(ln)}</text>')
    out.append("</svg>")
    open(path, "w").write("\n".join(out))
    return path


if __name__ == "__main__":
    print("\nPPTX:", build_pptx(OUT + "slide5-B-redraw.pptx"))
    print("SVG :", build_svg(OUT + "slide5-B-redraw.svg"))
