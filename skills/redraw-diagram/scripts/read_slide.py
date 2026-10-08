#!/usr/bin/env python3
"""
read_slide.py —— 把 PPT 里的一页（或页面上的一块）读成 spec.json。

spec 是一份**可复现的中间表示**：位置、颜色、文字、线型、图片，全都在里面。
build_deck.py 只认 spec、不认 pptx —— 这就是"读"和"画"分开的地方。

用法：
  python3 read_slide.py deck.pptx --slide 6 --render page6.png --out spec.json
  python3 read_slide.py deck.pptx --slide 5 --region 747,209,1224,616 \
          --render page5.png --out spec.json --offset=-359,45

--region 是**原页面像素坐标**（和 --render 同一坐标系），用来"只要页面的一部分"。
--offset 是输出画布上的平移；不给且给了 region 时，自动把选区居中。

分类顺序很重要：**先按对象类型分（图片/线/文本框/…），再按几何类型分**。
反过来会把文本框误判成矩形（因为文本框的 prst 也叫 rect）。
"""

import argparse
import os
import sys

from PIL import Image
from pptx import Presentation
from pptx.enum.shapes import MSO_SHAPE_TYPE

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from common import (Xfrm, GroupCtx, emu_to_px, mode_fill, darkest, lum,
                    save_spec, ensure_dir)

A = '{http://schemas.openxmlformats.org/drawingml/2006/main}'
P = '{http://schemas.openxmlformats.org/presentationml/2006/main}'

PRST_KIND = {
    'roundRect': 'roundrect',
    'flowChartAlternateProcess': 'roundrect',
    'rect': 'boxrect',
    'diamond': 'diamond',
    'flowChartDecision': 'diamond',
    'ellipse': 'ellipse',
    'flowChartTerminator': 'ellipse',
}
ARROW_PRSTS = {'upArrow', 'downArrow', 'leftArrow', 'rightArrow', 'upDownArrow',
               'leftRightArrow', 'bentArrow', 'bentUpArrow', 'curvedDownArrow',
               'curvedUpArrow', 'curvedLeftArrow', 'curvedRightArrow',
               'stripedRightArrow', 'notchedRightArrow', 'circularArrow',
               'quadArrow', 'leftRightUpArrow'}
BRACE_PRSTS = {'leftBrace', 'rightBrace'}


def prst_of(sh):
    for e in sh._element.iter():
        if e.tag.endswith('}prstGeom'):
            return e.get('prst') or ''
    return ''


def text_of(sh):
    return sh.text_frame.text.strip() if sh.has_text_frame else ""


def ln_props(sh):
    """线条：宽度（EMU）、虚线样式、**XML 里的显式颜色**。"""
    ln = None
    for e in sh._element.iter():
        if e.tag == A + 'ln':
            ln = e
            break
    if ln is None:
        return 9525, 'solid', None
    w = int(ln.get('w')) if ln.get('w') else 9525
    pd = ln.find(A + 'prstDash')
    dash = pd.get('val') if pd is not None else 'solid'
    col = None
    if ln.find(A + 'noFill') is not None:
        col = 'NONE'
    else:
        sf = ln.find(A + 'solidFill')
        if sf is not None:
            c = sf.find(A + 'srgbClr')
            if c is not None:
                col = c.get('val')
    return w, dash, col


def _spPr(sh):
    for e in sh._element.iter():
        if e.tag.endswith('}spPr'):
            return e
    return None


def solid_of(sh, in_ln=False):
    sp = _spPr(sh)
    if sp is None:
        return None
    holder = sp.find(A + 'ln') if in_ln else sp
    if holder is None:
        return None
    sf = holder.find(A + 'solidFill')
    if sf is None:
        return None
    c = sf.find(A + 'srgbClr')
    return c.get('val') if c is not None else None


def has_nofill(sh):
    sp = _spPr(sh)
    return sp is not None and sp.find(A + 'noFill') is not None


def inset_of(sh):
    tb = sh._element.find(P + 'txBody')
    bp = tb.find(A + 'bodyPr') if tb is not None else None
    if bp is None:
        return (15.0, 7.5, 15.0, 7.5, 't')
    def g(k, d):
        return emu_to_px(int(bp.get(k))) if bp.get(k) else d
    ins = (g('lIns', 15.0), g('tIns', 7.5), g('rIns', 15.0), g('bIns', 7.5),
           bp.get('anchor') or 't')
    # ★ wrap 和 overflow 是两件事：wrap 决定换不换行，overflow 决定溢出怎么办。
    #   原图 bodyPr 是 wrap="square"（要换行）+ overflow="overflow"。
    #   把"不换行"当成默认，字就会被撑出框外。
    return ins + (bp.get('wrap') or 'square',)


def font_of(sh):
    tb = sh._element.find(P + 'txBody')
    if tb is None:
        return 18.0, None
    r = tb.find('.//' + A + 'r')
    if r is None:
        return 18.0, None
    rPr = r.find(A + 'rPr')
    if rPr is None:
        return 18.0, None
    sz = int(rPr.get('sz')) / 100.0 if rPr.get('sz') else 18.0
    face = None
    for tag in ('latin', 'ea'):
        e = rPr.find(A + tag)
        if e is not None and e.get('typeface'):
            face = e.get('typeface')
            break
    return sz, face


def in_region(box, region):
    if region is None:
        return True
    rx, ry, rw, rh = region
    cx, cy = box[0] + box[2] / 2, box[1] + box[3] / 2
    return rx <= cx <= rx + rw and ry <= cy <= ry + rh


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("pptx")
    ap.add_argument("--slide", type=int, required=True, help="第几页（从 1 开始）")
    ap.add_argument("--render", help="这一页的渲染图（150dpi 下与页面像素一致，用来取色）")
    ap.add_argument("--region", help="只要页面的一块：x,y,w,h（原页面像素）")
    ap.add_argument("--offset", help="输出画布上的平移 dx,dy（负数写成 --offset=-359,45）")
    ap.add_argument("--canvas", default="2000x1125")
    ap.add_argument("--out", default="spec.json")
    ap.add_argument("--media-dir", default=None)
    args = ap.parse_args()

    region = [float(v) for v in args.region.split(",")] if args.region else None
    cw, ch = [int(v) for v in args.canvas.lower().split("x")]

    ren = Image.open(args.render).convert("RGB") if args.render else None
    if ren is None:
        print("⚠ 没给 --render：只能取 XML 里的显式颜色，"
              "继承主题的颜色会缺失。", file=sys.stderr)

    prs = Presentation(args.pptx)
    slide = prs.slides[args.slide - 1]
    # ★ 默认目录要带 spec 的名字：否则同一个目录下的两个 spec 会互相覆盖图片
    media_dir = args.media_dir or (os.path.splitext(os.path.abspath(args.out))[0] + ".media")
    ensure_dir(media_dir)

    found = []

    def walk(shapes, ctx):
        for sh in shapes:
            xf = Xfrm(sh._element, A)
            if not xf.ok:
                continue
            box = ctx.box_px(xf)
            if sh.shape_type == MSO_SHAPE_TYPE.GROUP:
                walk(sh.shapes, ctx.child(xf))
                continue
            if not in_region(box, region):
                continue
            found.append({"box": box, "sh": sh, "xf": xf,
                          "type": str(sh.shape_type), "prst": prst_of(sh),
                          "text": text_of(sh)})

    walk(slide.shapes, GroupCtx((0, 0), (1.0, 1.0), (0, 0)))

    if region:
        rx, ry, rw, rh = region
        ox, oy = (cw - rw) / 2 - rx, (ch - rh) / 2 - ry
    elif args.offset:
        ox, oy = [float(v) for v in args.offset.split(",")]
    elif found:
        x0 = min(f["box"][0] for f in found)
        y0 = min(f["box"][1] for f in found)
        x1 = max(f["box"][0] + f["box"][2] for f in found)
        y1 = max(f["box"][1] + f["box"][3] for f in found)
        ox, oy = (cw - (x1 - x0)) / 2 - x0, (ch - (y1 - y0)) / 2 - y0
    else:
        ox, oy = 0, 0

    elements, unsupported = [], []
    for f in found:
        sh, t, sb = f["sh"], f["text"], f["box"]
        b = [sb[0] + ox, sb[1] + oy, sb[2], sb[3]]
        typ, prst = f["type"], f["prst"]
        e = None

        # ---------- 1. 先按对象类型 ----------
        if typ.startswith("PICTURE"):
            i = len([x for x in elements if x["kind"] == "picture"])
            ext = sh.image.ext
            fn = os.path.join(media_dir, f"img{i:02d}.{ext}")
            open(fn, "wb").write(sh.image.blob)
            e = {"kind": "picture", "box": b,
                 "src": os.path.relpath(fn, os.path.dirname(os.path.abspath(args.out)))}

        elif typ.startswith("LINE"):
            w, dash, col = ln_props(sh)
            x1, y1, x2, y2 = b[0], b[1], b[0] + b[2], b[1] + b[3]
            if f["xf"].flipH:
                x1, x2 = x2, x1
            if f["xf"].flipV:
                y1, y2 = y2, y1
            if col is None and ren is not None:
                cands = []
                for k in range(1, 10):
                    c = darkest(ren, sb[0] + sb[2] * k / 10 - 3,
                                sb[1] + sb[3] * k / 10 - 3,
                                sb[0] + sb[2] * k / 10 + 4,
                                sb[1] + sb[3] * k / 10 + 4)
                    if c:
                        cands.append(c)
                col = min(cands, key=lum) if cands else None
            e = {"kind": "line", "p1": [x1, y1], "p2": [x2, y2],
                 "width_pt": round(w / 12700.0, 3), "dash": dash,
                 "color": col if col and col != 'NONE' else "333333",
                 "arrow": "end"}

        elif typ.startswith("TABLE"):
            unsupported.append({"what": "表格", "name": sh.name, "box": b})

        elif typ.startswith("TEXT_BOX"):
            sz, face = font_of(sh)
            ins = inset_of(sh)
            col = darkest(ren, sb[0], sb[1], sb[0] + sb[2],
                          sb[1] + sb[3]) if (ren and t) else None
            e = {"kind": "text", "box": b, "text": t, "font_pt": sz,
                 "face": face, "color": col or "264180",
                 "insets": list(ins[:4]), "anchor": ins[4], "wrap": ins[5]}

        elif typ.startswith("FREEFORM"):
            unsupported.append({"what": "自由曲线图形（图标一类）",
                                "name": sh.name, "box": b})

        # ---------- 2. 再按几何类型 ----------
        elif typ.startswith("AUTO_SHAPE"):
            if prst in BRACE_PRSTS:
                rot = round(f["xf"].rot / 60000.0)
                col = None
                if ren is not None:
                    cx = sb[0] + sb[2] / 2 + (sb[3] / 2) * 0.5
                    cy = sb[1] + sb[3] / 2
                    col = darkest(ren, cx - 4, cy - 4, cx + 5, cy + 5)
                e = {"kind": "brace", "box": b, "rot": rot, "color": col or "404040"}

            elif prst in ARROW_PRSTS:
                fl = solid_of(sh)
                if fl is None and ren is not None:
                    fl = mode_fill(ren, sb)
                e = {"kind": "arrow", "prst": prst, "box": b, "fill": fl or "808080"}

            elif prst in PRST_KIND:
                kind = PRST_KIND[prst]
                if kind == "boxrect" and not t:
                    fl = None if has_nofill(sh) else solid_of(sh)
                    if fl is None and ren is not None:
                        fl = mode_fill(ren, sb)   # 继承主题的底色，只能从渲染图采
                    e = {"kind": "container", "box": b,
                         "line": solid_of(sh, in_ln=True), "fill": fl}
                else:
                    if kind == "boxrect":
                        kind = "rect"
                    fl = solid_of(sh)
                    if ren is not None:
                        mf = mode_fill(ren, sb)
                        if mf:
                            fl = mf
                    sz, face = font_of(sh)
                    fg = darkest(ren, sb[0], sb[1], sb[0] + sb[2],
                                 sb[1] + sb[3]) if (ren and t) else None
                    ins = inset_of(sh)
                    e = {"kind": kind, "box": b, "fill": fl, "text": t,
                         "font_pt": sz, "face": face, "fg": fg or "333333",
                         "insets": list(ins[:4]), "anchor": ins[4], "wrap": ins[5]}
            else:
                unsupported.append({"what": f"几何类型 {prst}",
                                    "name": sh.name, "box": b})

        if e:
            e["_src_box"] = [round(v, 1) for v in sb]
            elements.append(e)

    blanks = [e for e in elements
              if e["kind"] in ("roundrect", "diamond", "ellipse", "rect")
              and not e.get("text")]
    counts = {}
    for e in elements:
        counts[e["kind"]] = counts.get(e["kind"], 0) + 1

    spec = {"spec_version": "0.1",
            "source": {"file": os.path.basename(args.pptx), "slide": args.slide,
                       "region": region, "canvas": [cw, ch], "offset": [ox, oy]},
            "counts": counts, "elements": elements, "unsupported": unsupported}
    save_spec(spec, args.out)

    print(f"读入 {len(elements)} 个元素 -> {args.out}")
    print("  统计:", counts)
    if blanks:
        print(f"  ⚠ {len(blanks)} 个形状**没有文字**。")
        print("     先按「原件故意留空」处理，**不要自己填名字**，")
        print("     除非用户确认那是没画完。见 references/reading-rules.md")
    if unsupported:
        print(f"  ⚠ {len(unsupported)} 个对象暂时画不出来（不会进重画结果）：")
        for u in unsupported[:6]:
            print(f"       - {u['what']}  ({u['name']})")
        if len(unsupported) > 6:
            print(f"       … 还有 {len(unsupported) - 6} 个")


if __name__ == "__main__":
    main()
