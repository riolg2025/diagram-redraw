#!/usr/bin/env python3
"""
observe_pptx_shapes.py —— 入2：把 PPT 里的一页**清点**成 observations.json。

★ 它产出的是「观察」，不是「清单」。

    观察 = 原件里有什么   —— 机械清点，可复现，**只增不改**（证据）
    清单 = 这张图在说什么 —— 要判断、要归并、要问用户（结论）

    "有 18 个形状" 是观察；"这是一条流程、有 10 条关系" 是结论。
    两者不能混 —— 混了就是拿**画法**属性冒充**意义**属性。
    项目栽过：11 个线条对象，10 条关系，**两个数都对**。

所以这个脚本**不猜**：形状就是形状，线就是线，连接符的端点有就是有、没有就是没有。
归并成"一条关系"是下一步（agent）的事。

用法：
    python3 observe_pptx_shapes.py deck.pptx --slide 6 --out work/observations.json
    python3 observe_pptx_shapes.py deck.pptx --slide 5 --region 747,209,1224,616 \
            --render work/rendered/page-0005.png --out work/observations.json

--region 是**原页面像素坐标**，只用来**筛**（哪些对象属于这一块），**不做坐标变换**。
--render 是**整页**渲染图，用来采"继承主题、XML 里没写"的那些颜色。
"""

import argparse
import json
import os
import sys
from collections import Counter

from PIL import Image
from pptx import Presentation
from pptx.enum.shapes import MSO_SHAPE_TYPE

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from common import Xfrm, GroupCtx, emu_to_px, mode_fill, darkest, lum  # noqa: E402

A = '{http://schemas.openxmlformats.org/drawingml/2006/main}'
P = '{http://schemas.openxmlformats.org/presentationml/2006/main}'

OBS_VERSION = "0.1"

ARROW_PRSTS = {'upArrow', 'downArrow', 'leftArrow', 'rightArrow', 'upDownArrow',
               'leftRightArrow', 'bentArrow', 'bentUpArrow', 'curvedDownArrow',
               'curvedUpArrow', 'curvedLeftArrow', 'curvedRightArrow',
               'stripedRightArrow', 'notchedRightArrow', 'circularArrow',
               'quadArrow', 'leftRightUpArrow'}
BRACE_PRSTS = {'leftBrace', 'rightBrace'}


# --------------------------------------------------------------------------
# XML 小工具
# --------------------------------------------------------------------------
def _first(el, tag):
    for e in el.iter():
        if e.tag == tag:
            return e
    return None


def _first_local(el, local):
    """按**元素名**找，不管命名空间。

    ★ 为什么不能写死命名空间：`<p:spPr>` 在 **presentationml** 里，
      而它的孩子 `<a:solidFill>` 在 **drawingml** 里。
      写成 `a:spPr` 就永远找不到 —— 后果是**所有形状的填充都读不到**，
      重画出来底色全丢，而且不报错。
    """
    for e in el.iter():
        if e.tag.rsplit('}', 1)[-1] == local:
            return e
    return None


def prst_of(el):
    g = _first(el, A + 'prstGeom')
    return (g.get('prst') or '') if g is not None else ''


def text_of(sh):
    try:
        return sh.text_frame.text.strip() if sh.has_text_frame else ""
    except Exception:
        return ""


def sp_pr_of(el):
    """p:sp / p:cxnSp 的 <p:nvSpPr><p:cNvPr>，里面有 id 和 name。"""
    for tag in ('cNvPr',):
        for e in el.iter():
            if e.tag.endswith('}' + tag) and e.get('id'):
                return e
    return None


def cnv_id(el):
    e = sp_pr_of(el)
    return int(e.get('id')) if e is not None else None


def cnv_name(el):
    e = sp_pr_of(el)
    return e.get('name') if e is not None else None


def ln_props(el):
    """线条：宽度、虚线、**XML 里的显式颜色**（没写就是 None，表示继承主题）。"""
    ln = _first(el, A + 'ln')
    if ln is None:
        return {"pt": 0.75, "dash": "solid", "color": None, "src": "default"}
    w = int(ln.get('w')) if ln.get('w') else 9525
    pd = ln.find(A + 'prstDash')
    dash = pd.get('val') if pd is not None else 'solid'
    if ln.find(A + 'noFill') is not None:
        return {"pt": round(w / 12700.0, 3), "dash": dash, "color": None,
                "src": "nofill"}
    col = None
    sf = ln.find(A + 'solidFill')
    if sf is not None:
        c = sf.find(A + 'srgbClr')
        if c is not None:
            col = c.get('val')
        elif sf.find(A + 'schemeClr') is not None:
            # ★ 主题色只能在渲染图上看到长什么样。这里如实记下"它是个主题色"。
            return {"pt": round(w / 12700.0, 3), "dash": dash, "color": None,
                    "src": "scheme",
                    "scheme": sf.find(A + 'schemeClr').get('val')}
    return {"pt": round(w / 12700.0, 3), "dash": dash, "color": col,
            "src": "xml" if col else "unset"}


def fill_of(el):
    """填充：显式色 / 主题色 / 无色 / 没写。"""
    sp = _first_local(el, 'spPr')      # ★ p:spPr，不是 a:spPr
    if sp is None:
        return {"color": None, "src": "nospPr"}
    if sp.find(A + 'noFill') is not None:
        return {"color": None, "src": "nofill"}
    sf = sp.find(A + 'solidFill')
    if sf is None:
        return {"color": None, "src": "unset"}
    c = sf.find(A + 'srgbClr')
    if c is not None:
        return {"color": c.get('val'), "src": "xml"}
    s = sf.find(A + 'schemeClr')
    if s is not None:
        return {"color": None, "src": "scheme", "scheme": s.get('val')}
    return {"color": None, "src": "other"}


def font_of(el):
    tb = _first(el, P + 'txBody')
    if tb is None:
        return {"pt": None, "face": None}
    r = tb.find('.//' + A + 'r')
    rPr = r.find(A + 'rPr') if r is not None else None
    if rPr is None:
        return {"pt": None, "face": None}
    pt = int(rPr.get('sz')) / 100.0 if rPr.get('sz') else None
    face = None
    for tag in ('latin', 'ea'):
        e = rPr.find(A + tag)
        if e is not None and e.get('typeface'):
            face = e.get('typeface')
            break
    return {"pt": pt, "face": face}


def text_color_of(el):
    """文字颜色，**先从 XML 读**。

    ★ 规矩和线条一样：**XML 里显式写了就用它，采样只是渲染的近似。**
      漏了这条的后果实测过：蓝底白字的框，用"框内最深像素"去采，
      采到的是**蓝底本身** → 文字被写成背景色 → **整段字看不见**。
    """
    tb = _first(el, P + 'txBody')
    if tb is None:
        return {"color": None, "src": "none"}
    for tag in ('rPr', 'defRPr'):
        for rPr in tb.iter(A + tag):
            sf = rPr.find(A + 'solidFill')
            if sf is None:
                continue
            c = sf.find(A + 'srgbClr')
            if c is not None:
                return {"color": c.get('val'), "src": "xml"}
            s = sf.find(A + 'schemeClr')
            if s is not None:
                return {"color": None, "src": "scheme", "scheme": s.get('val')}
    return {"color": None, "src": "unset"}


def text_color_sample(img, box, fill_hex, tol=110):
    """兜底：从渲染图里找文字色。

    不能取"最深像素" —— 白字在深底上会被采成底色。
    改成：**框内出现最多的那个"和填充色明显不同"的颜色**。
    文字像素数仅次于底色，所以这条稳。
    """
    from collections import Counter
    x0, y0 = int(box[0]), int(box[1])
    x1, y1 = int(box[0] + box[2]), int(box[1] + box[3])
    cnt = Counter()
    for yy in range(y0 + 2, y1 - 2):
        for xx in range(x0 + 2, x1 - 2):
            if 0 <= xx < img.width and 0 <= yy < img.height:
                cnt[img.getpixel((xx, yy))] += 1
    if not cnt:
        return None
    if not fill_hex:
        return "%02X%02X%02X" % min(cnt, key=lambda p: sum(p))
    fill = tuple(int(fill_hex[i:i + 2], 16) for i in (0, 2, 4))
    for px, _n in cnt.most_common(60):
        if sum(abs(a - b) for a, b in zip(px, fill)) > tol:
            return "%02X%02X%02X" % px
    return None


def insets_of(el):
    tb = _first(el, P + 'txBody')
    bp = tb.find(A + 'bodyPr') if tb is not None else None
    if bp is None:
        return None
    def g(k, d):
        return emu_to_px(int(bp.get(k))) if bp.get(k) else d
    return {"l": g('lIns', 15.0), "t": g('tIns', 7.5),
            "r": g('rIns', 15.0), "b": g('bIns', 7.5),
            "anchor": bp.get('anchor') or 't',
            # ★ wrap 和 overflow 是两件事，别混
            "wrap": bp.get('wrap') or 'square'}


def cxn_ends(el):
    """连接符的端点绑定：<a:stCxn id= idx=/> / <a:endCxn .../>。没有就是 None。"""
    out = {}
    for tag, key in (('stCxn', 'st'), ('endCxn', 'end')):
        e = _first(el, A + tag)
        if e is not None and e.get('id'):
            out[key] = {"shape": int(e.get('id')),
                        "idx": int(e.get('idx')) if e.get('idx') else None}
        else:
            out[key] = None
    return out


def line_pts(box, xf):
    """一条线的两个端点（考虑 flip）。"""
    x0, y0, w, h = box
    x1, y1, x2, y2 = x0, y0, x0 + w, y0 + h
    if xf.flipH:
        x1, x2 = x2, x1
    if xf.flipV:
        y1, y2 = y2, y1
    return [x1, y1], [x2, y2]


# --------------------------------------------------------------------------
def in_region(box, region):
    if region is None:
        return True
    rx, ry, rw, rh = region
    cx, cy = box[0] + box[2] / 2, box[1] + box[3] / 2
    return rx <= cx <= rx + rw and ry <= cy <= ry + rh


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("pptx")
    ap.add_argument("--slide", type=int, required=True)
    ap.add_argument("--render", help="**整页**渲染图，用来采主题色")
    ap.add_argument("--region", help="只要页面上的一块：x,y,w,h（原页面像素，只筛不变换）")
    ap.add_argument("--out", default="observations.json")
    ap.add_argument("--media-dir", help="图片导出目录（默认 <out>.media）")
    args = ap.parse_args()

    region = [float(v) for v in args.region.split(",")] if args.region else None
    ren = Image.open(args.render).convert("RGB") if args.render else None
    if ren is None:
        print("⚠ 没给 --render：XML 里没写颜色（继承主题）的对象会缺颜色。", file=sys.stderr)

    prs = Presentation(args.pptx)
    slide = prs.slides[args.slide - 1]
    canvas = [emu_to_px(prs.slide_width), emu_to_px(prs.slide_height)]

    media_dir = args.media_dir or (os.path.splitext(os.path.abspath(args.out))[0] + ".media")
    os.makedirs(media_dir, exist_ok=True)
    out_dir = os.path.dirname(os.path.abspath(args.out))

    shapes, connectors, groups, unsupported = [], [], [], []
    z = 0
    pic_i = 0

    def walk(shapes_iter, ctx, group_path):
        nonlocal z, pic_i
        for sh in shapes_iter:
            el = sh._element
            xf = Xfrm(el, A)
            if not xf.ok:
                continue
            box = ctx.box_px(xf)
            sid = cnv_id(el)
            oid = "sp%d" % sid if sid is not None else "sp?"
            ev = ["p:%s#%s" % ("cxnSp" if el.tag == P + 'cxnSp' else "sp", sid)]

            if sh.shape_type == MSO_SHAPE_TYPE.GROUP:
                if in_region(box, region):
                    groups.append({
                        "oid": oid, "shape_id": sid, "name": cnv_name(el),
                        "box": [round(v, 1) for v in box],
                        "rot": round(xf.rot / 60000.0),
                        "members": [("sp%d" % cnv_id(c._element))
                                    for c in sh.shapes if cnv_id(c._element) is not None],
                        "evidence": ev,
                    })
                walk(sh.shapes, ctx.child(xf), group_path + [oid])
                continue

            if not in_region(box, region):
                continue

            base = {
                "oid": oid, "shape_id": sid, "name": cnv_name(el),
                "type": str(sh.shape_type), "prst": prst_of(el),
                "box": [round(v, 1) for v in box],
                "rot": round(xf.rot / 60000.0),
                "flipH": xf.flipH, "flipV": xf.flipV,
                "group_path": group_path,
                "z": z,
                "text": text_of(sh),
                "evidence": ev,
            }
            z += 1

            # ---- 连接符：单独一张表 ----
            if el.tag == P + 'cxnSp' or str(sh.shape_type).startswith("LINE"):
                p1, p2 = line_pts(box, xf)
                ends = cxn_ends(el)
                connectors.append(dict(base, kind="connector",
                                       p1=[round(v, 1) for v in p1],
                                       p2=[round(v, 1) for v in p2],
                                       st_cxn=ends["st"], end_cxn=ends["end"],
                                       line=ln_props(el)))
                continue

            # ---- 图片 ----
            if str(sh.shape_type).startswith("PICTURE"):
                try:
                    ext = sh.image.ext
                    fn = os.path.join(media_dir, "img%02d.%s" % (pic_i, ext))
                    with open(fn, "wb") as fh:
                        fh.write(sh.image.blob)
                    src = os.path.relpath(fn, out_dir)
                except Exception as e:
                    src = None
                    unsupported.append(dict(base, what="图片导出失败", why=repr(e)))
                pic_i += 1
                shapes.append(dict(base, kind="picture", src=src))
                continue

            # ---- 表格 / 自由曲线：画不出来，如实记 ----
            if str(sh.shape_type).startswith("TABLE"):
                unsupported.append(dict(base, what="表格", why="不支持"))
                continue
            if str(sh.shape_type).startswith("FREEFORM"):
                unsupported.append(dict(base, what="自由曲线图形（图标一类）",
                                        why="FREEFORM 画不出来"))
                continue

            # ---- 其余：形状 ----
            fl = fill_of(el)
            # 只要没拿到颜色（继承主题 / 没写），就从渲染图上采。
            # 除了 nofill —— 那是**明确说没有填充**，不该给它硬采一个。
            if fl["color"] is None and fl["src"] != "nofill" and ren is not None:
                mf = mode_fill(ren, box)
                if mf:
                    fl = {"color": mf, "src": "sampled"}

            ln = ln_props(el)
            if ln.get("color") is None and ln.get("src") != "nofill" \
                    and ren is not None:
                cands = []
                for k in range(1, 10):
                    cx = box[0] + box[2] * k / 10
                    cy = box[1] + box[3] * k / 10
                    c = darkest(ren, cx - 3, cy - 3, cx + 4, cy + 4)
                    if c:
                        cands.append(c)
                if cands:
                    ln = dict(ln, color=min(cands, key=lum), src="sampled")

            fg = {"color": None, "src": None}
            if base["text"]:
                fg = text_color_of(el)                      # ★ XML 优先
                if fg["color"] is None and fg["src"] != "none" and ren is not None:
                    sampled = text_color_sample(ren, box, fl.get("color"))
                    if sampled:
                        fg = {"color": sampled, "src": "sampled"}

            shape_kind = "text" if str(sh.shape_type).startswith("TEXT_BOX") else "shape"
            if base["prst"] in BRACE_PRSTS:
                shape_kind = "brace"
            elif base["prst"] in ARROW_PRSTS:
                shape_kind = "arrow"

            shapes.append(dict(base, kind=shape_kind,
                               fill=fl, line=ln,
                               font=font_of(el), insets=insets_of(el),
                               fg=fg))

    walk(slide.shapes, GroupCtx((0, 0), (1.0, 1.0), (0, 0)), [])

    textless = [s for s in shapes
                if s["kind"] in ("shape", "arrow") and not s["text"]]
    bound = [c for c in connectors if c["st_cxn"] and c["end_cxn"]]

    obs = {
        "observation_version": OBS_VERSION,
        "source": {
            "path": args.pptx, "slide": args.slide,
            "canvas": [round(v, 1) for v in canvas], "unit": "px@150dpi",
            "region": region, "render": args.render,
        },
        "shapes": shapes,
        "connectors": connectors,
        "groups": groups,
        "unsupported": unsupported,
        "stats": {
            "shapes": len(shapes),
            "connectors": len(connectors),
            "connectors_bound": len(bound),
            "groups": len(groups),
            "textless": len(textless),
            "unsupported": len(unsupported),
        },
    }
    with open(args.out, "w", encoding="utf-8") as f:
        json.dump(obs, f, ensure_ascii=False, indent=1)

    # ---------------- 人读的摘要 ----------------
    st = obs["stats"]
    print("观察到（第 %d 页）：" % args.slide)
    print("  形状 %d   连接符 %d（其中**端点绑定的 %d**）   组合 %d"
          % (st["shapes"], st["connectors"], st["connectors_bound"], st["groups"]))
    print("  没有文字的形状 %d" % st["textless"])
    print("  画不出来的 %d" % st["unsupported"])
    print()
    k = Counter(s["kind"] for s in shapes)
    print("  按类别:", dict(k))
    if connectors:
        print("  连接符端点情况：")
        print("     两端都绑定的  %d   ← 关系在文件里是明写的" % len(bound))
        print("     只绑一端的    %d"
              % len([c for c in connectors if bool(c["st_cxn"]) ^ bool(c["end_cxn"])]))
        print("     两端都没绑的  %d   ← ⚠ 方向和端点只能推，推出来必标「我推的」"
              % len([c for c in connectors if not c["st_cxn"] and not c["end_cxn"]]))
    print()
    print("  ⚠ 这只是**观察**，不是清单。")
    print("     「有 %d 个形状」和「图上有几个东西」**不是一回事**——" % st["shapes"])
    print("     归并成意义单位、判断关系，是下一步（要判断，可能要问用户）。")
    if textless:
        print()
        print("  ⚠ %d 个形状没有文字：先按「原件故意留空」处理，**不要自己填名字**，"
              % len(textless))
        print("     除非用户确认那是没画完。见 references/reading-rules.md")
    if unsupported:
        print()
        print("  画不出来的：")
        for u in unsupported[:6]:
            print("     - %s  (%s)" % (u["what"], u["name"]))
        if len(unsupported) > 6:
            print("     … 还有 %d 个" % (len(unsupported) - 6))
    print()
    print("->", args.out)


if __name__ == "__main__":
    main()
