#!/usr/bin/env python3
"""
check_b.py —— Check B：**清单 vs 原件**。

它查的是「**我读错了没有**」。Check A 查的是"执行有没有走样"，
两件事完全不同：

    Check A  输出 vs 清单   —— 我说要画什么，画出来了吗、画对了吗
    Check B  清单 vs 原件   —— 我说原件是这样，原件真的是这样吗

★ 为什么必须**直接读原文件**，不能拿"观察"当基准：
  草稿本身就是从观察生成的，拿观察比 = 自己和自己比，永远通过。
  只有回到原文件重新提取一遍，才是独立的一双眼睛。

它只做**能机器做的那部分**：
    ✅ 文字      —— 直接从 PPTX 里提取
    ❌ 关系      —— 要从几何推，机器做不了，得靠**另一双眼睛**反读

两类结果，严重性不同：
    ① 编造（**失败**）：清单说"原件写着 X"，而原件里没有
       → 这是"我编了一个东西，还标成是原件的"，最危险
    ② 遗漏（**警告**）：原件里有 X，清单里哪儿都没有
       → 可能是漏了，也可能是**有意不画**（比如画不出来的形状里的字）

用法：
    python3 check_b.py work/spec.json repo/examples/deck/测试架构图.pptx --slide 6
    python3 check_b.py work/spec.json deck.pptx --slide 6 --strict   # 遗漏也当失败
"""

import argparse
import json
import os
import re
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from common import Xfrm, GroupCtx  # noqa: E402

A = '{http://schemas.openxmlformats.org/drawingml/2006/main}'

# 「来自原件」的来源档位 —— 这些的文字**必须**在原文件里找得到。
# inferred / domain / user 不在其中：那几类本来就不承诺"原件里写着了"。
FROM_ORIGINAL = ("original", "measured", "composed")


def load(path):
    import json as j
    with open(path, encoding="utf-8") as f:
        return j.load(f)


def norm(s):
    """规范化：去掉所有空白。比的是字符序列，不是分段方式。"""
    return re.sub(r"\s+", "", str(s or ""))


# --------------------------------------------------------------------------
def walk_abs(shapes, ctx=None):
    """递归遍历，**带上组坐标变换**，产出 (shape, 页面像素框)。

    ★ 不能直接用 sh.left：组内形状给的是**组内相对坐标**。
      实测踩过：拿它做区域过滤，把组里的东西全筛掉了（"虚拟机 Agent" 全在组里）。
      项目自己也记过同一个坑：组坐标要减 chOff。
    """
    from pptx.enum.shapes import MSO_SHAPE_TYPE
    ctx = ctx or GroupCtx((0, 0), (1.0, 1.0), (0, 0))
    for sh in shapes:
        xf = Xfrm(sh._element, A)
        if not xf.ok:
            continue
        box = ctx.box_px(xf)
        if sh.shape_type == MSO_SHAPE_TYPE.GROUP:
            for sub in walk_abs(sh.shapes, ctx.child(xf)):
                yield sub
        else:
            yield sh, box


def extract_pptx_text(path, slide_no, region=None):
    """**独立**地从原 pptx 里取文字（不经过 observe 那套代码）。

    region —— 只取**范围之内**的文字。用户确认"只重画左边那块"之后，
    右上文字卡、标题这些本来就不该被算成"遗漏"。
    """
    from pptx import Presentation
    DPI = 150.0
    prs = Presentation(path)
    blocks = []
    slides = [prs.slides[slide_no - 1]] if slide_no else list(prs.slides)
    for sl in slides:
        for sh, box in walk_abs(sl.shapes):
            # ★ 表格/图表**没有 text_frame** ✗ —— 原来这里一句 continue 就把它们
            #   全跳过了，于是"表格拼出来的图"在原件那一侧**等于没有文字**：
            #   观察工具漏了（已修：记进 unsupported），Check B 也漏了 —— 两头都瞎 ✗
            #   实测第 1 页：32 个形状里有 **10 个表格、35 个非空单元格**。
            if getattr(sh, "has_table", False) and sh.has_table:
                if region:
                    cx, cy = box[0] + box[2] / 2, box[1] + box[3] / 2
                    if not (region[0] <= cx <= region[0] + region[2]
                            and region[1] <= cy <= region[1] + region[3]):
                        continue
                cells = [c.text for r in sh.table.rows for c in r.cells]
                t = "\n".join(x for x in cells if norm(x))
                if norm(t):
                    blocks.append(t)
                continue
            if not sh.has_text_frame:
                continue
            if region:
                cx, cy = box[0] + box[2] / 2, box[1] + box[3] / 2
                if not (region[0] <= cx <= region[0] + region[2]
                        and region[1] <= cy <= region[1] + region[3]):
                    continue
            t = sh.text_frame.text
            if norm(t):
                blocks.append(t)
    # SmartArt 的文字不在 shape 里，在 ppt/diagrams/data*.xml
    import zipfile
    from lxml import etree
    try:
        z = zipfile.ZipFile(path)
        for nm in z.namelist():
            if not re.match(r"ppt/diagrams/data\d+\.xml$", nm):
                continue
            root = etree.fromstring(z.read(nm))
            txt = "".join(t.text or "" for t in root.iter(A + 't'))
            if norm(txt):
                blocks.append(txt)
    except Exception:
        pass
    return blocks


def spec_texts(spec):
    """清单里所有"声称有文字"的条目：(id, text, provenance)；外加用户改过的。

    ★ 层 1（entities）和层 2（layout）**都要看**。
      实测漏过：绿色条的标题「自动化运维平台能力通道」挂在 layout 上
      （它归属一个分组，不是实体），只查 entities 就会把它误报成"遗漏"。

    ★★ 但对账**只用 `text`（原件文字），绝不回退到 `text_user`**。
      原先写的是 `e.get("text") or e.get("text_user")` ✗ ——
      原件文字为空时会拿**用户的更正**去跟原件比，把用户的修改当成"我编的"。
      「我读对了没有」和「用户改了什么」是**两件事**，分开报。
    """
    out, corrected = [], []
    for e in spec.get("entities") or []:
        t = e.get("text")                      # ★ 只取原件文字
        if t:
            out.append((e.get("id"), t, e.get("provenance")))
        if e.get("text_user") and e["text_user"] != t:
            corrected.append((e.get("id"), t, e["text_user"]))   # 单独记
    for it in spec.get("layout") or []:
        tx = it.get("text") or {}
        t = "".join(str(x) for x in (tx.get("lines") or []))
        prov = tx.get("prov") or it.get("provenance")
        # ★ prov=user 的文字是**用户改的**，不是"我读的" → 不参与对账
        if t and prov != "user":
            out.append((it.get("id"), t, prov))
    # 同一个字符串（实体和它的 layout 各一份）只留一条，免得重复报
    seen, uniq = set(), []
    for i, t, p in out:
        k = norm(t)
        if k in seen:
            continue
        seen.add(k)
        uniq.append((i, t, p))
    return uniq, corrected


# --------------------------------------------------------------------------
def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("spec")
    ap.add_argument("original", help="原文件（pptx）。图片输入做不了 Check B，会明说")
    ap.add_argument("--slide", type=int, default=None, help="只比这一页（1 起）")
    ap.add_argument("--strict", action="store_true", help="遗漏也算失败")
    args = ap.parse_args()

    spec = load(args.spec)

    ext = args.original.rsplit(".", 1)[-1].lower()
    if ext not in ("pptx", "docx"):
        print("Check B：**做不了** —— 原件是 %s，没有文字层。" % ext)
        print("  回原件对账这条路对图片输入不成立（只能靠 OCR，那又是同一双眼睛）。")
        print("  → 这一项要交给「另一双眼睛」：给一个不共享上下文的读者反读输出。")
        return 0

    # 范围从清单里读 —— 用户确认过的那个范围
    region = None
    for f in (spec.get("source") or {}).get("files") or []:
        if f.get("region"):
            region = f["region"]
            break
    blocks = extract_pptx_text(args.original, args.slide, region)
    if region:
        print("Check B：清单 vs 原件（只看清单里声明的范围 %s）" % (region,))
    else:
        print("Check B：清单 vs 原件")
    blob = norm("".join(blocks))
    block_set = {norm(b) for b in blocks}

    items, corrected = spec_texts(spec)
    fabricated, spliced = [], []
    for sid, text, prov in items:
        if prov not in FROM_ORIGINAL:
            continue
        n = norm(text)
        if n in block_set:
            continue
        if n and n in blob:
            spliced.append((sid, text, prov))
        else:
            fabricated.append((sid, text, prov))

    spec_blob = norm("".join(t for _i, t, _p in items))
    missing = []
    for b in blocks:
        n = norm(b)
        if n and n not in spec_blob:
            missing.append(b)

    if corrected:
        print("  ★ 用户改过的文字   %d（**不是对账错误**，分开列）" % len(corrected))
        for cid, old, new in corrected:
            print("       %-8s 原件 %r → 用户改成 %r" % (cid, old, new))
    print("  原件文字块        %d" % len(blocks))
    print("  清单里带文字的     %d（其中标「来自原件」的 %d）"
          % (len(items),
             len([1 for _i, _t, p in items if p in FROM_ORIGINAL])))
    print()

    if fabricated:
        print("✗ ① 编造 %d 处 —— 清单说是原件的，原件里找不到：" % len(fabricated))
        for sid, text, prov in fabricated:
            print("   %-10s [%s] %r" % (sid, prov, text[:50]))
        print()
    if spliced:
        print("⚠ ② 拼接 %d 处 —— 在原件里找得到，但不是**完整的一块**：" % len(spliced))
        for sid, text, prov in spliced:
            print("   %-10s [%s] %r" % (sid, prov, text[:50]))
        print("   （跨几个字段/几处证据拼出来的，来源档位应该是 composed 才对）")
        print()
    if missing:
        print("⚠ ③ 遗漏 %d 块 —— 原件里有，清单里哪儿都没有：" % len(missing))
        for b in missing[:10]:
            print("   %r" % norm(b)[:60])
        if len(missing) > 10:
            print("   … 还有 %d 块" % (len(missing) - 10))
        print("   （可能是漏了，也可能是**有意不画**——比如画不出来的形状里的字）")
        print()

    if not (fabricated or spliced or missing):
        print("✓ 文字这一块，清单和原件对得上")
    print()
    print("  ⚠ 关系这一项 Check B **做不了**（要从几何推，机器不行）。")
    print("     它只能交给「另一双眼睛」：让一个不共享上下文的读者反读输出。")

    if fabricated:
        return 1
    if args.strict and (spliced or missing):
        return 1
    return 0


if __name__ == "__main__":
    sys.exit(main())
