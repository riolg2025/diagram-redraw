#!/usr/bin/env python3
"""
mark_slide.py —— 生成「标号图」：把读到的东西标在原图上，给用户确认。

**这是整个流程的第一句话。** 不是"我看到了 18 个节点"这种文字，
而是一张标了号的图 + 一句话："这些我都看到了，对吗？"

为什么这样做（实测有效）：
  * 用户**扫一眼就能答**（"对"/"少了一个"），不用想、不用打字、不用描述风格
  * 它把"AI 想当然"的风险**摁死在第一步**
  * 它对"看不懂技术文档"的用户友好 —— 图片谁都会看

用法：
  python3 mark_slide.py spec.json orig.png --out marked.png
"""

import argparse
import os
import sys

from PIL import Image, ImageDraw, ImageFont

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from common import load_spec

BOLD = "/System/Library/Fonts/Supplemental/Arial Bold.ttf"
CN = "/System/Library/Fonts/STHeiti Medium.ttc"

# 每类一个颜色，和输出里的说明一一对应
COLOR = {
    "node": (206, 38, 38),        # 红 = 方框/节点
    "line": (222, 122, 12),       # 橙 = 连线
    "container": (21, 94, 190),   # 蓝 = 分区/容器
    "picture": (200, 40, 40),     # 红框 = 图片
    "text": (110, 110, 110),      # 灰 = 文字
}
NODE_KINDS = ("rect", "roundrect", "diamond", "ellipse")


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("spec")
    ap.add_argument("render", help="这一页的原图（和 spec 同一坐标系）")
    ap.add_argument("--out", required=True)
    ap.add_argument("--offset", default=None,
                    help="spec 是基于原页面坐标时，标号图不用平移（默认按 spec 的 offset 反向回推）")
    args = ap.parse_args()

    spec = load_spec(args.spec)
    ox, oy = spec["source"]["offset"]

    im = Image.open(args.render).convert("RGB")
    d = ImageDraw.Draw(im)
    try:
        f = ImageFont.truetype(BOLD, 17)
    except Exception:
        f = ImageFont.load_default()

    def bbox(el):
        """统一取外框：line 用 p1/p2 算，其它用 box。"""
        if "box" in el:
            return el["box"]
        (x1, y1), (x2, y2) = el["p1"], el["p2"]
        return [min(x1, x2), min(y1, y2), abs(x2 - x1), abs(y2 - y1)]

    def back(b):
        """spec 坐标 -> 原图坐标"""
        return [b[0] - ox, b[1] - oy, b[2], b[3]]

    def dot(x, y, label, col, r=13):
        d.ellipse([x - r, y - r, x + r, y + r], fill=col)
        d.ellipse([x - r - 1, y - r - 1, x + r + 1, y + r + 1],
                  outline=(255, 255, 255), width=3)
        w = d.textlength(label, font=f)
        d.text((x - w / 2, y - 10), label, fill=(255, 255, 255), font=f)

    counts = {}
    blanks = []
    n = c = g = 0
    for el in spec["elements"]:
        b = back(bbox(el))
        k = el["kind"]
        if k in NODE_KINDS:
            n += 1
            lbl = f"N{n}"
            dot(b[0] + 9, b[1] + 9, lbl, COLOR["node"])
            counts["节点"] = n
            if not el.get("text"):
                blanks.append((lbl, el))
        elif k == "line":
            c += 1
            mx, my = b[0] + b[2] / 2, b[1] + b[3] / 2
            dot(mx, my - 16 if b[3] < 6 else my, f"C{c}", COLOR["line"])
            counts["连线"] = c
        elif k == "container":
            g += 1
            d.rectangle([b[0], b[1], b[0] + b[2], b[1] + b[3]],
                        outline=COLOR["container"], width=4)
            ty = b[1] - 20 if b[1] - 20 > 20 else b[1] + b[3] + 20
            dot(b[0] + 6, ty, f"G{g}", COLOR["container"])
            counts["分区"] = g
        elif k == "picture":
            d.rectangle([b[0], b[1], b[0] + b[2], b[1] + b[3]],
                        outline=COLOR["picture"], width=4)
        elif k == "brace":
            c += 1
            dot(b[0] + b[2] / 2, b[1] + b[3] / 2, f"C{c}", COLOR["line"])
        elif k == "arrow":
            c += 1
            dot(b[0] + b[2] / 2, b[1] + b[3] / 2, f"C{c}", COLOR["line"])

    pics = len([e for e in spec["elements"] if e["kind"] == "picture"])
    texts = len([e for e in spec["elements"] if e["kind"] == "text"])
    counts.update({"图片": pics, "文字": texts})

    im.save(args.out)

    print("=" * 62)
    print("这一页我读到的：")
    print("   ", "   ".join(f"{k} {v}" for k, v in counts.items()))
    print(f"    标号图: {args.out}")
    print("=" * 62)
    print("请用户**扫一眼这张图**，然后回答：")
    print("  ① 有没有漏的？  ② 有没有把两个东西数成一个的？")
    if blanks:
        print()
        print(f"⚠ 其中 {len(blanks)} 个框**原图里就没有文字**："
              f"{', '.join(l for l, _ in blanks)}")
        print("  → **不要自己填名字。** 先按「原件故意留空」处理")
        print("    （意思通常是：这里放你自己定义的内容）。")
        print("     除非用户确认那是「没画完」，才问要不要补。")
    unsup = spec.get("unsupported") or []
    if unsup:
        print()
        print(f"⚠ {len(unsup)} 个对象我画不出来，不会出现在重画结果里：")
        seen = {}
        for u in unsup:
            seen[u["what"]] = seen.get(u["what"], 0) + 1
        for w, k in seen.items():
            print(f"    - {w} ×{k}")
        print("  → 这个要**主动告诉用户**，别让他以为都画上了。")


if __name__ == "__main__":
    main()
