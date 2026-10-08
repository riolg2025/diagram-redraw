#!/usr/bin/env python3
"""
mark_reading.py —— 把「我读到的」标在原图上，给用户确认。

★ 这是整个流程的第一句话，也是产品跟"截图"正面竞争的手段：
    截图给你的是原图；我给你的是"原图 + 我的理解"，**你可以一眼否掉它**。

★ 它标的是「我读到的**意义**」，不是「文件里有多少个形状」。
  所以输入是一份 reading.json（我的解读），不是 observations.json。

  「有 44 个形状」和「图上有什么」不是一回事 —— 项目栽过：
  11 个线条对象，10 条关系，**两个数都对**。

用法：
    python3 mark_reading.py work/reading.json --out work/marked.png
"""

import argparse
import json
import os
import sys

from PIL import Image, ImageDraw, ImageFont

BOLD = "/System/Library/Fonts/Supplemental/Arial Bold.ttf"
CN = "/System/Library/Fonts/STHeiti Medium.ttc"

COLOR = {
    "node": (206, 38, 38),        # 红 = 流程节点
    "line": (222, 122, 12),       # 橙 = 关系（箭头）
    "container": (21, 94, 190),   # 蓝 = 分区 / 容器
    "label": (120, 120, 120),     # 灰 = 说明文字（不是节点）
    "picture": (150, 40, 150),    # 紫 = 图片 / 截图
    "unsupported": (30, 30, 30),  # 黑 = 画不出来的
}
PREFIX = {"node": "N", "line": "C", "container": "G", "label": "T",
          "picture": "P", "unsupported": "X"}


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("reading")
    ap.add_argument("--out", required=True)
    ap.add_argument("--scale", type=float, default=1.0)
    ap.add_argument("--md", help="同时输出一份人读的 markdown")
    args = ap.parse_args()

    R = json.load(open(args.reading, encoding="utf-8"))
    im = Image.open(R["render"]).convert("RGB")
    if args.scale != 1.0:
        im = im.resize((int(im.width * args.scale), int(im.height * args.scale)))
    d = ImageDraw.Draw(im)
    sc = args.scale
    try:
        f = ImageFont.truetype(BOLD, int(17 * sc))
        fsm = ImageFont.truetype(CN, int(15 * sc))
    except Exception:
        f = fsm = ImageFont.load_default()

    def dot(x, y, label, col, r=None):
        r = r or int(13 * sc)
        d.ellipse([x - r, y - r, x + r, y + r], fill=col)
        d.ellipse([x - r - 1, y - r - 1, x + r + 1, y + r + 1],
                  outline=(255, 255, 255), width=max(2, int(3 * sc)))
        w = d.textlength(label, font=f)
        d.text((x - w / 2, y - 10 * sc), label, fill=(255, 255, 255), font=f)

    counts = {}
    for it in R["items"]:
        role = it["role"]
        col = COLOR.get(role, (100, 100, 100))
        b = it.get("box")
        if b:
            b = [v * sc for v in b]
        if role in ("node", "label", "unsupported") and b:
            dot(b[0] + 9 * sc, b[1] + 9 * sc, it["id"], col)
        elif role == "container" and b:
            d.rectangle([b[0], b[1], b[0] + b[2], b[1] + b[3]], outline=col,
                        width=max(2, int(4 * sc)))
            ty = b[1] - 20 * sc if b[1] - 20 * sc > 20 else b[1] + b[3] + 20 * sc
            dot(b[0] + 6 * sc, ty, it["id"], col)
        elif role == "picture" and b:
            d.rectangle([b[0], b[1], b[0] + b[2], b[1] + b[3]], outline=col,
                        width=max(2, int(4 * sc)))
            dot(b[0] + 9 * sc, b[1] + 9 * sc, it["id"], col)
        elif role == "line":
            if it.get("mark"):
                mx, my = it["mark"][0] * sc, it["mark"][1] * sc
            else:
                (x1, y1), (x2, y2) = it["p1"], it["p2"]
                mx, my = (x1 + x2) / 2 * sc, (y1 + y2) / 2 * sc
            dot(mx, my, it["id"], col)
        counts[role] = counts.get(role, 0) + 1

    im.save(args.out)

    if args.md:
        cn2 = {"node": "流程节点", "line": "关系", "container": "分区 / 容器",
               "label": "说明文字", "picture": "图片", "unsupported": "画不出来的"}
        order = ["container", "node", "label", "line", "picture", "unsupported"]
        L = ["# 我读到的", "",
             "> 原图：`%s`　标号图：`%s`" % (R["render"], os.path.basename(args.out)),
             "> **这是「我读到的」，请你核对**：有没有漏的、有没有把两个当成一个的。", ""]
        for role in order:
            grp = [x for x in R["items"] if x["role"] == role]
            if not grp:
                continue
            L += ["## %s（%d）" % (cn2[role], len(grp)), "",
                  "| 编号 | 说的是什么 |", "|---|---|"]
            for x in grp:
                L.append("| **%s** | %s |" % (x["id"], x.get("says") or ""))
            L.append("")
        if R.get("warnings"):
            L += ["## ⚠️ 需要你定的", ""]
            for w in R["warnings"]:
                L.append("- " + w)
            L.append("")
        open(args.md, "w", encoding="utf-8").write("\n".join(L))
        print("   人读版：%s" % args.md)

    # ---------------- 给用户看的话 ----------------
    cn = {"node": "流程节点", "line": "关系", "container": "分区/容器",
          "label": "说明文字", "picture": "图片", "unsupported": "画不出来的"}
    print("=" * 64)
    print("这一页我读到的（%s）：" % os.path.basename(R["render"]))
    print("   " + "   ".join("%s %d" % (cn.get(k, k), v)
                            for k, v in counts.items()))
    print("   标号图：%s" % args.out)
    print("=" * 64)
    print("**请扫一眼这张图**，然后回答：")
    print("  ① 有没有漏掉的？   ② 有没有把两个东西当成一个的？")
    print("  ③ 有没有我标成节点、其实是别的东西的？")
    for w in R.get("warnings") or []:
        print()
        print("⚠ " + w)
    print()
    print("（颜色：红=流程节点 橙=关系 蓝=分区 灰=说明文字 紫=图片 黑=画不出来）")


if __name__ == "__main__":
    main()
