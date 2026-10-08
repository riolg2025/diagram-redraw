#!/usr/bin/env python3
"""
verify.py —— 把重画结果渲染出来，和原图做逐像素差分。

**这是唯一能发现"看起来对、其实错"的办法。**
踩过的坑全靠它抓出来：文字整体偏移、描边粗一倍、菱形画成方框、
以及最阴的一次——脚本报错没跑，我却在"验证"上一轮的旧文件。

用法：
  python3 verify.py orig.png redraw.png --outdir diff/ [--crop x,y,w,h]
"""

import argparse
import os
import sys

import numpy as np
from PIL import Image


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("orig")
    ap.add_argument("redraw")
    ap.add_argument("--outdir", required=True)
    ap.add_argument("--crop", help="只比这一块：x,y,w,h")
    ap.add_argument("--threshold", type=int, default=90)
    args = ap.parse_args()

    os.makedirs(args.outdir, exist_ok=True)
    A = Image.open(args.orig).convert("RGB")
    B = Image.open(args.redraw).convert("RGB")
    if A.size != B.size:
        sys.exit(f"尺寸不同：{A.size} vs {B.size}（先确认两边同一渲染尺度）")
    box = tuple(int(v) for v in args.crop.split(",")) if args.crop else None
    if box:
        x, y, w, h = box
        A = A.crop((x, y, x + w, y + h))
        B = B.crop((x, y, x + w, y + h))

    a = np.asarray(A, dtype=int)
    b = np.asarray(B, dtype=int)
    d = np.abs(a - b).sum(axis=2)
    sig = (d > args.threshold)

    print(f"尺寸 {A.size[0]}x{A.size[1]}")
    print(f"平均差 {d.mean():.1f}   显著差异像素 {sig.mean()*100:.1f}%")
    Image.fromarray(np.clip(d, 0, 255).astype("uint8")).save(
        os.path.join(args.outdir, "diff.png"))

    rows = sig.sum(axis=1)
    bands, run = [], None
    for i, v in enumerate(rows):
        if v > 60 and run is None:
            run = i
        elif v <= 60 and run is not None:
            bands.append((run, i - 1, int(rows[run:i].max())))
            run = None
    if bands:
        print("差异集中的横带（峰值>60）：")
        for a_, b_, c in bands[:20]:
            print(f"   y {a_}..{b_}  峰值 {c}")
        print("  （整行且峰值≈宽度 的，通常只是边框 1px 取整偏移，可忽略）")
    print(f"差分图 -> {os.path.join(args.outdir, 'diff.png')}")

    if sig.mean() > 0.15:
        print("⚠ 显著差异超过 15%，先别交付——大概率有结构性的错。")


if __name__ == "__main__":
    main()
