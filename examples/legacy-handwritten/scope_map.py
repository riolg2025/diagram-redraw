#!/usr/bin/env python3
"""
dryrun-02 第一步：范围确认图。

按新流程，第一句话不是"我看到了什么"，而是
"你要我理解的是哪一块、要重画的是哪一块"。

第 5 页正好是三块东西塞在一页里，这个问题很实。
"""
from PIL import Image, ImageDraw, ImageFont

PNG = "diagram-redraw/dryrun-02/orig/page-0003.png"
OUT = "diagram-redraw/dryrun-02/scope-map.png"

REGIONS = [
    ("A", 63, 209, 625, 640, (206, 38, 38),
     "A · 左侧：EasyOps / Agent 部署示意"),
    ("B", 747, 209, 1224, 616, (21, 94, 190),
     "B · 右侧：四层架构（流程编排 → 执行）"),
    ("C", 60, 895, 1845, 210, (120, 120, 120),
     "C · 底部两段文字说明（这不是图）"),
]

im = Image.open(PNG).convert("RGB")
# 稍微调暗，让框更醒目
im = Image.blend(im, Image.new("RGB", im.size, (255, 255, 255)), 0.12)
d = ImageDraw.Draw(im)
f = ImageFont.truetype("/System/Library/Fonts/STHeiti Medium.ttc", 24)
fb = ImageFont.truetype("/System/Library/Fonts/Supplemental/Arial Bold.ttf", 22)

for tag, x, y, w, h, col, label in REGIONS:
    d.rectangle([x, y, x + w, y + h], outline=col, width=5)
    # 标签贴在框外侧左上
    ty = y - 34 if y - 34 > 10 else y + 8
    tw = d.textlength(label, font=f)
    d.rectangle([x - 2, ty - 4, x + tw + 16, ty + 32], fill=col)
    d.text((x + 6, ty), label, fill=(255, 255, 255), font=f)

im.save(OUT)
print("写出:", OUT)
