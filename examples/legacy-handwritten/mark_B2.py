#!/usr/bin/env python3
"""
dryrun-02：B 区「要问你的东西」标号图。

坐标从 inventory-B.json 按**形状名**查，不再手抄——手抄过一次，错得很隐蔽。
（上一版错在组坐标变换漏减 chOff：所有尺寸都对、位置整体偏 55px 右 79px 下。）
"""
import json
from PIL import Image, ImageDraw, ImageFont

PNG = "diagram-redraw/dryrun-02/orig/page-0003.png"
OUT = "diagram-redraw/dryrun-02/B-questions.png"
INV = "diagram-redraw/dryrun-02/inventory-B.json"

inv = json.load(open(INV))
by = {r["name"]: r for r in inv}

GROUPS = [
    ("D", (150, 60, 190), [
        ("直接箭头连接符 163", "虚线 2.25pt sysDash  蓝框→子流程"),
        ("直接箭头连接符 165", "虚线 1.5pt dash     人工审核→蓝框"),
        ("直接箭头连接符 167", "虚线 1.5pt dash     自动化工具→蓝框"),
    ]),
    ("T", (25, 25, 25), [("箭头: 下 2059", "粗黑箭头 人工审核→Powershell服务")]),
    ("Y", (232, 150, 0), [
        ("箭头: 上下 2056", "黄色细箭头 工具层→执行层"),
        ("箭头: 右 2060", "黄色大箭头（出画）"),
    ]),
    ("W", (0, 138, 122), [
        ("箭头: 上下 164", "双向箭头 dbscript↔sql client"),
        ("箭头: 上下 165", "双向箭头 其他工具↔对应接口"),
    ]),
    ("K", (21, 94, 190), [
        ("左大括号 178", "分组括号 rot=270°"),
        ("左大括号 179", "分组括号 rot=90°"),
    ]),
]

im0 = Image.open(PNG).convert("RGB")
mask = Image.new("L", im0.size, 70)
ImageDraw.Draw(mask).rectangle([700, 200, 1990, 840], fill=255)
im = Image.composite(im0, Image.blend(im0, Image.new("RGB", im0.size, (255, 255, 255)), 0.55), mask)
d = ImageDraw.Draw(im)
f = ImageFont.truetype("/System/Library/Fonts/Supplemental/Arial Bold.ttf", 18)


def center(name):
    r = by.get(name)
    if not r:
        return None
    x0, y0, x1, y1 = r["box_px"]
    return (x0 + x1) / 2, (y0 + y1) / 2


print("采样（用未调暗的原图，坐标已修正）：")
marks = []
for tag, col, items in GROUPS:
    for i, (name, desc) in enumerate(items, 1):
        c = center(name)
        if c is None:
            print(f"  !! 找不到 {name}")
            continue
        p = im0.getpixel((int(c[0]), int(c[1])))
        print(f"  {tag}{i} {desc:38s} ({c[0]:5.0f},{c[1]:5.0f}) #{p[0]:02X}{p[1]:02X}{p[2]:02X}")
        marks.append((c[0], c[1], f"{tag}{i}", col))

for x, y, lab, c in marks:
    r = 15
    d.ellipse([x - r, y - r, x + r, y + r], fill=c)
    d.ellipse([x - r - 1, y - r - 1, x + r + 1, y + r + 1], outline=(255, 255, 255), width=3)
    w = d.textlength(lab, font=f)
    d.text((x - w / 2, y - 11), lab, fill=(255, 255, 255), font=f)

pics = sorted([r for r in inv if r["type"] == "PICTURE"],
              key=lambda r: (r["box_px"][1], r["box_px"][0]))
for i, r in enumerate(pics, 1):
    b = r["box_px"]
    d.rectangle(b, outline=(200, 40, 40), width=4)
    d.ellipse([b[0] - 15, b[1] - 15, b[0] + 15, b[1] + 15], fill=(200, 40, 40),
              outline=(255, 255, 255), width=3)
    w = d.textlength(f"P{i}", font=f)
    d.text((b[0] - w / 2, b[1] - 11), f"P{i}", fill=(255, 255, 255), font=f)

im.save(OUT)
print(f"\n写出 {OUT}；标记 {len(marks)} 个，图片 {len(pics)} 张")
