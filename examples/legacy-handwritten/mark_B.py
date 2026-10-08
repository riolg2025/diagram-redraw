#!/usr/bin/env python3
"""
dryrun-02：给 B 区出「标号图」，专门标出**我要问你的东西**。

不是把 63 个对象全编号——那没法看。只标：
  虚线箭头 / 粗黑箭头 / 黄色箭头 / 双向箭头 / 分组括号 / 图片 / 空框
"""
import json
from PIL import Image, ImageDraw, ImageFont
from pptx import Presentation

A = '{http://schemas.openxmlformats.org/drawingml/2006/main}'
PNG = "diagram-redraw/dryrun-02/orig/page-0003.png"
OUT = "diagram-redraw/dryrun-02/B-marked.png"

# B 区的对象（已做组内坐标换算），位置来自 inventory-B.json
inv = json.load(open("diagram-redraw/dryrun-02/inventory-B.json"))
by = {}
for r in inv:
    by.setdefault(r["name"], r)


def box(name):
    r = by.get(name)
    return r["box_px"] if r else None


im = Image.open(PNG).convert("RGB")
# 只把 B 区调亮，其余压暗
d = ImageDraw.Draw(im)
mask = Image.new("L", im.size, 90)
ImageDraw.Draw(mask).rectangle([747, 209, 1971, 825], fill=255)
im = Image.composite(im, Image.blend(im, Image.new("RGB", im.size, (255, 255, 255)), 0.55), mask)
d = ImageDraw.Draw(im)

f = ImageFont.truetype("/System/Library/Fonts/Supplemental/Arial Bold.ttf", 17)
fc = ImageFont.truetype("/System/Library/Fonts/STHeiti Medium.ttc", 19)

COL = {"虚": (150, 60, 190), "粗": (32, 32, 32), "黄": (230, 150, 0),
       "双": (0, 140, 120), "括": (21, 94, 190), "图": (200, 40, 40),
       "空": (230, 110, 20), "层": (90, 90, 90)}


def marker(x, y, label, col, r=13):
    d.ellipse([x - r, y - r, x + r, y + r], fill=col)
    d.ellipse([x - r - 1, y - r - 1, x + r + 1, y + r + 1], outline=(255, 255, 255), width=2)
    w = d.textlength(label, font=f)
    d.text((x - w / 2, y - 10), label, fill=(255, 255, 255), font=f)


def ring(r, col, label, lw=4, outer=False):
    x0, y0, x1, y1 = r
    d.rectangle([x0 - 3, y0 - 3, x1 + 3, y1 + 3], outline=col, width=lw)
    marker(x0 - 3, y0 - 3 if not outer else y0 + 16, label, col)


# ---- 空框：6 个 roundRect + 2 个 diamond ----
blanks = ["矩形: 圆角 7", "矩形: 圆角 55", "矩形: 圆角 57", "矩形: 圆角 59",
          "矩形: 圆角 65", "矩形: 圆角 84", "菱形 145", "菱形 155"]
for i, n in enumerate(blanks, 1):
    b = box(n)
    if b:
        ring(b, COL["空"], f"E{i}", lw=3)

# ---- 图片 ----
pics = [r for r in inv if r["type"] == "PICTURE"]
for i, r in enumerate(sorted(pics, key=lambda r: (r["box_px"][1], r["box_px"][0])), 1):
    b = r["box_px"]
    d.rectangle(b, outline=COL["图"], width=4)
    marker(b[0] + 4, b[1] + 4, f"P{i}", COL["图"])

# ---- 四层 ----
for i, (lab, y) in enumerate([("GUI 编排层", 292), ("流程引擎层", 470),
                              ("工具服务层", 700), ("工具执行层", 820)], 1):
    d.text((762, y), f"L{i} {lab}", fill=(255, 255, 255), font=fc,
           stroke_width=4, stroke_fill=COL["层"])

im.save(OUT)
print("写出:", OUT)
print("空框", len(blanks), "个；图片", len(pics), "张")
for i, r in enumerate(sorted(pics, key=lambda r: (r["box_px"][1], r["box_px"][0])), 1):
    print(f"  P{i} {r['name'][:14]:16s} {r['box_px']}")
