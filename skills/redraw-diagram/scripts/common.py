#!/usr/bin/env python3
"""
common.py —— diagram-redraw 各脚本共用的东西。

这里放的每一条都对应一个**踩过的坑**，不是随手抽象出来的：

  * EMU_TO_PX / group_transform —— 组内元素给的是**组内坐标**，不减 chOff 会整体平移
  * render_size / px_of          —— 画布像素坐标与 EMU 的换算
  * mode_fill / darkest          —— 取色：填充用众数（压掉文字），文字/线条用最深
  * explicit_line_color          —— XML 里显式写了颜色就用它；采样会被穿过的元素污染
  * is_neutral                   —— 判断是不是"灰/黑"这类结构色
"""

import json
import os
from collections import Counter

# 幻灯片默认 13.333in x 7.5in，渲染 2000x1125 px => 150 dpi
DEFAULT_DPI = 150.0
EMU_PER_INCH = 914400


def emu_to_px(emu, dpi=DEFAULT_DPI):
    return emu / EMU_PER_INCH * dpi


class Xfrm:
    """一个 shape 的 a:xfrm：off / ext，组还带 chOff / chExt。"""

    def __init__(self, el, ns):
        self.ns = ns
        x = None
        for e in el.iter():
            if e.tag == ns + 'xfrm':
                x = e
                break
        if x is None:
            self.ok = False
            return
        o, ex = x.find(ns + 'off'), x.find(ns + 'ext')
        if o is None or ex is None:
            self.ok = False
            return
        self.ok = True
        self.off = (int(o.get('x')), int(o.get('y')))
        self.ext = (int(ex.get('cx')), int(ex.get('cy')))
        co, ce = x.find(ns + 'chOff'), x.find(ns + 'chExt')
        self.ch = ((int(co.get('x')), int(co.get('y'))),
                   (int(ce.get('cx')), int(ce.get('cy')))) if (
            co is not None and ce is not None) else None
        self.flipH = x.get('flipH') in ('1', 'true')
        self.flipV = x.get('flipV') in ('1', 'true')
        self.rot = int(x.get('rot')) if x.get('rot') else 0

    def scale_in(self):
        """组自身的 缩放：child 坐标系 -> 组坐标系"""
        if self.ch:
            return (self.ext[0] / self.ch[1][0], self.ext[1] / self.ch[1][1])
        return (1.0, 1.0)


class GroupCtx:
    """
    递归遍历组时携带的上下文。

    ★ 关键：abs = parent_off + (child_off - chOff) * scale
      漏掉 (child_off - chOff) 会让所有尺寸都对、位置整体偏移。
    """

    def __init__(self, off, scale, ch=(0, 0)):
        self.off = off
        self.scale = scale
        self.ch = ch

    def child(self, xf):
        if not xf.ch:
            return GroupCtx(self.off, self.scale, self.ch)
        return GroupCtx(
            (self.off[0] + (xf.off[0] - self.ch[0]) * self.scale[0],
             self.off[1] + (xf.off[1] - self.ch[1]) * self.scale[1]),
            xf.scale_in(), xf.ch[0])

    def box_px(self, xf, dpi=DEFAULT_DPI):
        x = self.off[0] + (xf.off[0] - self.ch[0]) * self.scale[0]
        y = self.off[1] + (xf.off[1] - self.ch[1]) * self.scale[1]
        w = xf.ext[0] * self.scale[0]
        h = xf.ext[1] * self.scale[1]
        return [emu_to_px(x, dpi), emu_to_px(y, dpi),
                emu_to_px(w, dpi), emu_to_px(h, dpi)]


# --------------------------------------------------------------------------
# 取色
# --------------------------------------------------------------------------
def mode_fill(img, box):
    """框内出现最多的颜色 = 填充色（文字是少数，会被众数压掉）。"""
    x0, y0 = int(box[0]), int(box[1])
    x1, y1 = int(box[0] + box[2]), int(box[1] + box[3])
    c = Counter()
    for yy in range(y0 + 4, y1 - 4, 2):
        for xx in range(x0 + 4, x1 - 4, 2):
            if 0 <= xx < img.width and 0 <= yy < img.height:
                c[img.getpixel((xx, yy))] += 1
    if not c:
        return None
    return "%02X%02X%02X" % c.most_common(1)[0][0]


def darkest(img, x0, y0, x1, y1):
    """区域里最深的像素 —— 取文字/线条颜色，比单点采样稳。"""
    best, bv = None, 10 ** 9
    for yy in range(int(y0), int(y1), 2):
        for xx in range(int(x0), int(x1), 2):
            if 0 <= xx < img.width and 0 <= yy < img.height:
                p = img.getpixel((xx, yy))
                if sum(p) < bv:
                    bv, best = sum(p), p
    return "%02X%02X%02X" % best if best else None


def lum(hexcolor):
    return sum(int(hexcolor[i:i + 2], 16) for i in (0, 2, 4))


# --------------------------------------------------------------------------
# spec 读写
# --------------------------------------------------------------------------
SPEC_VERSION = "0.1"


# --------------------------------------------------------------------------
# 文字宽度估算 / 折行
# --------------------------------------------------------------------------
# **为什么必须自己折行**：SVG 的 <text> **不会自动换行**，PPTX 会。
# SVG 侧不折行，长句就直接冲出框 —— 实测两句话叠在一行（用户发现的）。
#
# ⚠️ 这是**估算**，不是量出来的。宁可估宽一点（早换行），不要估窄（溢出）。
_EM = {True: 1.0, False: 0.52}        # 全角 / 半角，单位是 em


def _is_wide(ch):
    return ord(ch) > 0x2E7F            # 中日韩 + 全角标点


def text_width_px(s, font_pt, dpi=None):
    """估一行字有多宽（像素）。"""
    if not s:
        return 0.0
    em = sum(_EM[_is_wide(c)] for c in s)
    return em * float(font_pt or 10.0) * (dpi or DEFAULT_DPI) / 72.0


def wrap_text(s, max_px, font_pt, dpi=None):
    """按估算宽度折行，返回折好的每一行。

    中日韩可以逐字断；西文尽量在空格处断，断不了才硬断。
    """
    if not s:
        return [""]
    per_pt = (dpi or DEFAULT_DPI) / 72.0
    out, cur, w = [], "", 0.0
    for ch in s:
        cw = _EM[_is_wide(ch)] * float(font_pt or 10.0) * per_pt
        if cur and w + cw > max_px:
            if not _is_wide(ch) and " " in cur:
                cut = cur.rfind(" ")
                out.append(cur[:cut])
                cur = cur[cut + 1:] + ch
                w = text_width_px(cur, font_pt, dpi)
            else:
                out.append(cur)
                cur, w = ch, cw
        else:
            cur += ch
            w += cw
    out.append(cur)
    return out


def save_spec(spec, path):
    spec.setdefault("spec_version", SPEC_VERSION)
    with open(path, "w", encoding="utf-8") as f:
        json.dump(spec, f, ensure_ascii=False, indent=1)
    return path


def load_spec(path):
    with open(path, encoding="utf-8") as f:
        return json.load(f)


def ensure_dir(p):
    os.makedirs(p, exist_ok=True)
    return p
