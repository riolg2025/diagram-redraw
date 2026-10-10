#!/usr/bin/env python3
"""
reconcile.py —— 「重排」路线的对账工具。

★ 它**不跟原件比像素** ✗ —— 重排路线的版面是**故意**不一样的，
  整页像素差会把「故意不一样」和「画错了」混在一起报警 ✗（`verify.py` 就是那样）。

  它比的是：**清单里说该表达的，输出里表达了没有。**

三查：
  ① 内容   清单里每个实体的文字，输出里有吗
  ② 关系   每条关系，几何上表达了没有（**线**和**邻近**两种都算）
  ③ 分区   输出的分区有没有重叠（SVG 的 data-pptx-bounds）

★ 要能对账，输出里的 id 必须**对得上清单的 id**（实体 id / 关系 id）。
  这是硬要求：对不上就直接报错，不猜。

用法：
    python3 reconcile.py <输出.svg> --spec <清单.json>
"""

import argparse
import json
import math
import os
import re
import sys

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, HERE)
from common import line_polyline  # noqa: E402

NEAR = 10.0          # 端点离实体这么近就算"贴上了"（px）
FAR = 0.75           # 邻近表达：A 到 B 的距离，要小于 A 到最近**无关**实体的距离的这个比例


# --------------------------------------------------------------------------
# 抽取：把「清单」和「输出 SVG」都归一成同一种形状
# --------------------------------------------------------------------------
def _box_of(o):
    b = o.get("box")
    return [float(v) for v in b] if b else None


def from_spec(spec):
    """清单（带 layout）→ {boxes:{id:[box]}, texts:{id:text}, lines:{rel:[pts]}}"""
    boxes, texts, lines = {}, {}, {}
    for e in spec.get("entities") or []:
        t = e.get("text_user") or e.get("text") or ""
        if t:
            texts[e["id"]] = t
        if e.get("annotates"):
            texts[e["id"]] = t
    for it in spec.get("layout") or []:
        i, kind = it.get("id"), it.get("kind")
        if kind == "line":
            pts = line_polyline(it)
            lines.setdefault(it.get("for") or i, []).extend(pts)
        elif _box_of(it):
            key = it.get("for") or i
            boxes.setdefault(key, []).append(_box_of(it))
            tx = it.get("text") or {}
            s = "".join(str(x) for x in (tx.get("lines") or []))
            if s:
                texts.setdefault(key, s)
    return boxes, texts, lines


def from_svg(path):
    """输出 SVG → 同样的形状。**靠 id 对账**：id 必须就是清单里的 id。"""
    src = open(path, encoding="utf-8").read()
    boxes, texts, lines, groups = {}, {}, {}, {}

    for m in re.finditer(r'<rect\b([^>]*)/?>', src):
        a = dict(re.findall(r'([\w:-]+)="([^"]*)"', m.group(1)))
        i = a.get("id")
        if not i or "x" not in a:
            continue
        boxes.setdefault(i, []).append(
            [float(a["x"]), float(a["y"]), float(a["width"]), float(a["height"])])

    # 文字：挂到同一个 <g> 里最近的带 id 的 rect 上；找不到就按 text 自己的 id
    for m in re.finditer(r'<text\b([^>]*)>(.*?)</text>', src, re.S):
        a = dict(re.findall(r'([\w:-]+)="([^"]*)"', m.group(1)))
        t = re.sub(r"<[^>]+>", "", m.group(2)).strip()
        if not t or a.get("id"):
            continue
        x, y = float(a.get("x", 0)), float(a.get("y", 0))
        best, bd = None, 1e9
        for i, bs in boxes.items():
            for b in bs:
                if b[0] - 40 <= x <= b[0] + b[2] + 40 and b[1] - 40 <= y <= b[1] + b[3] + 40:
                    d = abs((b[1] + b[3] / 2) - y)
                    if d < bd:
                        best, bd = i, d
        if best:
            texts.setdefault(best, t)

    for m in re.finditer(r'<line\b([^>]*)/?>', src):
        a = dict(re.findall(r'([\w:-]+)="([^"]*)"', m.group(1)))
        i = a.get("id")
        if not i:
            continue
        lines.setdefault(i, []).extend(
            [(float(a["x1"]), float(a["y1"])), (float(a["x2"]), float(a["y2"]))])

    for m in re.finditer(r'<g\b([^>]*)>', src):
        a = dict(re.findall(r'([\w:-]+)="([^"]*)"', m.group(1)))
        if a.get("id") and a.get("data-pptx-bounds"):
            b = [float(v) for v in a["data-pptx-bounds"].split()]
            groups[a["id"]] = b
            boxes.setdefault(a["id"], []).append(b)   # ★ 分组自身的框也是它的锚
    return boxes, texts, lines, groups


# --------------------------------------------------------------------------
# 几何
# --------------------------------------------------------------------------
def _in(b, p, m=NEAR):
    return b[0] - m <= p[0] <= b[0] + b[2] + m and b[1] - m <= p[1] <= b[1] + b[3] + m


def _dist_box(b, p):
    dx = max(b[0] - p[0], 0, p[0] - (b[0] + b[2]))
    dy = max(b[1] - p[1], 0, p[1] - (b[1] + b[3]))
    return math.hypot(dx, dy)


def _gap(a, b):
    """两个框之间的最短距离（相交为 0）。"""
    dx = max(a[0] - (b[0] + b[2]), b[0] - (a[0] + a[2]), 0)
    dy = max(a[1] - (b[1] + b[3]), b[1] - (a[1] + a[3]), 0)
    return math.hypot(dx, dy)


def _inside(b, outer):
    """b 是不是**装进**了 outer 里（bbox 包含）。"""
    return (b[0] >= outer[0] - 1 and b[1] >= outer[1] - 1 and
            b[0] + b[2] <= outer[0] + outer[2] + 1 and
            b[1] + b[3] <= outer[1] + outer[3] + 1)


def _union(bs):
    x0 = min(b[0] for b in bs); y0 = min(b[1] for b in bs)
    x1 = max(b[0] + b[2] for b in bs); y1 = max(b[1] + b[3] for b in bs)
    return [x0, y0, x1 - x0, y1 - y0]


# --------------------------------------------------------------------------
def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("output", help="输出 SVG")
    ap.add_argument("--spec", required=True, help="清单（①那一步的产物）")
    ap.add_argument("--quiet", action="store_true")
    a = ap.parse_args()

    S = json.load(open(a.spec, encoding="utf-8"))
    E = {e["id"]: e for e in S.get("entities") or []}
    G = {g["id"]: g for g in S.get("groups") or []}
    REL = {r["id"]: r for r in S.get("relations") or []}

    o_boxes, o_texts, o_lines, o_groups = from_svg(a.output)
    bad = 0
    print("重排对账：%s  ←→  %s" % (os.path.basename(a.output), os.path.basename(a.spec)))
    print()

    # ── ① 内容 ────────────────────────────────────────────────────────
    print("① 内容：清单里每个实体的文字，输出里有吗")
    blob = "".join(o_texts.values()) + "".join(
        re.sub(r"<[^>]+>", "", t) for t in re.findall(r"<text[^>]*>(.*?)</text>",
                                                      open(a.output, encoding="utf-8").read(), re.S))
    def _sq(x):        # ★ 去掉所有空白再比：实体文字可能是两行（"虚拟机\nAgent"）
        return re.sub(r"\s+", "", str(x or ""))
    bsq = _sq(blob)
    miss = [i for i, e in E.items()
            if _sq(e.get("text_user") or e.get("text")) and
            _sq(e.get("text_user") or e.get("text")) not in bsq]
    print("   实体 %d 个，缺 %d 个%s" % (len(E), len(miss), "：" + "、".join(miss) if miss else " ✓"))
    gmiss = []
    for gid, g in G.items():
        ti = (g.get("title") or {}).get("text")
        if ti and not _sq(ti) in bsq:
            gmiss.append("%s 的标题「%s」" % (gid, ti))
    if gmiss:
        print("   分组标题缺 %d 个：%s" % (len(gmiss), "、".join(gmiss)))
    bad += len(miss) + len(gmiss)
    print()

    # ── ② 关系 ────────────────────────────────────────────────────────
    print("② 关系：每条关系，几何上表达了没有（线 / 邻近 都算）")
    anchors = {}
    for i in list(E) + list(G):
        bs = o_boxes.get(i) or []
        if not bs and i in G:
            bs = o_boxes.get(i) or o_boxes.get("ct-" + i) or []
        if bs:
            anchors[i] = _union(bs)
    n_line = n_bad = n_contain = 0
    review = []
    for rid, r in REL.items():
        A, B = r.get("from"), r.get("to")
        an, bn = anchors.get(A), anchors.get(B)
        if not an or not bn:
            print("   ✗ %-5s 端点画不出来（%s 或 %s 在输出里找不到）" % (rid, A, B))
            n_bad += 1
            continue
        pts = o_lines.get(rid)
        if pts:
            # ★ 这一条是**机械可查**的：有线，两端就得都贴上
            hitA = any(_in(an, p) for p in pts)
            hitB = any(_in(bn, p) for p in pts)
            if hitA and hitB:
                n_line += 1
            else:
                print("   ✗ %-5s 有线，但两端没都贴上（贴 %s=%s 贴 %s=%s）"
                      % (rid, A, B, hitA, hitB))
                n_bad += 1
            continue
        # ★ 没有线 → **先看是不是"装在框里"**。
        #   「A 在 B 里面」是**机械可查的事实** ✓ 不是判断 ✓ ——
        #   所以它该由工具直接判，不该丢给人 ✗
        #   （`membership` 的标准表达就是包含：「格 ∈ 表」「表 ∈ 带」都不需要画线）
        if _inside(an, bn) or _inside(bn, an):
            n_contain += 1
            continue

        # ★ 既没有线、又没装在框里 → **工具不判"邻近算不算说得清"** ✗
        #   那是个判断，得人（或另一双眼睛）看。这里只把**证据**摆出来。
        gAB = _gap(an, bn)
        mid_a = (an[0] + an[2] / 2, an[1] + an[3] / 2)
        mid_b = (bn[0] + bn[2] / 2, bn[1] + bn[3] / 2)
        ali = []
        if abs(mid_a[0] - mid_b[0]) <= 4:
            ali.append("水平中心对齐")
        if abs(mid_a[1] - mid_b[1]) <= 4:
            ali.append("垂直中心对齐")
        review.append((rid, A, B, gAB, "、".join(ali) or "**没有对齐**"))
    print("   有线且两端都贴上        %d 条" % n_line)
    if n_contain:
        print("   装在一个框里（包含）     %d 条  ← **机械可查** ✓" % n_contain)
    if review:
        print("   没有线（**工具不判，需要人/另一双眼睛看**）%d 条：" % len(review))
        for rid, A, B, g, ali in review:
            print("      %-5s %s → %s    相距 %-5.0f  %s" % (rid, A, B, g, ali))
        print("      ↳ 判据：`link`/`order` 都允许用位置，**但前提是「一眼看得出来」**")
        print("        请对照 design-step2.md 第三节里写的理由，看这几条成不成立。")
    if n_bad or review:
        pass
    bad += n_bad
    print()

    # ── ③ 分区 ────────────────────────────────────────────────────────
    print("③ 分区：输出的分区有没有重叠")
    if not o_groups:
        print("   （输出里没有 data-pptx-bounds，跳过）")
    else:
        ks = sorted(o_groups)
        ov = 0
        for i in range(len(ks)):
            for j in range(i + 1, len(ks)):
                A, B = o_groups[ks[i]], o_groups[ks[j]]
                # ★ 容器套容器是**正常**的（表格本来就装在带里）✗
                #   只有"**部分**重叠、互不包含"才算失控 ✗
                #   （主线 spec_check 早就学过这一条，这里照它）
                if _inside(A, B) or _inside(B, A):
                    continue
                ix = max(0, min(A[0] + A[2], B[0] + B[2]) - max(A[0], B[0]))
                iy = max(0, min(A[1] + A[3], B[1] + B[3]) - max(A[1], B[1]))
                if ix * iy > 1.0:
                    print("   ✗ %s 和 %s **部分**重叠 %.0f×%.0f px" % (ks[i], ks[j], ix, iy))
                    ov += 1
        print("   %d 个分区，重叠 %d 处%s" % (len(ks), ov, " ✓" if not ov else ""))
        bad += ov
    print()

    print("=" * 60)
    if bad:
        print("✗ %d 处没对上" % bad)
    else:
        print("✓ 清单里说该表达的，输出里都表达了")
        print("  ⚠️ 这只说明「该表达的表达了」——**不说明「画得好不好看」** ✗")
    return 0 if bad == 0 else 1


if __name__ == "__main__":
    sys.exit(main())
