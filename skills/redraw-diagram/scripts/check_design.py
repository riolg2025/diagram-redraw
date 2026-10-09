#!/usr/bin/env python3
"""
check_design.py —— A 类：版面**品味**体检。

和 B 类检查的分工：

    B 类（check_a / spec_check）  版面**失控** —— 越界、冲突、放不下
                                  → **硬失败，必须修**
    A 类（这个脚本）              版面**品味** —— 不齐、不一致、太挤、颜色杂
                                  → **只报告，不拦**

★ 判据是**设计常识**，不是"和原件比"。
  （第 7 项 A 半的判据如果建成"和原件差多少"，就掉回"画法"里去了 —— 见 §十一）

★ 这些规则**同时也是生成规则**（A 的第 1 件事）：
  同一份东西，正向用来画、反向用来查。免得"检查说好、画出来不好"。

⚠️ 会误报的地方我都标了 `?`，并且**把数字打出来让你自己判** —— 不要盲信。

用法：
    python3 check_design.py work/spec.json
    python3 check_design.py work/spec.json --only 字号,线宽
"""

import argparse
import json
import os
import sys
from collections import Counter, defaultdict

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from common import wrap_text, line_polyline, pt_poly_dist  # noqa: E402


def load(p):
    with open(p, encoding="utf-8") as f:
        return json.load(f)


def _rgb(h):
    try:
        return tuple(int(str(h)[i:i + 2], 16) for i in (0, 2, 4))
    except Exception:
        return None


def _near(a, b, tol=5):
    """两个颜色是不是"其实同一个"。

    ⚠️ 阈值 5 是**有意选的**，不是随手写：
        差 ≤5  基本看不出差别，多半是手滑定成了两个值      → 算重复
        差 ≥6  开始能看出是**有意的深浅**（容器更浅、卡片更实）→ 不算重复
    一开始我写的是 12，结果把"有意的 7 级灰差"也报成了重复 —— 阈值太宽。
    """
    ra, rb = _rgb(a), _rgb(b)
    return bool(ra and rb and max(abs(x - y) for x, y in zip(ra, rb)) <= tol)


def _dist(a, b):
    ra, rb = _rgb(a), _rgb(b)
    return max(abs(x - y) for x, y in zip(ra, rb)) if ra and rb else 999


def _cls(items):
    """给对象分"类"：宽高接近 + 同一类画法。

    ⚠️ 这是**启发式**，会分错。所以要人工看一眼报告里的类。
    """
    out = defaultdict(list)
    for it in items:
        b = it.get("box")
        if not b:
            continue
        out[(round(b[2] / 4) * 4, round(b[3] / 4) * 4, it.get("kind"),
             it.get("fill"))].append(it)
    return {k: v for k, v in out.items() if len(v) >= 3}


# ---------------------------------------------------------------- 各项检查
def c_font(spec, R):
    """字号：同类一致吗？有层级吗？"""
    L = [it for it in spec["layout"] if (it.get("text") or {}).get("lines")]
    allsz = Counter((it["text"].get("font_pt")) for it in L)
    R("字号分布", "%s" % dict(sorted(allsz.items(), key=lambda kv: -(kv[0] or 0))),
      "%d 种字号 / %d 个带文字的对象" % (len(allsz), len(L)))
    for k, v in sorted(_cls(L).items(), key=lambda kv: -len(kv[1])):
        szs = Counter(it["text"].get("font_pt") for it in v)
        if len(szs) > 1:
            R("字号不一致", "类(%dx%d) 混用了 %s" % (k[0], k[1], dict(szs)),
              "同类元素字号应当一致", warn=True)
    # 层级：容器的标题不该比成员小
    for g in spec.get("groups") or []:
        mem = set(g.get("members") or [])
        mem_sz = [it["text"].get("font_pt") for it in L
                  if it.get("for") in mem and (it.get("text") or {}).get("lines")]
        for it in L:
            if it.get("for") != g["id"] or not (it.get("text") or {}).get("lines"):
                continue
            if mem_sz and it["text"].get("font_pt") and \
                    it["text"]["font_pt"] < max(mem_sz):
                R("层级倒挂", "%s 的标题 %.0fpt < 成员最大 %.0fpt"
                  % (g["id"], it["text"]["font_pt"], max(mem_sz)),
                  "容器的标题一般不该比它所辖元素小")


def c_line(spec, R):
    """线宽 / 箭头：一致吗？"""
    ls = [it for it in spec["layout"] if it.get("kind") == "line"]
    if not ls:
        return
    w = Counter((it.get("line") or {}).get("pt") for it in ls)
    R("线宽分布", "%s" % dict(w), "%d 条线" % len(ls))
    if len(w) > 1:
        main = w.most_common(1)[0][0]
        odd = [it["id"] for it in ls if (it.get("line") or {}).get("pt") != main]
        R("线宽不一致", "多数是 %.2fpt，这 %d 条不是：%s"
          % (main, len(odd), ", ".join(odd[:8])),
          "同类的线应当一样粗（粗细不同会让人以为是不同含义）", warn=True)
    # ★ 线条**颜色**也要查。我第一版漏了这条，结果自己撞出来：
    #   C10 那两条折线用的是 0056D1，其他 9 条是 0084CF —— 两种蓝、两种粗细，
    #   看起来就像"这是两类不同的关系"。
    lc = Counter((it.get("line") or {}).get("color") for it in ls)
    if len(lc) > 1:
        main = lc.most_common(1)[0][0]
        odd = [it["id"] for it in ls if (it.get("line") or {}).get("color") != main]
        R("线条颜色不一致", "多数是 #%s，这 %d 条不是：%s"
          % (main, len(odd), ", ".join(odd[:8])),
          "同类的线应当同色（颜色不同会让人以为是不同含义）", warn=True)
    a = Counter((it.get("arrow"), it.get("arrow_type")) for it in ls)
    if len(a) > 1:
        R("箭头不一致", "%s" % dict(a), "同类关系箭头应当一致", warn=True)
    else:
        R("箭头一致", "%s" % (list(a)[0],), "")


def c_padding(spec, R):
    """文字内边距：一致吗？够不够？"""
    L = [it for it in spec["layout"]
         if (it.get("text") or {}).get("lines") and (it["text"] or {}).get("insets")]
    if not L:
        return
    c = Counter(tuple(round(v, 1) for v in it["text"]["insets"]) for it in L)
    R("内边距分布", "%s" % {str(k): v for k, v in c.most_common()},
      "%d 个带文字的对象" % len(L))
    if len(c) > 1:
        main = c.most_common(1)[0][0]
        odd = [it["id"] for it in L
               if tuple(round(v, 1) for v in it["text"]["insets"]) != main]
        R("内边距不一致", "多数是 %s，这 %d 个不是：%s"
          % (str(main), len(odd), ", ".join(odd[:8])),
          "文字离框边的距离不一致，看上去会「忽紧忽松」", warn=True)
    small = [it["id"] for it in L
             if min(it["text"]["insets"][:2]) < 6.0]
    if small:
        R("内边距偏小", "这 %d 个文字离框边不到 6px：%s"
          % (len(small), ", ".join(small[:8])), "文字会显得贴边", warn=True)


def c_color(spec, R):
    """颜色：种类克制吗？有没有几乎一样的两种？"""
    L = spec["layout"]
    fills = [it.get("fill") for it in L if it.get("fill")]
    texts = [(it.get("text") or {}).get("color") for it in L
             if (it.get("text") or {}).get("lines")]
    R("填充色", "%d 种 %s" % (len(set(fills)), dict(Counter(fills).most_common(9))), "")
    R("文字色", "%d 种 %s" % (len(set(texts)), dict(Counter(texts).most_common(9))), "")
    for name, vals in (("填充", sorted(set(fills))), ("文字", sorted(set(texts)))):
        for i, a in enumerate(vals):
            for b in vals[i + 1:]:
                if _near(a, b):
                    R("颜色近似重复", "%s色 %s 和 %s 只差 %d/255"
                      % (name, a, b, _dist(a, b)),
                      "几乎是同一个颜色，留着两种会显得不干净", warn=True)
    if len(set(fills)) + len(set(texts)) > 12:
        R("颜色偏多", "填充 %d 种 + 文字 %d 种" % (len(set(fills)), len(set(texts))),
          "技术图一般控制在 6–8 种以内", warn=True)


def _pt_seg_dist(p, a, b):
    ax, ay = a; bx, by = b; px, py = p
    dx, dy = bx - ax, by - ay
    L = dx * dx + dy * dy
    if L == 0:
        return ((px - ax) ** 2 + (py - ay) ** 2) ** 0.5
    t = max(0.0, min(1.0, ((px - ax) * dx + (py - ay) * dy) / L))
    cx, cy = ax + t * dx, ay + t * dy
    return ((px - cx) ** 2 + (py - cy) ** 2) ** 0.5


def c_label_center(spec, R):
    """说明文字要落在**它标注的那条线**上（横向居中）。

    ★ 用户看标号图发现的：T6–T9 和 C6–C9 一一对应却没居中；T3 和 C2 也没居中。
      量出来 9 对里偏了 5 对（T6 −16.8、T8 −25.3、T5 +11.8、T2 −9.4、T3 −8.2）。

    ⚠️ 折线（bentConnector）**不能拿包围盒中点当中心** —— L 形的中心不在路径上。
      所以这里算的是"标签中心到**路径**的最近距离"。
    """
    L = spec["layout"]
    skipped = set()
    lines = defaultdict(list)
    for it in L:
        if it.get("kind") == "line":
            lines[it.get("for")].append(it)
    for it in L:
        t = it.get("text") or {}
        if not t.get("lines") or not it.get("box"):
            continue
        # ★ 用清单里写明的 `annotates`，**不去猜"最近的线"** ——
        #   猜就是"机械猜意义"，这个项目在关系端点上栽过一次（10 条错 5 条）。
        segs = lines.get(it.get("annotates"))
        if not segs:
            continue
        # ★ 只对**单段直线**严格查。
        #   折线/多段（比如 C10 由两段拼成、标签正好落在中间的缺口里）
        #   本来就不该"落在线上" —— 拿这条去套会得出"离线 42px"的假警报。
        #   ⚠️ 也就是说折线类的标注**现在没人查**，如实记下来。
        if len(segs) != 1 or (segs[0].get("prst") or "straightConnector1") \
                != "straightConnector1":
            skipped.add(it.get("annotates"))
            continue
        b = it["box"]
        c = (b[0] + b[2] / 2.0, b[1] + b[3] / 2.0)
        # ★ 用**实际折线路径**，不是 p1→p2 那条对角线。
        #   折线是 L 形，拿对角线算距离会得出"离线 103px"这种假警报。
        d = min(pt_poly_dist(c, line_polyline(s)) for s in segs)
        if d > 6.0:
            R("标注没居中", "%s（%s）的中心离线 %.1f px" % (it["id"], it.get("annotates"), d),
              "说明文字应当落在它标注的那条线上", warn=True)
    if skipped:
        R("（折线标注没查）", "这些关系是折线或多段，跳过了：%s"
          % "、".join(sorted(x for x in skipped if x)),
          "折线没有唯一的中轴，规则不好定 —— 这一档暂时没人查")


def c_text_padding(spec, R):
    """文字离框边够不够 —— 用**量出来的**墨迹高度模型。

    ★ 用户看标号图发现的：N1/N4 因为文字换行，上下很挤。
      实测：它们上下只留 7/5 px，而别的框是 15–30 px。
      （B 类的"放不下"查不出这个 —— 它确实**放得下**，只是难看。）

    墨迹高度 = (行数-1) × 行距 + 0.88 × 字号
      0.88 是在原件渲染图上量的：18pt 单行→30px(0.80)、12pt 单行→21px(0.85)、
      16pt 两行→70px(算出来 69.3) —— 取 0.88 偏保守一点。
    """
    DPI = 150.0
    for it in spec["layout"]:
        t = it.get("text") or {}
        b = it.get("box")
        if not t.get("lines") or not isinstance(b, list) or len(b) != 4:
            continue
        ins = t.get("insets") or [15.0, 7.5, 15.0, 7.5]
        pt = float(t.get("font_pt") or 10.0)
        aw = max(1.0, b[2] - float(ins[0]) - float(ins[2]))
        n = 0
        for one in t["lines"]:
            n += (len(wrap_text(str(one), aw, pt)) if t.get("wrap", True) else 1)
        ppm = DPI / 72.0
        lh = pt * (float(t.get("lnspc") or 120000) / 100000.0) * ppm
        ink = (n - 1) * lh + 0.88 * pt * ppm
        pad = (b[3] - ink) / 2.0
        if pad < 8.0:
            R("文字太挤", "%s 折成 %d 行后墨迹 %.0fpx，框高 %.0f，上下只剩 %.1f px"
              % (it["id"], n, ink, b[3], pad),
              "文字离上下边不到 8px 会显得挤（改框高或缩字号）", warn=True)


def c_inside_container(spec, R):
    """对象的框不该戳出它所属分组容器的框。"""
    by = {it["id"]: it for it in spec["layout"]}
    for g in spec.get("groups") or []:
        cont = [it for it in spec["layout"]
                if it.get("for") == g["id"] and it.get("kind") in ("rect", "roundrect")]
        if not cont:
            continue
        cb = cont[0]["box"]
        for it in spec["layout"]:
            if it.get("for") != g["id"] or not it.get("box") or it is cont[0]:
                continue
            b = it["box"]
            if (b[0] < cb[0] - 0.5 or b[1] < cb[1] - 0.5
                    or b[0] + b[2] > cb[0] + cb[2] + 0.5
                    or b[1] + b[3] > cb[1] + cb[3] + 0.5):
                R("戳出容器", "%s 的框超出了 %s 的框" % (it["id"], g["id"]),
                  "容器里的东西不该露出容器外", warn=True)


def c_margin(spec, R):
    """画布留白：内容别顶到边上。"""
    cw, ch = spec["source"]["canvas"]
    bs = [it["box"] for it in spec["layout"] if it.get("box")]
    bs += [[p[0], p[1], 0, 0] for it in spec["layout"]
           if it.get("kind") == "line" for p in (it.get("p1"), it.get("p2")) if p]
    if not bs:
        return
    x0 = min(b[0] for b in bs); y0 = min(b[1] for b in bs)
    x1 = max(b[0] + b[2] for b in bs); y1 = max(b[1] + b[3] for b in bs)
    m = {"左": x0, "上": y0, "右": cw - x1, "下": ch - y1}
    R("画布留白", " ".join("%s%.0f" % (k, v) for k, v in m.items()),
      "画布 %dx%d" % (cw, ch))
    tight = [k for k, v in m.items() if v < 6]
    if tight:
        R("留白不足", "内容顶到了画布的 %s 边（<6px）" % "、".join(tight),
          "看上去像没排完；画布应当比内容大一圈", warn=True)


def c_align(spec, R):
    """对齐：同类元素该在同一条线上。"""
    L = [it for it in spec["layout"] if it.get("box")]
    for k, v in sorted(_cls(L).items(), key=lambda kv: -len(kv[1])):
        ys = Counter(round(it["box"][1]) for it in v)
        if len(ys) <= 1 or max(ys.values()) < 2:
            continue
        main = max(ys, key=ys.get)
        # ★ 只有**离得近**的才算"同一行"。
        #   不这么干会误报：图上方一个同样尺寸的标签，和下面一排标签
        #   会被当成一类，于是"y 不一致" —— 可它们根本不在同一行。
        odd = [y for y, n in ys.items() if n == 1 and abs(y - main) <= 100]
        if odd:
            R("没对齐", "类(%dx%d) 里 y 不一致：%s（多数在 %s，最大偏 %.0fpx）"
              % (k[0], k[1], sorted(odd)[:6], main,
                 max(abs(y - main) for y in odd)),
              "同类元素一般应当对齐同一条线", warn=True)


def c_group_pad(spec, R):
    """分组容器到成员的留白。"""
    by = {it["id"]: it for it in spec["layout"]}
    for g in spec.get("groups") or []:
        cont = [it for it in spec["layout"]
                if it.get("for") == g["id"] and it.get("kind") in ("rect", "roundrect")]
        if not cont:
            continue
        cb = cont[0]["box"]
        ms = [by[it["id"]]["box"] for it in spec["layout"]
              if it.get("for") in set(g.get("members") or [])
              and by.get(it["id"], {}).get("box")
              and it.get("kind") not in ("line", "verbatim")]
        if not ms:
            continue
        x0 = min(m[0] for m in ms); y0 = min(m[1] for m in ms)
        x1 = max(m[0] + m[2] for m in ms); y1 = max(m[1] + m[3] for m in ms)
        pad = {"左": x0 - cb[0], "上": y0 - cb[1],
               "右": cb[0] + cb[2] - x1, "下": cb[1] + cb[3] - y1}
        R("分组留白", "%s %s" % (g["id"], " ".join("%s%.0f" % (k, v)
                                                  for k, v in pad.items())),
          "容器 %dx%d" % (cb[2], cb[3]))
        lr = abs(pad["左"] - pad["右"])
        if lr > 6:
            R("左右留白不匀", "%s 左 %.0f 右 %.0f（差 %.0f）"
              % (g["id"], pad["左"], pad["右"], lr),
              "容器左右留白不一致，看着会歪", warn=True)


def c_spacing(spec, R):
    """同组内间距：一致吗？（只在分组里比，跨组的比了会误报）"""
    by = {it["id"]: it for it in spec["layout"]}
    for g in spec.get("groups") or []:
        ms = [by[it["id"]] for it in spec["layout"]
              if it.get("for") in set(g.get("members") or [])
              and by.get(it["id"], {}).get("box")
              and it.get("kind") not in ("line", "verbatim")]
        # 只挑"排成一行"的（y 基本一致）
        if len(ms) < 3:
            continue
        ys = [m["box"][1] for m in ms]
        if max(ys) - min(ys) > 4:
            continue
        ms.sort(key=lambda m: m["box"][0])
        gaps = [round(ms[i + 1]["box"][0] - (ms[i]["box"][0] + ms[i]["box"][2]), 1)
                for i in range(len(ms) - 1)]
        R("组内间距", "%s %s" % (g["id"], gaps), "%d 个成员排成一行" % len(ms))
        if gaps and max(gaps) - min(gaps) > 4:
            R("间距不匀", "%s 的间距在 %.1f–%.1f 之间（差 %.1f）"
              % (g["id"], min(gaps), max(gaps), max(gaps) - min(gaps)),
              "同一行里的间距应当一致", warn=True)


CHECKS = [("字号", c_font), ("线", c_line), ("内边距", c_padding),
          ("标注居中", c_label_center), ("文字留白", c_text_padding),
          ("容器包含", c_inside_container),
          ("颜色", c_color), ("留白", c_margin), ("对齐", c_align),
          ("分组", c_group_pad), ("间距", c_spacing)]


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("spec")
    ap.add_argument("--only", help="只跑这几项，逗号分隔")
    args = ap.parse_args()
    spec = load(args.spec)

    only = set(args.only.split(",")) if args.only else None
    findings = []

    def R(title, detail, note="", warn=False):
        findings.append((title, detail, note, warn))

    print("=" * 70)
    print("A 类体检（版面品味）—— **只报告，不拦**")
    print("  判据是设计常识，不是「和原件比」。")
    print("=" * 70)
    for name, fn in CHECKS:
        if only and name not in only:
            continue
        before = len(findings)
        print()
        print("── %s " % name + "─" * (60 - len(name)))
        fn(spec, R)
        for t, d, n, w in findings[before:]:
            print("   %s %-14s %s" % ("⚠" if w else "·", t, d))
            if n:
                print("     %s%s" % (" " * 15, n))

    bad = [f for f in findings if f[3]]
    print()
    print("=" * 70)
    if bad:
        print("共 %d 项建议（A 类，不拦交付）：" % len(bad))
        for t, d, _n, _w in bad:
            print("   ⚠ %-14s %s" % (t, d))
    else:
        print("✓ 没看出明显的版面问题")
    print()
    print("⚠ A 类**没有硬判据**：上面每一条都请当作「提示」，不是「错误」。")
    print("  真要判断，得**看只属于我们自己的成品图**，不要和原件并排看")
    print("  （和原件并排，眼睛会自动去比「像不像」，而不是「好不好」）。")
    print("=" * 70)
    return 0


if __name__ == "__main__":
    sys.exit(main())
