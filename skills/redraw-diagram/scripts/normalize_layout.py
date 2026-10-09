#!/usr/bin/env python3
"""
normalize_layout.py —— 按设计规则，把清单的**排版**自动拉齐一遍。

和 check_design.py 的关系：

    check_design      反向查：哪里不齐、不一致
    normalize_layout  正向画：把它拉齐
    —— **同一套规则**。免得检查说"不齐"、画出来还是那样。

★ 只动画法：坐标、尺寸、线宽线色、内边距、填充灰。
  **不碰内容、关系、分组**（那是意义）。

⚠️ 三条"要判断"的事，我**不自动做**，只报出来：
    · 颜色层次（谁该更浅）—— 那是设计决定，不是规则
    · 字号层级（谁该更大）—— 同上
    · 折线类标注的居中 —— 折线没有唯一中轴，规则不好定

用法：
    python3 normalize_layout.py work/spec.json --out work/spec-norm.json
    python3 normalize_layout.py work/spec.json --out ... --pad 24 --dry-run
"""

import argparse
import json
import os
import sys
from collections import Counter, defaultdict

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from common import (line_polyline, pt_poly_dist, wrap_text)  # noqa: E402

MIN_PAD = 8.0          # 文字离框边最少留这么多（见 check_design 的同名规则）
MAJORITY = 0.7         # 要"统一"某一项，多数派至少得占这么多 —— 免得把有意的差异也抹平
NEAR_COLOR = 5         # 差 ≤5/255 算"其实是同一个颜色"


def load(p):
    with open(p, encoding="utf-8") as f:
        return json.load(f)


def _rgb(h):
    try:
        return tuple(int(str(h)[i:i + 2], 16) for i in (0, 2, 4))
    except Exception:
        return None


def _near(a, b, tol=NEAR_COLOR):
    ra, rb = _rgb(a), _rgb(b)
    return bool(ra and rb and max(abs(x - y) for x, y in zip(ra, rb)) <= tol)


def _majority(counter, min_share=None):
    """多数派；占比不够就返回 None（表示"不该硬统一"）。

    ★ 门槛**按项不同**：
        线条颜色/线宽 —— 0.7。因为**真的可能有第二种颜色**（比如标红的异常路径），
                          没有压倒性多数就不要抹平。
        内边距       —— 0.5。**没有哪种排版是"故意用两种内边距"的**，
                          两种就是手滑，多数派够了就该统一。
      （实测：内边距 21 vs 11 = 65.6%，用 0.7 会被挡掉、统一不了。）
    """
    if not counter:
        return None
    top, n = counter.most_common(1)[0]
    total = sum(counter.values())
    return top if total and n / total >= (min_share or MAJORITY) else None


# --------------------------------------------------------------------------
# 每一条规则：吃 spec，改 spec，返回改了什么（人话）
# 顺序有讲究 —— 后面的依赖前面的结果。
# --------------------------------------------------------------------------
def r_line_style(spec):
    """"同类的线应当一样粗、一样色"。取多数派，套给少数派。"""
    ls = [it for it in spec["layout"] if it.get("kind") == "line"]
    if len(ls) < 3:
        return []
    done = []
    for key, name in (("pt", "线宽"), ("color", "线条颜色")):
        c = Counter((it.get("line") or {}).get(key) for it in ls)
        c.pop(None, None)
        if len(c) < 2:
            continue
        top = _majority(c)
        if top is None:
            done.append("⚠ %s 有 %d 种、没有明显多数派，**没动**：%s"
                        % (name, len(c), dict(c)))
            continue
        odd = [it for it in ls if (it.get("line") or {}).get(key) != top]
        for it in odd:
            it["line"][key] = top
        done.append("%s：%d 条改成 %s（原来 %s）"
                    % (name, len(odd), top,
                       sorted({str((x.get("line") or {}).get(key)) for x in odd})))
    return done


def r_unify_insets(spec):
    """文字内边距统一。取多数派，而不是我拍一个值。"""
    L = [it for it in spec["layout"]
         if (it.get("text") or {}).get("lines") and (it["text"] or {}).get("insets")]
    if len(L) < 3:
        return []
    c = Counter(tuple(round(v, 1) for v in it["text"]["insets"]) for it in L)
    if len(c) < 2:
        return []
    top = _majority(c, min_share=0.5)
    if top is None:
        return ["⚠ 内边距有 %d 种、连一半都没有，**没动**：%s" % (len(c), dict(c))]
    odd = [it for it in L
           if tuple(round(v, 1) for v in it["text"]["insets"]) != top]
    for it in odd:
        it["text"]["insets"] = list(top)
    return ["内边距：%d 个统一成 %s（原来 %s）"
            % (len(odd), top, sorted({str(tuple(round(v, 1) for v in x["text"]["insets"]))
                                      for x in L if x not in odd}))]


def r_merge_near_colors(spec):
    """两种只差几个色阶的填充，合成一种（并到多数派）。

    ⚠️ 这只是"去掉手滑"。真想要**层次**（容器更浅、卡片更实）是**设计决定**，
       这里不做 —— 做了就是把我自己的审美塞进代码。
    """
    fills = [it.get("fill") for it in spec["layout"] if it.get("fill")]
    uniq = sorted(set(fills))
    merges = []
    for i, a in enumerate(uniq):
        for b in uniq[i + 1:]:
            if not _near(a, b):
                continue
            na, nb = fills.count(a), fills.count(b)
            keep, drop = (a, b) if na >= nb else (b, a)
            for it in spec["layout"]:
                if it.get("fill") == drop:
                    it["fill"] = keep
            merges.append("%s → %s（差 %d/255，留多数派）"
                          % (drop, keep, max(abs(x - y) for x, y in
                                             zip(_rgb(a), _rgb(b)))))
    return ["近重复填充色合并：%s" % "；".join(merges)] if merges else []


def _rows_in_group(spec):
    """找出每个分组里"排成一行"的成员（y 基本一致）。"""
    by = {it["id"]: it for it in spec["layout"]}
    out = []
    for g in spec.get("groups") or []:
        ms = [by[it["id"]] for it in spec["layout"]
              if it.get("for") in set(g.get("members") or [])
              and by.get(it["id"], {}).get("box")
              and it.get("kind") not in ("line", "verbatim")]
        if len(ms) < 3:
            continue
        ys = [m["box"][1] for m in ms]
        if max(ys) - min(ys) > 6:      # 不在同一行的，不碰
            continue
        ms.sort(key=lambda m: m["box"][0])
        out.append((g, ms))
    return out


def r_align_rows(spec):
    """"同类元素该在同一条线上" —— 把一行的 y 对齐到多数派。"""
    done = []
    for g, ms in _rows_in_group(spec):
        ys = Counter(round(m["box"][1], 1) for m in ms)
        if len(ys) < 2:
            continue
        top = ys.most_common(1)[0][0]
        for m in ms:
            m["box"][1] = top
        done.append("%s 的 %d 个成员 y 对齐到 %.1f" % (g["id"], len(ms), top))
    return done


def r_even_spacing(spec):
    """"同一行里的间距应当一致" —— 均分，并让左右外边距也一样宽。

    只在**同一行、宽度一样**的成员之间做 —— 宽度不一样时"均分"没有意义。
    """
    done = []
    for g, ms in _rows_in_group(spec):
        w = ms[0]["box"][2]
        if any(abs(m["box"][2] - w) > 0.5 for m in ms):
            continue
        # 容器的框：找分组自己的那个 rect
        cont = [it for it in spec["layout"]
                if it.get("for") == g["id"] and it.get("kind") in ("rect", "roundrect")]
        if not cont:
            continue
        cb = cont[0]["box"]
        n = len(ms)
        gap = (cb[2] - w * n) / (n + 1)
        if gap < 2:
            continue
        old = [round(ms[i + 1]["box"][0] - (ms[i]["box"][0] + w), 1)
               for i in range(n - 1)]
        for i, m in enumerate(ms):
            m["box"][0] = round(cb[0] + gap * (i + 1) + w * i, 1)
        done.append("%s 的 %d 个成员间距均分成 %.1f（原来 %s）"
                    % (g["id"], n, gap, old))
    return done


def r_center_labels(spec):
    """"说明文字要落在它标注的那条线上"。

    ⚠️ 只对**单段直线**做。折线没有唯一中轴，规则不好定 —— 跳过并报出来。
    """
    lines = defaultdict(list)
    for it in spec["layout"]:
        if it.get("kind") == "line":
            lines[it.get("for")].append(it)
    done, skipped = [], set()
    for it in spec["layout"]:
        t = it.get("text") or {}
        if it.get("kind") != "text" or not it.get("box") or not it.get("annotates"):
            continue
        segs = lines.get(it["annotates"], [])
        if len(segs) != 1 or (segs[0].get("prst") or "straightConnector1") \
                != "straightConnector1":
            skipped.add(it["annotates"])
            continue
        ln = segs[0]
        p1, p2 = ln["p1"], ln["p2"]
        b = it["box"]
        cx, cy = b[0] + b[2] / 2.0, b[1] + b[3] / 2.0
        if abs(p1[0] - p2[0]) < 1.0:            # 竖线 → 对 x
            d = p1[0] - cx
            b[0] = round(b[0] + d, 1)
        elif abs(p1[1] - p2[1]) < 1.0:          # 横线 → 对 y
            d = p1[1] - cy
            b[1] = round(b[1] + d, 1)
        else:
            skipped.add(it["annotates"])
            continue
        if abs(d) > 0.05:
            done.append("%s 挪 %.1fpx 到 %s 上" % (it["id"], d, it["annotates"]))
    if skipped:
        done.append("⚠ 折线/多段的关系**没查也没挪**：%s（没有唯一中轴）"
                    % "、".join(sorted(skipped)))
    return done


def r_text_padding(spec):
    """"文字离框边够不够" —— 不够就把**框加高**。

    只往下长（保持上边不动），免得整块往上飘去撞别的东西。
    ⚠️ 它可能撞到下面的东西 —— 所以跑完必须再过一遍 B 类检查（出界/重叠）。
    """
    done = []
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
        ppm = 150.0 / 72.0
        lh = pt * (float(t.get("lnspc") or 120000) / 100000.0) * ppm
        ink = (n - 1) * lh + 0.88 * pt * ppm
        # +0.6：刚好等于 MIN_PAD 时，浮点误差会让检查仍判"<8"
        need = ink + 2 * MIN_PAD + 0.6
        if need > b[3] + 0.5:
            done.append("%s 框高 %.0f → %.0f（折 %d 行要 %.0f + 上下各 %.0f）"
                        % (it["id"], b[3], round(need, 1), n, ink, MIN_PAD))
            b[3] = round(need, 1)
    return done


def r_expand_containers(spec):
    """"容器里的东西不该露出容器外" —— 把容器**扩大**到装得下。

    ★ 扩容器，不是挪内容。挪内容会动到意义相关的相对位置，风险大。
    """
    done = []
    for g in spec.get("groups") or []:
        cont = [it for it in spec["layout"]
                if it.get("for") == g["id"] and it.get("kind") in ("rect", "roundrect")]
        if not cont:
            continue
        cb = cont[0]["box"]
        kids = [it for it in spec["layout"]
                if it.get("for") == g["id"] and it.get("box") and it is not cont[0]]
        if not kids:
            continue
        x0 = min([cb[0]] + [k["box"][0] for k in kids])
        y0 = min([cb[1]] + [k["box"][1] for k in kids])
        x1 = max([cb[0] + cb[2]] + [k["box"][0] + k["box"][2] for k in kids])
        y1 = max([cb[1] + cb[3]] + [k["box"][1] + k["box"][3] for k in kids])
        if (x0 < cb[0] - 0.5 or y0 < cb[1] - 0.5
                or x1 > cb[0] + cb[2] + 0.5 or y1 > cb[1] + cb[3] + 0.5):
            old = list(cb)
            cb[0], cb[1] = round(x0, 1), round(y0, 1)
            cb[2], cb[3] = round(x1 - x0, 1), round(y1 - y0, 1)
            done.append("%s 容器 %s → %s（装下露出来的成员）"
                        % (g["id"], [round(v) for v in old], [round(v) for v in cb]))
    return done


def r_canvas_padding(spec, pad):
    """"内容别顶到画布边上" —— 整体留一圈白。**放最后做。**

    ★ 必须**幂等**。第一版是"无条件加一圈"，于是跑第二遍又加一圈 ——
      画布越跑越大。实测发现的（把结果再喂进去，47 个对象全又平移了 24px）。
      改成：每次都**从内容包围盒重新算**（内容左上角落到 (pad,pad)，画布 = 内容 + 2×pad）。
      跑几遍结果都一样。
    """
    items = [it for it in spec["layout"] if it.get("box") or it.get("p1")]
    if not items or pad <= 0:
        return []
    xs0, ys0, xs1, ys1 = [], [], [], []
    for it in items:
        if it.get("box"):
            b = it["box"]
            xs0.append(b[0]); ys0.append(b[1])
            xs1.append(b[0] + b[2]); ys1.append(b[1] + b[3])
        for k in ("p1", "p2"):
            if it.get(k):
                xs0.append(it[k][0]); ys0.append(it[k][1])
                xs1.append(it[k][0]); ys1.append(it[k][1])
    x0, y0, x1, y1 = min(xs0), min(ys0), max(xs1), max(ys1)
    dx, dy = round(pad - x0, 1), round(pad - y0, 1)
    for it in items:
        if it.get("box"):
            it["box"] = [round(it["box"][0] + dx, 1), round(it["box"][1] + dy, 1),
                         it["box"][2], it["box"][3]]
        for k in ("p1", "p2"):
            if it.get(k):
                it[k] = [round(it[k][0] + dx, 1), round(it[k][1] + dy, 1)]
    cw, ch = spec["source"]["canvas"]
    nw, nh = round(x1 - x0 + 2 * pad), round(y1 - y0 + 2 * pad)
    spec["source"]["canvas"] = [nw, nh]
    if abs(dx) < 0.05 and abs(dy) < 0.05 and cw == nw and ch == nh:
        return []
    return ["画布 %dx%d → %dx%d（内容包围盒 + 四边 %.0f；平移 %+.1f,%+.1f）"
            % (cw, ch, nw, nh, pad, dx, dy)]


RULES = [("线宽/线色", r_line_style),
         ("内边距", r_unify_insets),
         ("近重复颜色", r_merge_near_colors),
         ("行内对齐", r_align_rows),
         ("间距均分", r_even_spacing),
         ("标注居中", r_center_labels),
         ("文字留白", r_text_padding),
         ("容器包含", r_expand_containers)]


def normalize(spec, pad=24.0, only=None):
    log = []
    for name, fn in RULES:
        if only and name not in only:
            continue
        out = fn(spec)
        if out:
            log.append((name, out))
    out = r_canvas_padding(spec, pad)
    if out:
        log.append(("画布留白", out))
    # 台账：改了排版就留痕
    spec.setdefault("changes", []).append({
        "at": "", "round": (max([c.get("round", 0) for c in spec["changes"]] or [0]) + 1),
        "target": "layout（排版）", "by": "ai",
        "from": "照抄原件坐标", "to": "按设计规则自动拉齐",
        "reason": "normalize_layout.py；只动画法，内容/关系未动",
        "detail": [x for _n, xs in log for x in xs]})
    return log


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("spec")
    ap.add_argument("--out", required=True)
    ap.add_argument("--pad", type=float, default=24.0, help="画布留白，0 表示不加")
    ap.add_argument("--only", help="只跑这几条，逗号分隔")
    ap.add_argument("--dry-run", action="store_true", help="只报会改什么，不写文件")
    args = ap.parse_args()

    spec = load(args.spec)
    log = normalize(spec, args.pad, set(args.only.split(",")) if args.only else None)

    print("=" * 66)
    print("normalize_layout —— 按设计规则拉齐排版")
    print("  只动画法：坐标 / 尺寸 / 线宽线色 / 内边距 / 填充")
    print("  不碰：内容、关系、分组（那是意义）")
    print("=" * 66)
    if not log:
        print("\n✓ 没有可拉齐的地方（清单已经是齐的）")
    for name, items in log:
        print("\n── %s" % name)
        for x in items:
            print("   %s %s" % ("⚠" if x.startswith("⚠") else "·", x))
    print()
    print("=" * 66)
    print("⚠ 跑完**必须**再过一遍 B 类检查（spec_check + check_a）：")
    print("  加高/扩容器可能撞到别的东西，这里不负责收尾。")
    print("=" * 66)

    if args.dry_run:
        print("\n（--dry-run：没有写文件）")
        return 0
    with open(args.out, "w", encoding="utf-8") as f:
        json.dump(spec, f, ensure_ascii=False, indent=1)
    print("\n-> %s" % args.out)
    return 0


if __name__ == "__main__":
    sys.exit(main())
