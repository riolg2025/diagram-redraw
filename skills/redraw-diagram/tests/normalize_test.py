#!/usr/bin/env python3
"""
normalize_test.py —— 证明 normalize_layout 干的是它说的那件事。

四条断言：
  ① **不动意义**：entities / relations / groups 前后必须一模一样
  ② **幂等**：跑一遍和跑两遍结果相同（踩过坑：第一版"无条件加留白"，
     跑第二遍又加一圈，画布越跑越大）
  ③ **真的会拉齐**：造一份不齐的，规则该动的地方都动了
  ④ **不该动的别动**：线条有两种颜色但**没有多数派**时，不许抹平

用法：
    python3 tests/normalize_test.py
"""

import copy
import json
import os
import subprocess
import sys
import tempfile

HERE = os.path.dirname(os.path.abspath(__file__))
SCRIPTS = os.path.join(os.path.dirname(HERE), "scripts")
PY = sys.executable
sys.path.insert(0, SCRIPTS)

import normalize_layout as NL  # noqa: E402


def load(p):
    with open(p, encoding="utf-8") as f:
        return json.load(f)


def strip(s):
    s = copy.deepcopy(s)
    s.pop("changes", None)
    return s


def base_spec():
    """一份最小但"不齐"的清单：两种线宽、两种内边距、两条标签没居中。"""
    return {
        "spec_version": "0.2",
        "task": {"id": "t", "round": 1, "created": ""},
        "source": {"canvas": [1000, 600], "unit": "px@150dpi",
                   "files": [{"path": "x.pptx", "kind": "pptx_shapes",
                              "slide": 1, "region": [0, 0, 1000, 600]}]},
        "entities": [
            {"id": "N1", "role": ["shape"], "text": "甲", "box_src": [10, 10, 50, 50],
             "provenance": "original", "evidence": ["p:sp#1"], "status": "confirmed"},
            {"id": "N2", "role": ["shape"], "text": "乙", "box_src": [300, 10, 50, 50],
             "provenance": "original", "evidence": ["p:sp#2"], "status": "confirmed"},
            {"id": "T1", "role": ["annotation"], "text": "说明", "box_src": [0, 100, 80, 40],
             "provenance": "original", "evidence": ["p:sp#3"], "status": "confirmed"},
        ],
        "relations": [
            {"id": "C1", "from": "N1", "to": "N2", "kind": "sequence",
             "direction": "forward", "provenance": "measured",
             "evidence": ["p:cxnSp#4"], "status": "confirmed"},
        ],
        "groups": [],
        "layout": [
            {"id": "s1", "for": "N1", "kind": "rect", "box": [50.0, 50.0, 60.0, 60.0],
             "z": 0, "provenance": "measured", "evidence": ["p:sp#1"],
             "line": {"color": "0084CF", "pt": 1.75, "dash": "solid"}},
            {"id": "s2", "for": "N2", "kind": "rect", "box": [340.0, 50.0, 60.0, 60.0],
             "z": 1, "provenance": "measured", "evidence": ["p:sp#2"],
             "line": {"color": "0084CF", "pt": 1.75, "dash": "solid"}},
            {"id": "s3", "for": "T1", "kind": "text", "box": [40.0, 150.0, 80.0, 40.0],
             "z": 2, "provenance": "measured", "evidence": ["p:sp#3"],
             "annotates": "C1", "text": {"lines": ["说明"], "font_pt": 10.0,
                                         "wrap": True, "insets": [4.0, 4.0, 4.0, 4.0]}},
            {"id": "s4", "for": "C1", "kind": "line", "prst": "straightConnector1",
             "p1": [200.0, 200.0], "p2": [200.0, 500.0], "arrow": "end", "z": 3,
             "provenance": "measured", "evidence": ["p:cxnSp#4"],
             "line": {"color": "FF0000", "pt": 0.75, "dash": "solid"}},
        ],
        "changes": [], "open_questions": [], "findings": [], "unsupported": [],
    }


def main():
    ok = True

    # ---- ① 不动意义 + ② 幂等 -------------------------------------------
    a = base_spec()
    b = copy.deepcopy(a)
    NL.normalize(b, pad=24.0)
    c = copy.deepcopy(b)
    NL.normalize(c, pad=24.0)

    for tbl in ("entities", "relations", "groups"):
        if a[tbl] != b[tbl]:
            print("✗ ① %s 被改了 —— 归一化**不许动意义**" % tbl)
            ok = False
    if ok:
        print("① 不动意义（entities / relations / groups）✓")

    if strip(b) == strip(c):
        print("② 幂等（跑一遍 = 跑两遍）✓")
    else:
        print("✗ ② 不幂等 —— 跑第二遍又改了东西")
        ok = False

    # ---- ③ 该拉齐的拉齐了 ----------------------------------------------
    moved = b["layout"][2]["box"][0] != a["layout"][2]["box"][0]
    print("③ 标注居中：%s" % ("✓ 挪了" if moved else "✗ 没动"))
    ok = ok and moved
    # ★ 别断言"画布必须变大" —— 规则是「画布 = 内容包围盒 + 2×留白」，
    #   内容比画布小的时候画布**会变小**，那是对的。
    #   真正该成立的是：**内容左上角正好落在 (pad,pad)**，且四边留白都是 pad。
    # ⚠️ 包围盒必须**连线的端点一起算** —— 只算 box 会把线排除在外，
    #    下边留白就会算成几百像素（我第一版就是这样，误判成失败）。
    xs0, ys0, xs1, ys1 = [], [], [], []
    for it in b["layout"]:
        if it.get("box"):
            bb = it["box"]
            xs0.append(bb[0]); ys0.append(bb[1])
            xs1.append(bb[0] + bb[2]); ys1.append(bb[1] + bb[3])
        for k in ("p1", "p2"):
            if it.get(k):
                xs0.append(it[k][0]); ys0.append(it[k][1])
                xs1.append(it[k][0]); ys1.append(it[k][1])
    cw, ch = b["source"]["canvas"]
    pads = {"左": min(xs0), "上": min(ys0), "右": cw - max(xs1), "下": ch - max(ys1)}
    if all(abs(v - 24.0) < 0.2 for v in pads.values()):
        print("   四边留白都正好 24：%s ✓" % {k: round(v, 1) for k, v in pads.items()})
    else:
        print("✗ 四边留白不对：%s" % {k: round(v, 1) for k, v in pads.items()})
        ok = False

    # ---- ④ 没有多数派就别抹平 ------------------------------------------
    d = base_spec()
    # 三红两蓝：红的 60%，不到 70% → 不该统一
    d["layout"][0]["line"]["color"] = "FF0000"
    d["layout"][1]["line"]["color"] = "FF0000"
    d["layout"][2]["line"] = {"color": "00FF00", "pt": 1.75, "dash": "solid"}
    d["layout"][3]["line"]["color"] = "0000FF"
    NL.normalize(d, pad=0)
    colors = {it["line"]["color"] for it in d["layout"] if it.get("line")}
    if len(colors) > 1:
        print("④ 线条颜色没有多数派时**没抹平**（留了 %d 种）✓" % len(colors))
    else:
        print("✗ ④ 把没有多数派的颜色也抹平了 —— 会抹掉有意的差异")
        ok = False

    print()
    print("normalize_test：%s" % ("全部通过 ✓" if ok else "有失败 ✗"))
    return 0 if ok else 1


if __name__ == "__main__":
    sys.exit(main())
