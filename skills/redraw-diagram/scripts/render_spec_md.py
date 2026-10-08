#!/usr/bin/env python3
"""
render_spec_md.py —— 把最终清单渲染成**人读版**（spec.md）。

★ 这是「视图」，不是产物。只有 spec.json 能被编辑，视图永远是渲染出来的。
  头部会记下**源文件的哈希** —— 改了源忘了重渲染，一眼能看出来。

用法：
    python3 render_spec_md.py work/spec.json --out work/spec.md
"""

import argparse
import hashlib
import json
import os
import sys
from collections import Counter, defaultdict

PROV_CN = {"original": "原件明写", "measured": "我量出来的",
           "composed": "我拼出来的", "inferred": "我推的",
           "domain": "行内常识补的", "user": "用户说的"}
ROLE_CN = {"shape": "流程节点 / 形状", "container": "容器", "relation_carrier": "关系载体",
           "annotation": "说明文字", "decoration": "装饰"}
KIND_CN = {"rect": "矩形", "roundrect": "圆角矩形", "diamond": "菱形",
           "ellipse": "椭圆", "line": "线", "brace": "大括号", "arrow": "箭头",
           "picture": "图片", "text": "文本框", "container": "容器",
           "verbatim": "**原样搬运**"}


def load(p):
    with open(p, encoding="utf-8") as f:
        return json.load(f)


def sha8(p):
    return hashlib.sha256(open(p, "rb").read()).hexdigest()[:8]


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("spec")
    ap.add_argument("--out", required=True)
    args = ap.parse_args()

    S = load(args.spec)
    src = S.get("source") or {}
    L = []
    A = L.append

    A("# 清单（人读版）")
    A("")
    A("> ⚠️ **这是从 `%s` 渲染出来的视图，不许直接改。**"
      % os.path.basename(args.spec))
    A("> 源文件 sha256 前 8 位：`%s`　要改就改源，然后重新渲染这一份。"
      % sha8(args.spec))
    A("")
    A("| | |")
    A("|---|---|")
    A("| 画布 | %s %s |" % (src.get("canvas"), src.get("unit") or ""))
    for f in src.get("files") or []:
        A("| 来源 | `%s` %s%s |" % (
            f.get("path"), f.get("kind"),
            "　范围 %s" % (f.get("region"),) if f.get("region") else ""))
    A("| 层 1 | 实体 %d　关系 %d　分组 %d |"
      % (len(S.get("entities") or []), len(S.get("relations") or []),
         len(S.get("groups") or [])))
    A("| 层 2 | %d 个对象 |" % len(S.get("layout") or []))
    A("")

    # ---- 来源分布：这是用户核对时该盯的 ----
    prov = Counter()
    for t in ("entities", "relations", "groups", "layout"):
        for it in S.get(t) or []:
            prov[it.get("provenance")] += 1
    A("## 信息来源")
    A("")
    A("| 档 | 意思 | 条数 |")
    A("|---|---|---|")
    for k, cn in PROV_CN.items():
        if prov.get(k):
            A("| `%s` | %s | %d |" % (k, cn, prov[k]))
    A("")
    weak = prov.get("inferred", 0) + prov.get("domain", 0)
    if weak:
        A("> ⚠️ **其中有 %d 条是我推的 / 行内常识补的** —— 这几条最该核。" % weak)
        A("")

    # ---- 容器 / 分组 ----
    G = S.get("groups") or []
    if G:
        A("## 分区 / 容器（%d）" % len(G))
        A("")
        A("| 编号 | 说的是什么 | 包含 |")
        A("|---|---|---|")
        for g in G:
            A("| **%s** | %s | %s |"
              % (g["id"], g.get("label") or "", "、".join(g.get("members") or []) or "—"))
        A("")

    # ---- 实体 ----
    E = S.get("entities") or []
    by_role = defaultdict(list)
    for e in E:
        by_role[(e.get("role") or ["shape"])[0]].append(e)
    A("## 实体（%d）" % len(E))
    A("")
    for role in ("shape", "container", "relation_carrier", "annotation", "decoration"):
        grp = by_role.get(role)
        if not grp:
            continue
        A("### %s（%d）" % (ROLE_CN.get(role, role), len(grp)))
        A("")
        A("| 编号 | 文字 | 来源 | 状态 |")
        A("|---|---|---|---|")
        for e in grp:
            t = e.get("text") or "—"
            if e.get("text_user"):
                t = "~~%s~~ → **%s**" % (e.get("text") or "", e["text_user"])
            A("| **%s** | %s | %s | %s |"
              % (e["id"], t.replace("\n", "⏎"), PROV_CN.get(e.get("provenance"), ""),
                 {"unconfirmed": "待确认", "confirmed": "已确认",
                  "changed": "**用户改过**"}.get(e.get("status"), "")))
        A("")
    if not E:
        A("（空）")
        A("")

    # ---- 关系 ----
    R = S.get("relations") or []
    A("## 关系（%d）" % len(R))
    A("")
    A("| 编号 | 从 | 到 | 方向 | 来源 |")
    A("|---|---|---|---|---|")
    for r in R:
        A("| **%s** | %s | %s | %s | %s |"
          % (r["id"], r.get("from"), r.get("to"), r.get("direction"),
             PROV_CN.get(r.get("provenance"), "")))
    A("")

    # ---- 层 2 概况 ----
    k = Counter(it.get("kind") for it in S.get("layout") or [])
    A("## 层 2：怎么画的（%d 个对象）" % sum(k.values()))
    A("")
    A("| 画法 | 个数 |")
    A("|---|---|")
    for kind, n in k.most_common():
        A("| %s | %d |" % (KIND_CN.get(kind, kind), n))
    A("")
    vb = [it for it in S.get("layout") or [] if it.get("kind") == "verbatim"]
    if vb:
        A("> ⚠️ **其中 %d 个是原样搬运的**（不是重画的）。它们**保真**，"
          "但点开是**一堆节点**，改起来不如预设形状顺手。" % len(vb))
        A("")

    # ---- 台账 ----
    C = S.get("changes") or []
    if C:
        A("## 改动台账（%d）—— 我动过原件什么" % len(C))
        A("")
        for c in C:
            A("- **%s**：`%s` → `%s`　（%s）" % (
                c.get("target"), c.get("from"), c.get("to"),
                "用户" if c.get("by") == "user" else "AI"))
            if c.get("reason"):
                A("  - 理由：%s" % c["reason"])
        A("")

    # ---- 发现 / 画不出来 ----
    F = S.get("findings") or []
    if F:
        A("## 关于原件的发现（%d）—— **不是我的错**" % len(F))
        A("")
        for f in F:
            A("- **%s**：%s" % (f.get("id"), f.get("what")))
        A("")
    U = S.get("unsupported") or []
    if U:
        A("## 画不出来 / 范围外（%d）" % len(U))
        A("")
        agg = Counter((u.get("what"), u.get("reason")) for u in U)
        for (what, why), n in agg.most_common():
            A("- %s ×%d　— %s" % (what, n, why))
        A("")

    with open(args.out, "w", encoding="utf-8") as f:
        f.write("\n".join(L))
    print("人读版清单 -> %s（源 %s，sha %s）"
          % (args.out, os.path.basename(args.spec), sha8(args.spec)))


if __name__ == "__main__":
    main()
