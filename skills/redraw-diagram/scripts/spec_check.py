#!/usr/bin/env python3
"""
spec_check.py —— 校验「清单」（spec.json）自洽。

★ 它保证的是**清单自洽**，不是**清单对**。
   "我读得对不对"要靠 Check B（回原件对账），那是另一件事。
   这里只回答一个问题：**这份清单自己站得住吗？**

为什么需要它：
   清单是整个流程的枢纽。它内部矛盾的话，后面每一步都在错的东西上干活，
   而且**错得很安静**——没人会去读一份 JSON 找矛盾。

用法：
    python3 spec_check.py work/spec.json            # 只校验
    python3 spec_check.py work/spec.json --summary  # 校验 + 人读的概览
"""

import argparse
import json
import sys
from collections import Counter, defaultdict

SPEC_VERSION = "0.2"

PROVENANCE = ("original", "measured", "composed", "inferred", "domain", "user")
PROV_CN = {"original": "原件明写", "measured": "我量的", "composed": "我拼的",
           "inferred": "我推的", "domain": "行内常识", "user": "用户说的"}
ROLES = ("shape", "container", "relation_carrier", "annotation", "decoration")
STATUSES = ("unconfirmed", "confirmed", "changed")
REL_KINDS = ("sequence", "dependency", "containment", "association")
REL_DIRS = ("forward", "backward", "both", "none")
LAYOUT_KINDS = ("rect", "roundrect", "diamond", "ellipse", "line", "brace",
                "arrow", "picture", "text", "container", "verbatim")
BOX_KINDS = ("rect", "roundrect", "diamond", "ellipse", "brace", "arrow",
             "picture", "text", "container")


# --------------------------------------------------------------------------
def load(path):
    with open(path, encoding="utf-8") as f:
        return json.load(f)


def save(spec, path):
    spec.setdefault("spec_version", SPEC_VERSION)
    with open(path, "w", encoding="utf-8") as f:
        json.dump(spec, f, ensure_ascii=False, indent=1)
    return path


# --------------------------------------------------------------------------
class Problems:
    def __init__(self):
        self.items = []

    def add(self, where, msg):
        self.items.append((where, msg))

    def __bool__(self):
        return bool(self.items)

    def __len__(self):
        return len(self.items)


def check(spec):
    """返回 Problems。空 = 自洽。"""
    p = Problems()

    # ---- 1. 顶层 -------------------------------------------------------
    if spec.get("spec_version") != SPEC_VERSION:
        p.add("spec_version",
              "应该是 %r，实际 %r" % (SPEC_VERSION, spec.get("spec_version")))
    for k in ("source", "entities", "relations", "groups", "layout"):
        if k not in spec:
            p.add(k, "缺这个字段")
        elif not isinstance(spec[k], list) and k != "source":
            p.add(k, "应该是数组")
    if not isinstance(spec.get("source"), dict):
        p.add("source", "应该是对象")
    for k in ("changes", "open_questions", "findings", "unsupported"):
        if k in spec and not isinstance(spec[k], list):
            p.add(k, "应该是数组")

    src = spec.get("source") or {}
    if not isinstance(src.get("canvas"), list) or len(src.get("canvas") or []) != 2:
        p.add("source.canvas", "应该是 [宽, 高]")

    # ---- 2. id 全局唯一 -------------------------------------------------
    ids = {}
    for table in ("entities", "relations", "groups", "layout"):
        for i, it in enumerate(spec.get(table) or []):
            if not isinstance(it, dict):
                p.add("%s[%d]" % (table, i), "应该是对象")
                continue
            i_d = it.get("id")
            if not i_d:
                p.add("%s[%d]" % (table, i), "缺 id")
                continue
            if i_d in ids:
                p.add("%s[%d].id" % (table, i),
                      "id %r 和 %s 里的撞了（id 必须全局唯一）" % (i_d, ids[i_d]))
            else:
                ids[i_d] = "%s[%d]" % (table, i)

    ent_ids = {e.get("id") for e in (spec.get("entities") or []) if isinstance(e, dict)}
    grp_ids = {g.get("id") for g in (spec.get("groups") or []) if isinstance(g, dict)}
    rel_ids = {r.get("id") for r in (spec.get("relations") or []) if isinstance(r, dict)}
    meaning_ids = ent_ids | grp_ids | rel_ids

    # ---- 3/4. 引用存在 --------------------------------------------------
    for i, r in enumerate(spec.get("relations") or []):
        if not isinstance(r, dict):
            continue
        at = "relations[%d]" % i
        for side in ("from", "to"):
            v = r.get(side)
            if not v:
                p.add("%s.%s" % (at, side), "缺")
            elif v not in ent_ids and v not in grp_ids:
                p.add("%s.%s" % (at, side), "指向不存在的实体/分组：%r" % v)
        _enum(p, at + ".kind", r.get("kind"), REL_KINDS)
        _enum(p, at + ".direction", r.get("direction"), REL_DIRS)
        _enum(p, at + ".status", r.get("status"), STATUSES)

    for i, g in enumerate(spec.get("groups") or []):
        if not isinstance(g, dict):
            continue
        at = "groups[%d]" % i
        for m in g.get("members") or []:
            if m not in ent_ids:
                p.add(at + ".members", "指向不存在的实体：%r" % m)

    # ---- 5/10. layout --------------------------------------------------
    for i, s in enumerate(spec.get("layout") or []):
        if not isinstance(s, dict):
            continue
        at = "layout[%d]" % i
        f = s.get("for")
        if not f:
            p.add(at + ".for", "缺（每个输出对象都要说清它在表达谁）")
        elif f not in meaning_ids:
            p.add(at + ".for", "指向不存在的东西：%r" % f)
        kind = _enum(p, at + ".kind", s.get("kind"), LAYOUT_KINDS)
        # ★ z（叠放次序）是**必须**的。实测栽过：漏了它的条目按 0 排，
        #   会被后面的不透明形状整个盖住 —— 位置文字全对，就是看不见，
        #   而 Check A 只查位置文字，结构上发现不了。
        if s.get("z") is None:
            p.add(at + ".z", "缺少 z（叠放次序）：没有它就保证不了画的次序和原件一致")
        # ★ 所有对象必须在画布内。出界 = 被裁掉 = **内容真的少了**，
        #   实测栽过：声明的范围漏了 4 个灰色底框，它们画到画布外被裁。
        _cv = src.get("canvas") or [0, 0]
        cw, ch = (_cv + [0, 0])[:2]
        pts = []
        if kind != "line" and isinstance(s.get("box"), list) and len(s["box"]) == 4:
            b = s["box"]
            pts = [(b[0], b[1]), (b[0] + b[2], b[1] + b[3])]
        elif isinstance(s.get("p1"), list) and isinstance(s.get("p2"), list):
            pts = [(s["p1"][0], s["p1"][1]), (s["p2"][0], s["p2"][1])]
        for (px, py) in pts:
            if px < -0.5 or py < -0.5 or px > cw + 0.5 or py > ch + 0.5:
                p.add(at + ".box",
                      "对象出界（点 %.0f,%.0f 不在画布 %.0f×%.0f 内）—— "
                      "出界会被裁掉，等于**内容少了**" % (px, py, cw, ch))
                break
        if kind == "verbatim" and not s.get("src"):
            p.add(at + ".src", "kind=verbatim（原样搬运）必须给 src（要搬的那份 XML）")
        has_box = "box" in s
        has_pts = "p1" in s or "p2" in s
        if kind == "line":
            if has_box:
                p.add(at, "line 用 p1/p2，不该有 box")
            if not (s.get("p1") and s.get("p2")):
                p.add(at, "line 缺 p1/p2")
        else:
            if has_pts:
                p.add(at, "非 line 不该有 p1/p2")
            if not has_box:
                p.add(at, "缺 box")
            elif not (isinstance(s.get("box"), list) and len(s["box"]) == 4):
                p.add(at + ".box", "应该是 [x, y, 宽, 高]")

    # ---- 6. 每个意义单位都要有人表达 ------------------------------------
    expressed = defaultdict(list)
    for s in spec.get("layout") or []:
        if isinstance(s, dict) and s.get("for"):
            expressed[s["for"]].append(s.get("id"))
    for table, label in (("entities", "实体"), ("relations", "关系"), ("groups", "分组")):
        for i, it in enumerate(spec.get(table) or []):
            if not isinstance(it, dict):
                continue
            i_d = it.get("id")
            if not i_d:
                continue
            n = len(expressed.get(i_d, []))
            if n == 0 and not it.get("not_drawn"):
                p.add("%s[%d] (%s)" % (table, i, i_d),
                      "%s没被任何 layout 对象表达。**如果是故意不画，显式写 not_drawn: true**"
                      % label)
            if n > 0 and it.get("not_drawn"):
                p.add("%s[%d] (%s)" % (table, i, i_d),
                      "标了 not_drawn，但有 %d 个 layout 对象在表达它" % n)

    # ---- 7/8. provenance 与 evidence -----------------------------------
    for table in ("entities", "relations", "groups", "layout"):
        for i, it in enumerate(spec.get(table) or []):
            if not isinstance(it, dict):
                continue
            at = "%s[%d]" % (table, i)
            prov = _enum(p, at + ".provenance", it.get("provenance"), PROVENANCE)
            ev = it.get("evidence")
            if prov and prov != "user":
                if not isinstance(ev, list) or not [x for x in ev if str(x).strip()]:
                    p.add(at + ".evidence",
                          "provenance=%r 必须给证据（只有 user 可以不给）" % prov)
            if table == "entities":
                _enum(p, at + ".status", it.get("status"), STATUSES)
                for r in it.get("role") or []:
                    if r not in ROLES:
                        p.add(at + ".role", "未知取值 %r" % r)

    # ---- 9. changes 是追加型 -------------------------------------------
    for i, c in enumerate(spec.get("changes") or []):
        if not isinstance(c, dict):
            continue
        at = "changes[%d]" % i
        for k in ("target", "from", "to", "by"):
            if k not in c:
                p.add(at, "缺 %s" % k)
        if c.get("by") not in ("user", "ai"):
            p.add(at + ".by", "只能是 user 或 ai，实际 %r" % c.get("by"))
        if c.get("by") == "user" and not c.get("evidence_ref"):
            p.add(at, "用户做的改动必须指向原话（evidence_ref）")

    return p


def _enum(p, where, value, allowed):
    if value is None:
        p.add(where, "缺（可选：%s）" % " / ".join(allowed))
        return None
    if value not in allowed:
        p.add(where, "未知取值 %r（可选：%s）" % (value, " / ".join(allowed)))
        return None
    return value


# --------------------------------------------------------------------------
def summary(spec):
    out = []
    src = spec.get("source") or {}
    out.append("画布 %s %s" % (src.get("canvas"),
                             "（%s）" % src.get("unit") if src.get("unit") else ""))
    for f in src.get("files") or []:
        out.append("  来源：%s  %s%s" % (
            f.get("path"), f.get("kind"),
            "  第 %s 页" % f.get("slide") if f.get("slide") else ""))

    counts = {t: len(spec.get(t) or [])
              for t in ("entities", "relations", "groups", "layout")}
    out.append("")
    out.append("层 1（在说什么）  实体 %d  关系 %d  分组 %d"
               % (counts["entities"], counts["relations"], counts["groups"]))
    out.append("层 2（长什么样）  %d 个对象" % counts["layout"])
    out.append("")

    prov = Counter()
    for t in ("entities", "relations", "groups", "layout"):
        for it in spec.get(t) or []:
            if isinstance(it, dict):
                prov[it.get("provenance")] += 1
    out.append("信息来源：")
    for k in PROVENANCE:
        if prov.get(k):
            out.append("   %-10s %-8s %d" % (k, PROV_CN[k], prov[k]))
    out.append("")

    un = [e.get("id") for e in spec.get("entities") or []
          if isinstance(e, dict) and e.get("status") == "unconfirmed"]
    ch = [e.get("id") for e in spec.get("entities") or []
          if isinstance(e, dict) and e.get("status") == "changed"]
    if un:
        out.append("未确认：%d 个  %s" % (len(un), " ".join(un[:12])))
    if ch:
        out.append("用户改过：%d 个  %s" % (len(ch), " ".join(ch[:12])))

    qs = spec.get("open_questions") or []
    if qs:
        out.append("")
        out.append("待问用户（%d）：" % len(qs))
        for q in qs:
            out.append("   [%s] %s" % (q.get("id"), q.get("ask")))
            if q.get("if_wrong"):
                out.append("        答错会改：%s" % q["if_wrong"])
            if q.get("default"):
                out.append("        默认：%s" % q["default"])
            if q.get("status"):
                out.append("        状态：%s" % q["status"])

    fs = spec.get("findings") or []
    if fs:
        out.append("")
        out.append("关于原件的发现（%d）——这些**不是我们的错**：" % len(fs))
        for f in fs:
            out.append("   [%s] %s" % (f.get("id"), f.get("what")))

    us = spec.get("unsupported") or []
    if us:
        out.append("")
        out.append("画不出来的（%d）：" % len(us))
        for u in us:
            out.append("   %s —— %s" % (u.get("what"), u.get("reason")))

    return "\n".join(out)


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("spec")
    ap.add_argument("--summary", action="store_true", help="额外打印人读的概览")
    args = ap.parse_args()

    spec = load(args.spec)
    probs = check(spec)

    if args.summary:
        print(summary(spec))
        print()

    if probs:
        print("✗ 清单不自洽，%d 处：" % len(probs))
        for where, msg in probs.items:
            print("  %-28s %s" % (where, msg))
        sys.exit(1)

    print("✓ 清单自洽（%d 个实体 / %d 条关系 / %d 个对象）"
          % (len(spec.get("entities") or []),
             len(spec.get("relations") or []),
             len(spec.get("layout") or [])))


if __name__ == "__main__":
    main()
