#!/usr/bin/env python3
"""
draft_spec.py —— 观察 → 清单**草稿**。

★ 它是草稿，不是成品。草稿的意思是：

    层 2（长什么样）**是完整的**  —— 位置、颜色、文字，观察里都有，直接搬
    层 1（在说什么）**是机械的**  —— 每个观察对象一个条目，**没有归并、没有判断**

    所以这个草稿**必然**是过细的：
    「有 44 个形状」和「图上有几个东西」不是一回事（项目栽过：11 个线条对象，
    10 条关系，两个数都对）。**归并、命名、判断关系，是下一步（agent）的活。**

为什么还要生成它：
    给 agent 一个起点，而不是让它对着 44 个对象的观察从零写。
    而且**全链路能在真材料上跑通**，Check B 才有原件可对。

刻意不做的两件事：
  1. **不把 PPT 的「组合」当成 spec 的 groups。**
     组合是**画法**（作者可能只是为了能一次性拖动），分组是**意义**。
     观察到的组合只记进 evidence，由 agent 判断是不是真有含义。
  2. **不给没有文字的形状填名字。** 按「原件故意留空」处理，text 留 null。

用法：
    python3 draft_spec.py work/observations.json --out work/spec.json
"""

import argparse
import json
import os
import sys

KIND_OF = {}          # 观察的 kind → 清单 layout 的 kind
ARROWISH = ("arrow",)


def kind_map(obs_shape):
    k = obs_shape.get("kind")
    if k == "picture":
        return "picture"
    if k == "text":
        return "text"
    if k == "brace":
        return "brace"
    if k == "arrow":
        return "arrow"
    # 形状：用预设几何判断圆角 / 菱形 / 椭圆
    prst = (obs_shape.get("prst") or "").lower()
    if "diamond" in prst:
        return "diamond"
    if "ellipse" in prst or "oval" in prst:
        return "ellipse"
    if "round" in prst:
        return "roundrect"
    return "rect"


def role_of(obs_shape):
    """**机械的**角色猜测 —— 依据是"文件里说它是什么"，不是"它是什么意思"。"""
    k = obs_shape.get("kind")
    if k == "text":
        return ["annotation"]
    if k in ARROWISH:
        return ["relation_carrier"]
    if k == "brace":
        return ["annotation"]
    return ["shape"]


def layout_of(obs, for_id, lid):
    """观察对象 → 清单的层 2 条目。"""
    kind = kind_map(obs)
    it = {"id": lid, "for": for_id, "kind": kind,
          "z": obs.get("z", 0)}
    if obs.get("prst"):
        it["prst"] = obs["prst"]
    if obs.get("rot"):
        it["rot"] = obs["rot"]

    # ---- 几何 ----
    if kind == "picture":
        it["box"] = obs["box"]
        if obs.get("src"):
            it["src"] = obs["src"]
    else:
        it["box"] = obs["box"]

    # ---- 颜色：如实反映"来自 XML"还是"从渲染图采的" ----
    prov = "original"
    fl = obs.get("fill") or {}
    if fl.get("color"):
        it["fill"] = fl["color"]
    if fl.get("src") == "sampled":
        prov = "measured"
    ln = obs.get("line") or {}
    # ★ 只有**真拿到颜色**才给描边。
    #   写成 `if ln.get("color") or ln.get("dash")` 是个坑：dash 永远是
    #   "solid" 这种非空字符串（恒为真），于是**每个形状都被加一圈线**，
    #   颜色未知时还默认成黑色 —— 实测重画出来每个框都镶了黑边。
    #   拿不到颜色就不要描边；src == "nofill" 是"明确没有描边"。
    if ln.get("src") != "nofill" and ln.get("color"):
        it["line"] = {"color": ln["color"],
                      "pt": ln.get("pt", 0.75),
                      "dash": ln.get("dash") or "solid"}
        if ln.get("alpha") is not None:
            it["line"]["alpha"] = ln["alpha"]
        if ln.get("src") == "sampled":
            prov = "measured"
    if kind == "arrow" and fl.get("color"):
        it["fill"] = fl["color"]
    if kind == "brace" and fl.get("color"):
        it["fill"] = None

    # ---- 文字 ----
    txt = obs.get("text") or ""
    if txt:
        fg = (obs.get("fg") or {})
        f = obs.get("font") or {}
        ins = obs.get("insets") or {}
        par = obs.get("para") or {}
        it["text"] = {
            "lines": txt.split("\n"),
            "font_pt": f.get("pt") or 10.0,
            "face": f.get("face") or "微软雅黑",
            "align": "l" if kind == "text" else "ctr",
            "color": fg.get("color") or "333333",
        }
        # ★ 排版属性必须传下去。漏了它们，文字会整体偏移、而且**换行位置会变**
        #   （内边距丢了 → 可用宽度变小 → "虚拟机" 被提前折成两行）。
        if ins:
            it["text"]["insets"] = [ins.get("l", 15.0), ins.get("t", 7.5),
                                    ins.get("r", 15.0), ins.get("b", 7.5)]
            it["text"]["anchor"] = ins.get("anchor") or "t"
            it["text"]["wrap"] = (ins.get("wrap") or "square") != "none"
        if par.get("lnspc") is not None:
            it["text"]["lnspc"] = par["lnspc"]
        if par.get("spc_before") is not None:
            it["text"]["spc_before"] = par["spc_before"]
        if fg.get("src") == "sampled":
            prov = "measured"

    it["provenance"] = prov
    it["evidence"] = list(obs.get("evidence") or [])
    return it


def line_item(c, for_id, lid):
    """连接符 → 清单的层 2（line）。"""
    it = layout_of(c, for_id, lid)
    it["kind"] = "line"
    it["p1"], it["p2"] = c["p1"], c["p2"]
    it.pop("box", None)
    it["arrow"] = c.get("arrow") or "end"
    if c.get("arrow_type"):
        it["arrow_type"] = c["arrow_type"]
    # ★ 保留原件连接符的预设几何。
    #   不保留的话，bentConnector2/3（折线）会被当成两点直线 —— 拐弯丢掉，斜穿。
    if c.get("prst"):
        it["prst"] = c["prst"]
    return it


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("observations")
    ap.add_argument("--out", default="spec.json")
    args = ap.parse_args()

    with open(args.observations, encoding="utf-8") as f:
        obs = json.load(f)

    shapes = obs.get("shapes") or []
    connectors = obs.get("connectors") or []
    src = obs.get("source") or {}

    entities, relations, layout = [], [], []
    by_sid = {s.get("shape_id"): s for s in shapes}

    # ---- 形状 → 实体 + 层2 ----
    for s in shapes:
        eid, lid = "e-%s" % s["oid"], "s-%s" % s["oid"]
        ev = list(s.get("evidence") or [])
        if s.get("group_path"):
            # 观察到的组合只记进 evidence —— 组合是画法，不是意义
            ev.append("组内: %s" % ",".join(s["group_path"]))
        entities.append({
            "id": eid,
            "role": role_of(s),
            "text": s.get("text") or None,
            "text_user": None,
            "box_src": s["box"],
            "provenance": "original",     # 形状、位置、文字都在文件里
            "evidence": ev,
            "status": "unconfirmed",
        })
        layout.append(layout_of(s, eid, lid))

    # ---- 连接符：两端都绑定 → 关系；否则 → 关系载体（还不知道连谁）----
    for c in connectors:
        ev = list(c.get("evidence") or [])
        st, en = c.get("st_cxn"), c.get("end_cxn")
        from_s, to_s = (st or {}).get("shape"), (en or {}).get("shape")
        from_e = "e-sp%d" % from_s if from_s else None
        to_e = "e-sp%d" % to_s if to_s else None
        known = from_e and to_e \
            and any(e["id"] == from_e for e in entities) \
            and any(e["id"] == to_e for e in entities)

        if known:
            rid, lid = "r-%s" % c["oid"], "s-%s" % c["oid"]
            relations.append({
                "id": rid,
                "from": from_e, "to": to_e,
                "kind": "association",
                "direction": "forward",
                "provenance": "measured",     # 端点是文件里明写的
                "evidence": ev + ["stCxn=%s endCxn=%s" % (from_s, to_s)],
                "status": "unconfirmed",
            })
            layout.append(line_item(c, rid, lid))
        else:
            # ⚠ 没有端点绑定 —— 关系只能**推**。这里不推，如实记成"关系载体"。
            eid, lid = "e-%s" % c["oid"], "s-%s" % c["oid"]
            why = []
            if not from_s and not to_s:
                why.append("两端都没绑定")
            else:
                why.append("只绑一端(st=%s end=%s)" % (from_s, to_s))
            entities.append({
                "id": eid,
                "role": ["relation_carrier"],
                "text": None,
                "text_user": None,
                "box_src": c["box"],
                "provenance": "original",
                "evidence": ev + why,
                "status": "unconfirmed",
            })
            layout.append(line_item(c, eid, lid))

    # ---- 组装 ----
    spec = {
        "spec_version": "0.2",
        "draft": {
            "by": "draft_spec.py",
            "note": "层 2 完整；层 1 是**机械草稿**——每个观察对象一个条目，"
                    "没有归并、没有判断。归并/命名/判断关系是 agent 的活。",
            "from": os.path.basename(args.observations),
        },
        "task": {"id": os.path.splitext(os.path.basename(args.out))[0],
                 "round": 0, "created": ""},
        "source": {
            "canvas": src.get("canvas"),
            "unit": src.get("unit") or "px@150dpi",
            "files": [{"path": os.path.basename(src.get("path") or ""),
                       "kind": "pptx_shapes", "slide": src.get("slide"),
                       "region": src.get("region")}],
        },
        "entities": entities,
        "relations": relations,
        "groups": [],
        "layout": layout,
        "changes": [],
        "open_questions": [],
        "findings": [],
        "unsupported": [
            {"what": u.get("what"), "name": u.get("name"),
             "reason": u.get("why"), "handled": "列入清单告知用户"}
            for u in obs.get("unsupported") or []
        ],
    }

    with open(args.out, "w", encoding="utf-8") as f:
        json.dump(spec, f, ensure_ascii=False, indent=1)

    print("草稿清单 -> %s" % args.out)
    print("  层 1：实体 %d（其中关系载体 %d）  关系 %d  分组 0"
          % (len(entities),
             len([e for e in entities if "relation_carrier" in e["role"]]),
             len(relations)))
    print("  层 2：%d 个对象" % len(layout))
    print("  画不出来的：%d" % len(spec["unsupported"]))
    print()
    print("  ⚠ 这是**草稿**。它必然过细——每个观察对象一个条目。")
    print("     归并成意义单位、判断关系、问用户，都还没做。")


if __name__ == "__main__":
    main()
