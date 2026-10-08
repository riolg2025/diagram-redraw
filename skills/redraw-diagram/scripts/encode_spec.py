#!/usr/bin/env python3
"""
encode_spec.py —— 把「我读到的」+「观察」编码成最终的清单（层 1 + 层 2）。

分工：
    观察 observations.json  —— 机械清点：原件里有什么（位置、文字、颜色、XML）
    解读 reading.json       —— 我的判断：图上**有什么意义**，各自对应原件里哪些对象
    → 这份清单             —— 层 1 用解读，层 2 用观察，两边靠 oid 对上

★ 它顺手做一件很重要的事：**覆盖检查**。
  「原件里在这个范围内的对象，我是不是每一个都认领了？」
  没被任何解读条目认领的对象，会明确列出来 —— 那可能就是**我漏读了**。
  这就是「读」这一步的 Check B。

用法：
    python3 encode_spec.py work/reading.json work/observations.json --out work/spec.json
"""

import argparse
import json
import os
import sys

KIND_TO_ROLE_FALLBACK = "shape"


def load(p):
    with open(p, encoding="utf-8") as f:
        return json.load(f)


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("reading")
    ap.add_argument("observations")
    ap.add_argument("--out", required=True)
    ap.add_argument("--task-id", default=None)
    args = ap.parse_args()

    R = load(args.reading)
    O = load(args.observations)

    items = {it["id"]: it for it in R["items"]}
    scope = R.get("scope") or {}
    s1 = items["S1"]["box"]
    ox, oy = s1[0], s1[1]
    canvas = [round(s1[2]), round(s1[3])]

    # ---- 谁认领了哪些原件对象 ----------------------------------------
    claim = {}
    for it in R["items"]:
        for oid in it.get("oids") or []:
            if oid in claim:
                sys.exit("✗ %s 被 %s 和 %s 同时认领了" % (oid, claim[oid], it["id"]))
            claim[oid] = it["id"]

    # ---- 原件里、落在范围内、但没人认领的对象（= 我可能漏读了）--------
    def in_scope(b):
        return not (b[0] + b[2] < ox or b[0] > ox + canvas[0]
                    or b[1] + b[3] < oy or b[1] > oy + canvas[1])

    all_obs = (O.get("shapes") or []) + (O.get("connectors") or [])
    orphan = [o for o in all_obs
              if o["oid"] not in claim and in_scope(o["box"])]
    # 范围外的：被 S2/S3 认领的，或者压根没在范围内的
    excluded = [o for o in all_obs
                if claim.get(o["oid"]) in ("S2", "S3")
                or (o["oid"] not in claim and not in_scope(o["box"]))]

    # ---- 层 1 ---------------------------------------------------------
    entities, relations, groups = [], [], []
    obs_by_oid = {o["oid"]: o for o in all_obs}

    for it in R["items"]:
        iid = it["id"]
        if iid in ("S1", "S2", "S3") or it.get("part_of"):
            continue
        role = it["role"]
        oids = it.get("oids") or []
        ev = []
        for oid in oids:
            ev += (obs_by_oid.get(oid, {}).get("evidence") or [])

        if role == "line":
            # ★ 关系的两端**必须**由解读显式给出 —— 不许代码从几何猜。
            #   实测猜过一次，10 条里错了 5 条（C5 猜成 N4→N8，C10 猜成 T4→T4）。
            #   那正是我们说过不能做的事：**机械猜意义**。
            frm, to = it.get("from"), it.get("to")
            if not frm or not to:
                sys.exit("✗ %s 是关系，但解读里没写 from / to。"
                         "关系连的是谁，是**读出来的判断**，不许代码猜。" % iid)
            relations.append({
                "id": iid,
                "from": frm or "?",
                "to": to or "?",
                "kind": it.get("rel_kind") or "sequence",
                "direction": it.get("direction") or "forward",
                "provenance": it.get("provenance") or "inferred",
                "evidence": ev + [it.get("says") or ""],
                "status": "confirmed" if it.get("confirmed") else "unconfirmed",
            })
            continue

        if role == "container":
            g = {"id": iid, "label": it.get("says") or "",
                 "members": it.get("members") or [],
                 "provenance": "original", "evidence": ev}
            groups.append(g)
            continue

        # 实体
        text = None
        for oid in oids:
            t = obs_by_oid.get(oid, {}).get("text")
            if t:
                text = t
                break
        e = {
            "id": iid,
            "role": [{"node": "shape", "label": "annotation",
                      "unsupported": "decoration"}.get(role, "shape")],
            "text": text,
            "text_user": it.get("text_user"),
            "box_src": it.get("box"),
            "provenance": "original" if text else "measured",
            "evidence": ev,
            "status": "changed" if it.get("text_user") else
                      ("confirmed" if it.get("confirmed") else "unconfirmed"),
        }
        if it.get("group"):
            e["group"] = it["group"]
        entities.append(e)

    # ---- 层 2：观察对象 → layout，坐标平移到新画布 ---------------------
    layout = []
    for o in all_obs:
        oid = o["oid"]
        owner = claim.get(oid)
        # S1/S2/S3 是**范围标记**，不是层 1 的条目，不能当 for 的目标
        if owner is None or owner in ("S1", "S2", "S3"):
            continue                      # 范围外，不画（但下面会如实列出来）
        if items[owner].get("part_of"):
            owner = items[owner]["part_of"]
        b = o["box"]
        nb = [round(b[0] - ox, 1), round(b[1] - oy, 1), round(b[2], 1), round(b[3], 1)]
        lid = "s-%s" % oid

        if o.get("kind") == "connector":
            it = {"id": lid, "for": owner, "kind": "line",
                  "p1": [round(o["p1"][0] - ox, 1), round(o["p1"][1] - oy, 1)],
                  "p2": [round(o["p2"][0] - ox, 1), round(o["p2"][1] - oy, 1)],
                  "arrow": o.get("line", {}).get("arrow") or "end",
                  "provenance": "measured", "evidence": o.get("evidence") or []}
            if o.get("prst"):
                it["prst"] = o["prst"]
            if o["line"].get("arrow_type"):
                it["arrow_type"] = o["line"]["arrow_type"]
            if o["line"].get("color"):
                it["line"] = {"color": o["line"]["color"],
                              "pt": o["line"].get("pt", 0.75),
                              "dash": o["line"].get("dash") or "solid"}
                if o["line"].get("alpha") is not None:
                    it["line"]["alpha"] = o["line"]["alpha"]
            layout.append(it)
            continue

        if o.get("kind") == "verbatim":
            layout.append({"id": lid, "for": owner, "kind": "verbatim",
                           "box": nb, "src": o["xml"],
                           "provenance": "measured",
                           "evidence": o.get("evidence") or []})
            continue

        if o.get("kind") == "picture":
            it = {"id": lid, "for": owner, "kind": "picture", "box": nb,
                  "provenance": "measured", "evidence": o.get("evidence") or []}
            if o.get("src"):
                it["src"] = o["src"]
            layout.append(it)
            continue

        kind = {"text": "text", "brace": "brace", "arrow": "arrow"}.get(
            o.get("kind"), "rect")
        if kind == "rect":
            prst = (o.get("prst") or "").lower()
            kind = ("diamond" if "diamond" in prst else
                    "ellipse" if ("ellipse" in prst or "oval" in prst) else
                    "roundrect" if "round" in prst else "rect")
        it = {"id": lid, "for": owner, "kind": kind, "box": nb,
              "rot": o.get("rot", 0), "z": o.get("z", 0),
              "provenance": "measured", "evidence": o.get("evidence") or []}
        if o.get("prst"):
            it["prst"] = o["prst"]
        prov = "measured"
        fl = o.get("fill") or {}
        if fl.get("color"):
            it["fill"] = fl["color"]
            if fl.get("src") == "xml":
                prov = "original"
        ln = o.get("line") or {}
        if ln.get("src") != "nofill" and ln.get("color"):
            it["line"] = {"color": ln["color"], "pt": ln.get("pt", 0.75),
                          "dash": ln.get("dash") or "solid"}
            if ln.get("src") == "sampled":
                prov = "measured"
        txt = o.get("text") or ""
        if txt:
            fg = o.get("fg") or {}
            f = o.get("font") or {}
            e = next((x for x in entities if x["id"] == owner), None)
            user_text = (e or {}).get("text_user")
            shown = user_text or txt
            it["text"] = {"lines": shown.split("\n"),
                          "font_pt": f.get("pt") or 10.0,
                          "face": f.get("face") or "微软雅黑",
                          "align": "l" if o.get("kind") == "text" else "ctr",
                          "color": fg.get("color") or "333333"}
            if user_text:
                # ★ 这段文字是**用户改过的**，不是原件写的 —— 必须标出来，
                #   否则 Check B 会把它当成"我编了一个原件里没有的字"。
                it["text"]["prov"] = "user"
            if fg.get("src") == "sampled":
                prov = "measured"
            ins = o.get("insets") or {}
            if ins:
                it["text"]["insets"] = [ins.get("l", 15.0), ins.get("t", 7.5),
                                        ins.get("r", 15.0), ins.get("b", 7.5)]
                it["text"]["anchor"] = ins.get("anchor") or "t"
                it["text"]["wrap"] = (ins.get("wrap") or "square") != "none"
            par = o.get("para") or {}
            if par.get("lnspc") is not None:
                it["text"]["lnspc"] = par["lnspc"]
        it["provenance"] = prov
        layout.append(it)

    # ---- 台账：用户改过的，单独记一条 --------------------------------
    changes = []
    t7 = items.get("T7") or {}
    if t7.get("text_user"):
        changes.append({
            "at": "", "round": 1, "target": "T7.text",
            "from": "6.部署到开发环境，制品晋级",
            "to": t7["text_user"],
            "by": "user",
            "reason": "原件写「开发环境」却站在测试环境上方，用户确认改为「测试环境」",
            "evidence_ref": "u004",
        })

    spec = {
        "spec_version": "0.2",
        "task": {"id": args.task_id or os.path.splitext(os.path.basename(args.out))[0],
                 "round": 1, "created": ""},
        "source": {
            "canvas": canvas, "unit": "px@150dpi",
            "files": [{"path": os.path.basename(O["source"].get("path") or ""),
                       "kind": "pptx_shapes", "slide": O["source"].get("slide"),
                       "region": s1}],
        },
        "entities": entities, "relations": relations, "groups": groups,
        "layout": layout, "changes": changes,
        "open_questions": [], "findings": [],
        "unsupported": [
            {"what": "不在重画范围（用户确认只重画 S1）",
             "name": o["oid"], "reason": "范围外", "handled": "不画"}
            for o in excluded
        ],
    }

    with open(args.out, "w", encoding="utf-8") as f:
        json.dump(spec, f, ensure_ascii=False, indent=1)

    # ---- 报告 ---------------------------------------------------------
    print("清单 -> %s" % args.out)
    print("  画布 %.0fx%.0f（原来那张图的左上角平移到 0,0）" % (canvas[0], canvas[1]))
    print("  层 1：实体 %d  关系 %d  分组 %d"
          % (len(entities), len(relations), len(groups)))
    print("  层 2：%d 个对象" % len(layout))
    print("  范围外被排除的对象：%d" % len(excluded))
    print()
    if orphan:
        print("  ⚠ **范围内有 %d 个原件对象没被任何条目认领** —— 可能是我漏读了：" % len(orphan))
        for o in orphan[:12]:
            b = o["box"]
            print("     %-8s %-12s (%.0f,%.0f %.0fx%.0f) %r"
                  % (o["oid"], o.get("kind"), b[0], b[1], b[2], b[3],
                     (o.get("text") or "")[:20]))
        print("     这是「读」这一步的 Check B：原件里的东西，我是不是都读到了。")
    else:
        print("  ✓ 范围内的原件对象**全部被认领**（没漏读）")


if __name__ == "__main__":
    main()
