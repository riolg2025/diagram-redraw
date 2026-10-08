#!/usr/bin/env python3
"""
mutation_test.py —— 证明 spec_check.py **有牙齿**。

只看"有效清单通过"是不够的：一个永远返回"通过"的校验器也能做到。
所以对**每一条不变式**各造一个变异，看校验器抓不抓得到。

用法：
    python3 tests/mutation_test.py            # 全部跑一遍
    python3 tests/mutation_test.py -v         # 打印每个变异报了什么
"""

import copy
import json
import os
import sys
import argparse

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, os.path.join(os.path.dirname(HERE), "scripts"))
import spec_check  # noqa: E402

FIXTURE = os.path.join(HERE, "fixture-min.json")


# --------------------------------------------------------------------------
# 每个变异 = (要证的不变式, 破坏它的函数)
# --------------------------------------------------------------------------
def _mut(name, fn):
    return (name, fn)


MUTATIONS = [
    _mut("① spec_version 不对", lambda s: s.__setitem__("spec_version", "9.9")),
    _mut("① 缺顶层字段 entities", lambda s: s.pop("entities")),
    _mut("① entities 不是数组", lambda s: s.__setitem__("entities", {})),
    _mut("① source 不是对象", lambda s: s.__setitem__("source", [])),
    _mut("① canvas 不是 [宽,高]", lambda s: s["source"].__setitem__("canvas", [2000])),
    _mut("① changes 不是数组", lambda s: s.__setitem__("changes", {})),

    _mut("② id 在表内重复",
         lambda s: s["entities"][1].__setitem__("id", "e01")),
    _mut("② id 跨表撞车",
         lambda s: s["layout"][0].__setitem__("id", "e01")),
    _mut("② layout 缺 id",
         lambda s: s["layout"][0].pop("id")),

    _mut("③ from 指向不存在",
         lambda s: s["relations"][0].__setitem__("from", "e99")),
    _mut("③ to 指向不存在",
         lambda s: s["relations"][0].__setitem__("to", "e99")),
    _mut("③ relation 缺 from",
         lambda s: s["relations"][0].pop("from")),
    _mut("⑧ relation.kind 非法",
         lambda s: s["relations"][0].__setitem__("kind", "banana")),
    _mut("⑧ relation.direction 非法",
         lambda s: s["relations"][0].__setitem__("direction", "sideways")),

    _mut("④ groups.members 指向不存在",
         lambda s: s["groups"][0]["members"].append("e99")),

    _mut("⑤ layout 缺 for",
         lambda s: s["layout"][0].pop("for")),
    _mut("⑤ layout.for 指向不存在",
         lambda s: s["layout"][0].__setitem__("for", "e99")),
    _mut("⑧ layout.kind 非法",
         lambda s: s["layout"][0].__setitem__("kind", "hexagon")),

    _mut("⑥ 实体没被任何对象表达",
         lambda s: s.__setitem__("layout",
                                 [x for x in s["layout"] if x.get("for") != "e03"])),
    _mut("⑥ 关系没被任何对象表达",
         lambda s: s.__setitem__("layout",
                                 [x for x in s["layout"] if x.get("for") != "r02"])),
    _mut("⑥ 分组没被任何对象表达",
         lambda s: s.__setitem__("layout",
                                 [x for x in s["layout"] if x.get("for") != "g01"])),
    _mut("⑥ 标了 not_drawn 却画了",
         lambda s: s["layout"].append({
             "id": "s99", "for": "e04", "kind": "rect",
             "box": [0.0, 0.0, 10.0, 10.0],
             "provenance": "measured", "evidence": ["p:sp#40"]})),

    _mut("⑦ provenance 无证据",
         lambda s: s["entities"][0].__setitem__("evidence", [])),
    _mut("⑦ evidence 只有空白",
         lambda s: s["entities"][0].__setitem__("evidence", ["  "])),
    _mut("⑦ 用户说的却没 evidence_ref（changes）",
         lambda s: s["changes"][0].pop("evidence_ref")),

    _mut("⑧ provenance 非法",
         lambda s: s["entities"][0].__setitem__("provenance", "guessed")),
    _mut("⑧ status 非法",
         lambda s: s["entities"][0].__setitem__("status", "maybe")),
    _mut("⑧ role 非法",
         lambda s: s["entities"][0].__setitem__("role", ["blob"])),
    _mut("⑧ changes.by 非法",
         lambda s: s["changes"][0].__setitem__("by", "robot")),

    _mut("⑩ verbatim 没给 src",
         lambda s: s["layout"].append({
             "id": "s98", "for": "e01", "kind": "verbatim",
             "box": [0.0, 0.0, 10.0, 10.0],
             "provenance": "measured", "evidence": ["p:sp#70"]})),

    _mut("⑩ line 用了 box",
         lambda s: s["layout"][3].__setitem__("box", [0, 0, 10, 10])),
    _mut("⑩ line 缺 p1/p2",
         lambda s: s["layout"][3].pop("p2")),
    _mut("⑩ 非 line 用了 p1/p2",
         lambda s: s["layout"][0].__setitem__("p1", [0, 0])),
    _mut("⑩ 非 line 缺 box",
         lambda s: s["layout"][0].pop("box")),
    _mut("⑩ box 不是四元组",
         lambda s: s["layout"][0].__setitem__("box", [1, 2, 3])),
]


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("-v", "--verbose", action="store_true")
    args = ap.parse_args()

    base = json.load(open(FIXTURE, encoding="utf-8"))

    # 基线必须先通过，否则后面全是假象
    probs = spec_check.check(copy.deepcopy(base))
    if probs:
        print("✗ 基线就不通过，变异测试没意义：")
        for w, m in probs.items:
            print("   %-28s %s" % (w, m))
        return 1
    print("基线：有效清单通过 ✓")

    caught = 0
    escaped = []
    for name, fn in MUTATIONS:
        s = copy.deepcopy(base)
        try:
            fn(s)
        except Exception as e:                       # 变异本身写错了
            escaped.append((name, "变异执行失败：%r" % e))
            continue
        probs = spec_check.check(s)
        if probs:
            caught += 1
            if args.verbose:
                print("  ✓ %-34s → %d 处：%s"
                      % (name, len(probs), probs.items[0][1][:60]))
        else:
            escaped.append((name, "**校验器没抓到**"))

    total = len(MUTATIONS)
    print()
    print("变异测试：%d/%d 被抓到" % (caught, total))
    if escaped:
        print()
        print("✗ 漏掉的（这些不变式名存实亡）：")
        for name, why in escaped:
            print("   %-34s %s" % (name, why))
        return 1
    print("✓ 全部抓到——每一条不变式都有牙齿")
    return 0


if __name__ == "__main__":
    sys.exit(main())
