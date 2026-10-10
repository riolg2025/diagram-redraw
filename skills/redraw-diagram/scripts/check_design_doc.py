#!/usr/bin/env python3
"""
check_design_doc.py —— 查②设计稿（design-format.md）的四条硬规则。

★ 它**查不了"这个决定好不好"** ✗ —— 那是审美，只有人能判。
  它只查"该有的有没有、不该有的有没有混进来"。

四条：
  ① 五节齐不齐
  ② **有没有坐标/尺寸漏进来**（除「五、角色锚点」外）
  ③ 清单里的关系，第三节漏了谁
  ④ 写了「位置」的，有没有写"为什么说得清"

用法：
    python3 check_design_doc.py work/design.md --spec work/spec.json
"""

import argparse
import json
import os
import re
import sys

# 五节的标题（用关键词匹配，不要求一字不差）
SECTIONS = [("一", ("拓扑判断", "topology")),
            ("二", ("Relationships", "关系")),
            ("三", ("用什么表达", "表达方式")),
            ("四", ("组合倾向", "Composition")),
            ("五", ("角色锚点", "锚点"))]

# ② 里绝对不许出现的东西
UNIT = re.compile(r"\d+(?:\.\d+)?\s*(?:px|pt|em|rem|mm|cm|in|%)", re.I)
PAIR = re.compile(r"\(\s*-?\d+(?:\.\d+)?\s*,\s*-?\d+(?:\.\d+)?\s*\)")
# 允许出现的（id、节号之类）
ALLOW = re.compile(r"\b(?:C|N|T|G|P)\d+\b")


def split_sections(text):
    """按 ## 标题切节。返回 [(标题, 正文)]。"""
    out, cur, buf = [], None, []
    for line in text.splitlines():
        if line.startswith("## "):
            if cur is not None:
                out.append((cur, "\n".join(buf)))
            cur, buf = line[3:].strip(), []
        else:
            buf.append(line)
    if cur is not None:
        out.append((cur, "\n".join(buf)))
    return out


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("design")
    ap.add_argument("--spec", help="清单（用来核对关系有没有漏表达）")
    a = ap.parse_args()

    text = open(a.design, encoding="utf-8").read()
    secs = split_sections(text)
    bad = 0

    print("② 设计稿检查：%s" % os.path.basename(a.design))
    print()

    # ① 五节齐不齐
    titles = [t for t, _ in secs]
    print("① 五节")
    found = {}
    for num, keys in SECTIONS:
        hit = next((t for t in titles if any(k in t for k in keys)), None)
        found[num] = hit
        print("   %s、%-10s %s" % (num, "/".join(keys[:2]), "✓ " + hit if hit else "✗ **缺**"))
        if not hit:
            bad += 1
    print()

    # ② 坐标/尺寸漏进来没有（第五节豁免）
    print("② 有没有坐标/尺寸漏进来（「五、角色锚点」豁免）")
    leaked = 0
    for t, body in secs:
        if found.get("五") and t == found["五"]:
            continue
        for i, line in enumerate(body.splitlines(), 1):
            if line.strip().startswith((">", "|--", "```")):
                continue          # 引用/表头/代码块里的说明不算
            clean = ALLOW.sub("", line)
            for pat, what in ((UNIT, "带单位的值"), (PAIR, "坐标对")):
                for m in pat.finditer(clean):
                    leaked += 1
                    print("   ✗ [%s] %s：%r" % (t[:6], what, m.group(0)))
    if not leaked:
        print("   ✓ 没有")
    else:
        print("   → %d 处。**量出来的数字不属于这份文件** ✗ 它们是「对输出的检查」，")
        print("     放到单独的检查记录里，或者放进第五节当锚点。")
        bad += leaked
    print()

    # ③ 关系有没有漏表达
    if a.spec:
        print("③ 每条关系有没有表达决定")
        S = json.load(open(a.spec, encoding="utf-8"))
        rels = [r["id"] for r in S.get("relations") or []]
        sec3 = next((b for t, b in secs if found.get("三") and t == found["三"]), "")
        miss = [r for r in rels if r not in sec3]
        print("   清单里有 %d 条关系" % len(rels))
        if miss:
            print("   ✗ 第三节里没写：%s" % "、".join(miss))
            bad += len(miss)
        else:
            print("   ✓ 每条都有")
        print()
    else:
        sec3 = ""

    # ④ 用「位置」的有没有写理由
    print("④ 写了「位置」的，有没有写「为什么说得清」")
    sec3 = sec3 or next((b for t, b in secs if found.get("三") and t == found["三"]), "")
    rows = {}
    for line in sec3.splitlines():
        if "|" not in line:
            continue
        cells = [c.strip() for c in line.strip().strip("|").split("|")]
        if len(cells) >= 4 and re.match(r"^[CNTG]\d+$", cells[0]):
            rows[cells[0]] = {"how": cells[2], "why": cells[3]}

    def reason_of(rid, seen=None):
        """理由可以是自己的，也可以是**引用另一条**（引用链要能解析）。"""
        seen = seen or set()
        if rid in seen or rid not in rows:
            return None
        seen.add(rid)
        why = rows[rid]["why"]
        ref = re.search(r"同\s*([CNTG]\d+)", why)
        if ref and len(why) < 12:          # 纯引用 → 顺着链找
            return reason_of(ref.group(1), seen)
        return why if len(why) >= 8 else None

    n_pos = n_ref = n_bad = 0
    for rel, r in rows.items():
        if "位置" not in r["how"]:
            continue
        n_pos += 1
        why = reason_of(rel)
        if why is None:
            print("   ✗ %s 用了「位置」但没写（也引用不到）为什么说得清" % rel)
            n_bad += 1
        elif re.search(r"同\s*[CNTG]\d+", r["why"]) and len(r["why"]) < 12:
            n_ref += 1
    bad += n_bad
    if n_pos and not n_bad:
        print("   ✓ %d 条用「位置」的都有理由%s" % (n_pos, "（其中 %d 条是引用别条）" % n_ref if n_ref else ""))
    if n_pos == 0:
        print("   （没有用「位置」的）")
    print()

    print("=" * 60)
    if bad:
        print("✗ %d 处不合格式" % bad)
    else:
        print("✓ 格式合规")
        print("  ⚠️ 注意：这只说明「格式对」，**不说明「决定好」** ——")
        print("     那些决定好不好，只有人（或另一双眼睛）能判。")
    return 0 if bad == 0 else 1


if __name__ == "__main__":
    sys.exit(main())
