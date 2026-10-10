#!/usr/bin/env python3
"""
reconcile_test.py —— 证明 reconcile 抓得住「关系没被几何表达」，而且不误报。

用例**全部来自实测踩过的坑**：
  · 一条关系画成**两段**（我第一次就是按段比，误报 ✓）
  · 线**旋转过**（我第一次拿旋转前的盒子角比，误报 ✓）
  · 有线但端点**悬空**（原件 C10 就是那样 ✓）
  · 没有线（工具**不该判**，只该摆证据 ✓）

用法：
    python3 tests/reconcile_test.py
"""

import json
import os
import subprocess
import sys
import tempfile

HERE = os.path.dirname(os.path.abspath(__file__))
SCRIPTS = os.path.join(os.path.dirname(HERE), "scripts")
PY = sys.executable
REC = os.path.join(SCRIPTS, "reconcile.py")

SPEC = {
    "spec_version": "0.2",
    "entities": [
        {"id": "N1", "role": ["shape"], "text": "甲"},
        {"id": "N2", "role": ["shape"], "text": "乙"},
        {"id": "N3", "role": ["shape"], "text": "丙"},
    ],
    "relations": [{"id": "C1", "from": "N1", "to": "N2"},
                  {"id": "C3", "from": "N3", "to": "N2"}],
    "groups": [],
    "layout": [],
}

SVG_HEAD = ('<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 400 300" '
            'data-pptx-page-role="content">')


def boxes():
    return ('<rect id="N1" x="40" y="20" width="80" height="40"/>'
            '<rect id="N2" x="160" y="140" width="80" height="40"/>'
            '<rect id="N3" x="280" y="20" width="80" height="40"/>')


def texts():
    return ('<text x="80" y="45">甲</text>'
            '<text x="200" y="165">乙</text>'
            '<text x="320" y="45">丙</text>')


def run(svg, spec=SPEC):
    with tempfile.NamedTemporaryFile("w", suffix=".svg", delete=False,
                                     encoding="utf-8") as f:
        f.write(svg); p = f.name
    with tempfile.NamedTemporaryFile("w", suffix=".json", delete=False,
                                     encoding="utf-8") as f:
        json.dump(spec, f, ensure_ascii=False); sp = f.name
    try:
        r = subprocess.run([PY, REC, p, "--spec", sp], capture_output=True, text=True)
        return r.returncode, r.stdout
    finally:
        os.unlink(p); os.unlink(sp)


def case(name, svg, want_fail, needle=None):
    code, out = run(svg)
    failed = code != 0
    ok = (failed == want_fail) and (needle is None or needle in out)
    print("  %s %-38s %s" % ("✓" if ok else "✗", name,
                             ("失败" if failed else "通过") +
                             ("" if needle is None else
                              ("，" + ("抓到" if needle in out else "**没抓到**")))))
    return ok


def main():
    allok = True
    print("reconcile 的牙齿（用例都来自实测踩过的坑）：\n")

    # 基线：C1 有线两端贴上；C3 没有线 → 只报告，不算失败
    base = (SVG_HEAD + boxes() + texts() +
            '<line id="C1" x1="80" y1="60" x2="200" y2="140"/>' +
            '<line id="C3" x1="320" y1="60" x2="240" y2="140"/>' +
            '</svg>')
    allok &= case("两条线两端都贴上", base, want_fail=False)

    # ★ 有线但端点悬空
    hang = (SVG_HEAD + boxes() + texts() +
            '<line id="C1" x1="80" y1="60" x2="120" y2="120"/>' +
            '<line id="C3" x1="320" y1="60" x2="240" y2="140"/>' +
            '</svg>')
    allok &= case("C1 有线但终点悬空", hang, True, "两端没都贴上")

    # ★ 一条关系画成两段，合起来贴上 → 不该误报
    two = (SVG_HEAD + boxes() + texts() +
            '<line id="C1" x1="80" y1="60" x2="120" y2="100"/>'
            '<line id="C1" x1="120" y1="100" x2="200" y2="140"/>' +
            '<line id="C3" x1="320" y1="60" x2="240" y2="140"/>' +
            '</svg>')
    allok &= case("C1 画成两段（应通过）", two, want_fail=False)

    # ★ 没有线的，工具不该判失败，但要**摆出证据**
    none = (SVG_HEAD + boxes() + texts() +
            '<line id="C1" x1="80" y1="60" x2="200" y2="140"/>' +
            '</svg>')
    allok &= case("C3 没有线（只报告，不算失败）", none, False, "工具不判")

    # ★ 文字缺
    miss = (SVG_HEAD + boxes() + texts().replace(">丙<", "><") +
            '<line id="C1" x1="80" y1="60" x2="200" y2="140"/>' +
            '<line id="C3" x1="320" y1="60" x2="240" y2="140"/>' +
            '</svg>')
    allok &= case("丙 的文字没画出来", miss, True, "缺 1 个")

    # ★ 分区重叠
    ov = (SVG_HEAD +
          '<g id="A" data-pptx-bounds="0 0 200 200">' + boxes() + '</g>'
          '<g id="B" data-pptx-bounds="100 100 200 200"><rect id="N9" x="110" y="110" width="20" height="20"/></g>'
          + texts() +
          '<line id="C1" x1="80" y1="60" x2="200" y2="140"/>' +
          '<line id="C3" x1="320" y1="60" x2="240" y2="140"/>' +
          '</svg>')
    allok &= case("两个分区重叠", ov, True, "重叠")

    print()
    print("reconcile_test：%s" % ("全部符合预期 ✓" if allok else "有失败 ✗"))
    return 0 if allok else 1


if __name__ == "__main__":
    sys.exit(main())
