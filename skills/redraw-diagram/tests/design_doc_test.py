#!/usr/bin/env python3
"""
design_doc_test.py —— 证明 check_design_doc 的四条规矩**都有牙齿**。

每条规矩造一个"刚好违规"的稿子，看它抓不抓得住；
再造一个"刚好合规"的，看它会不会误报。

用法：
    python3 tests/design_doc_test.py
"""

import json
import os
import subprocess
import sys
import tempfile

HERE = os.path.dirname(os.path.abspath(__file__))
SCRIPTS = os.path.join(os.path.dirname(HERE), "scripts")
PY = sys.executable
CHECK = os.path.join(SCRIPTS, "check_design_doc.py")

OK = """# ② 设计稿

## 一、拓扑判断
**topology = `yes`**
**理由**：这是一张流程图，关系就是内容。

## 二、Relationships（来源明说的）
| 关系 | 从 | 到 | 类型 |
|---|---|---|---|
| C1 | 甲 | 乙 | `order` |
| C3 | 丙 | 乙 | `link` |

## 三、每条关系用什么表达
| 关系 | 类型 | 表达方式 | 依据 |
|---|---|---|---|
| C1 | `order` | **线** | `order` → 一条读路径 |
| C3 | `link` | **位置**：丙贴在乙正左侧 | **说得清**：正左、无遮挡、中间没有别的东西 |

## 四、组合倾向
```
上：流程脊柱
下：环境
```
**分区**：flow · band

## 五、角色锚点
| 角色 | 用途 | 字号 | 填色 | 文字色 |
|---|---|---|---|---|
| 节点 | 主干 | 20 | `#0084CF` | 白 |
"""

SPEC = {
    "spec_version": "0.2",
    "relations": [{"id": "C1", "from": "N1", "to": "N2"},
                  {"id": "C3", "from": "N3", "to": "N2"}],
}


def run(doc, spec_path):
    with tempfile.NamedTemporaryFile("w", suffix=".md", delete=False,
                                     encoding="utf-8") as f:
        f.write(doc)
        p = f.name
    try:
        r = subprocess.run([PY, CHECK, p, "--spec", spec_path],
                           capture_output=True, text=True)
        return r.returncode, r.stdout
    finally:
        os.unlink(p)


def case(name, doc, spec, want_fail, needle=None):
    if spec:
        with tempfile.NamedTemporaryFile("w", suffix=".json", delete=False,
                                         encoding="utf-8") as f:
            json.dump(spec, f, ensure_ascii=False)
            sp = f.name
    else:
        sp = os.devnull
    try:
        code, out = run(doc, sp)
    finally:
        if spec:
            os.unlink(sp)
    failed = code != 0
    ok = (failed == want_fail) and (needle is None or needle in out)
    print("  %s %-34s %s" % ("✓" if ok else "✗", name,
                             ("失败" if failed else "通过") +
                             ("" if needle is None else
                              ("，" + ("抓到" if needle in out else "**没抓到**")))))
    return ok


def main():
    allok = True
    print("check_design_doc 的牙齿：\n")

    # 合规的不能被误杀
    allok &= case("合规稿", OK, SPEC, want_fail=False)

    # ① 缺节
    allok &= case("缺「五、角色锚点」", OK.replace("## 五、角色锚点",
                                                "## 别的"), SPEC, True, "缺")

    # ② 坐标漏进来
    allok &= case("第三节混进 265px",
                  OK.replace("无遮挡", "离它 265px"), SPEC, True, "带单位的值")
    allok &= case("第三节混进坐标对",
                  OK.replace("正左、无遮挡", "在 (375,177) 处"), SPEC, True, "坐标对")
    # 五节的锚点**允许**有数值
    allok &= case("第五节有数值（应允许）", OK, SPEC, want_fail=False)

    # ③ 关系漏表达
    allok &= case("漏了 C3 的表达决定",
                  OK.replace("| C3 | `link` | **位置**：丙贴在乙正左侧 | **说得清**：正左、无遮挡、中间没有别的东西 |\n", ""),
                  SPEC, True, "没写")

    # ④ 用「位置」没写理由
    allok &= case("「位置」没写理由",
                  OK.replace("**说得清**：正左、无遮挡、中间没有别的东西", "好"),
                  SPEC, True, "没写")

    # ④ 引用链要能解析
    ref = OK.replace("| C3 | `link` | **位置**：丙贴在乙正左侧 | **说得清**：正左、无遮挡、中间没有别的东西 |",
                     "| C3 | `link` | **位置**：丙贴在乙正左侧 | 同 C1 |")
    allok &= case("「位置」引用另一条（应允许）", ref, SPEC, want_fail=False)
    # 但引用到一条根本没有理由的，要抓
    ref2 = OK.replace("`order` → 一条读路径", "好").replace(
        "**说得清**：正左、无遮挡、中间没有别的东西", "同 C1")
    allok &= case("引用链指向没理由的（应抓）", ref2, SPEC, True, "引用不到")

    print()
    print("design_doc_test：%s" % ("全部符合预期 ✓" if allok else "有失败 ✗"))
    return 0 if allok else 1


if __name__ == "__main__":
    sys.exit(main())
