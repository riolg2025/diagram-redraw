#!/usr/bin/env python3
"""
render_pptx.py —— 把 PPT/PPTX 渲染成每一页的 PNG，用来"看原件"和"验重画结果"。

★ 为什么不能直接用 LibreOffice 渲染 pptx：
  实测这份 PPT **直接渲染 6 页全部是空白**（6 页的字节数一模一样）。
  必须先转 PDF、再从 PDF 取图（走 pdfium）才正常。
  这里把那条弯路固化下来了。

需要 DSH 自带的 LibreOffice kit（Node 调用）。路径可用环境变量覆盖：
  DSH_NODE   DSH_LO_CLI

用法：
  python3 render_pptx.py deck.pptx --pages 6 --outdir out --dpi 150
"""

import argparse
import json
import os
import shutil
import subprocess
import sys
import tempfile

DEFAULT_NODE = ("/Applications/DeepSeek Harness.app/Contents/Resources/runtime/"
                "primary-runtime/dependencies/node/bin/node")
DEFAULT_CLI = ("/Applications/DeepSeek Harness.app/Contents/Resources/app.asar.unpacked/"
               "dsh/node_modules/@deepseek-ai/libreoffice-kit/lib/cli.js")


def run(node, cli, *args):
    cmd = [node, cli] + list(args)
    r = subprocess.run(cmd, capture_output=True, text=True)
    if r.returncode != 0:
        raise RuntimeError(f"{' '.join(cmd[:3])} ... 失败：\n{r.stderr[-800:]}")
    return r.stdout.strip()


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("pptx")
    ap.add_argument("--pages", default=None, help="如 1,4,5；不给则全部")
    ap.add_argument("--outdir", required=True)
    ap.add_argument("--dpi", type=int, default=150)
    args = ap.parse_args()

    node = os.environ.get("DSH_NODE", DEFAULT_NODE)
    cli = os.environ.get("DSH_LO_CLI", DEFAULT_CLI)
    if not os.path.exists(node) or not os.path.exists(cli):
        sys.exit("找不到 LibreOffice kit。用 DSH_NODE / DSH_LO_CLI 指定。")

    src = os.path.abspath(args.pptx)
    os.makedirs(args.outdir, exist_ok=True)
    tmp = tempfile.mkdtemp(prefix="render-")
    try:
        pdf = os.path.join(tmp, "deck.pdf")
        # 第一步：转 PDF（直接渲染 pptx 会得到空白页）
        run(node, cli, "convert", "--input", src, "--output", pdf)
        if not os.path.exists(pdf) or os.path.getsize(pdf) < 1000:
            sys.exit("转 PDF 失败或产物过小。")

        # 第二步：从 PDF 取图
        pagedir = os.path.join(tmp, "pages")
        cmd = ["render", "--input", pdf, "--output-dir", pagedir, "--dpi", str(args.dpi)]
        if args.pages:
            cmd += ["--pages", args.pages]
        out = run(node, cli, *cmd)
        info = json.loads(out)

        want = set(int(v) for v in args.pages.split(",")) if args.pages else None
        made = []
        for im in info["images"]:
            page = im["page"]
            if want and page not in want:
                continue
            dst = os.path.join(args.outdir, f"page-{page:04d}.png")
            shutil.copyfile(im["path"], dst)
            made.append(dst)
            print(f"  第 {page} 页 -> {dst}  ({im['width']}x{im['height']})")
        if info.get("missingFonts"):
            print("  ⚠ 缺字体:", ", ".join(info["missingFonts"]))
            print("     缺字体会**改变文字排版**——"
                  "你看到的位置可能和原作者看到的不一样。")
        if not made:
            sys.exit("没有产出任何页。检查 --pages。")
    finally:
        shutil.rmtree(tmp, ignore_errors=True)


if __name__ == "__main__":
    main()
