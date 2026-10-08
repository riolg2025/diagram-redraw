#!/usr/bin/env python3
"""
svg2png.py —— 把 SVG 渲染成 PNG，**用于自检**。

★ 为什么必须有这个：
  这套工具链的 LibreOffice **不支持 svg -> 图**，所以 SVG 一度是
  "我交出去但自己看不见"的东西，连着错了两次（箭头没头、大括号画错）。
  没有这个脚本，SVG 就只能靠用户当眼睛。

本机探测结果（2026-10）：
  rsvg-convert / inkscape / cairosvg  —— 没有
  qlmanage                             —— 被沙箱挡住
  Chrome headless                      —— 可用  ← 走这条

用法：
  python3 svg2png.py in.svg out.png [--width 2000 --height 1125]
"""

import argparse
import os
import shutil
import subprocess
import sys
import tempfile
import time

CHROME = os.environ.get(
    "DSH_CHROME", "/Applications/Google Chrome.app/Contents/MacOS/Google Chrome")


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("svg")
    ap.add_argument("png")
    ap.add_argument("--width", type=int, default=2000)
    ap.add_argument("--height", type=int, default=1125)
    ap.add_argument("--timeout", type=int, default=40)
    args = ap.parse_args()

    if not os.path.exists(CHROME):
        sys.exit(f"找不到 Chrome（用 DSH_CHROME 指定）：{CHROME}")

    src = os.path.abspath(args.svg)
    dst = os.path.abspath(args.png)
    os.makedirs(os.path.dirname(dst), exist_ok=True)
    if os.path.exists(dst):
        os.remove(dst)

    # 拷成随机名，避免 Chrome 吃缓存（踩过：改了 SVG 但渲染出来还是旧的）
    tmpd = tempfile.mkdtemp(prefix="svgpng-")
    try:
        copy = os.path.join(tmpd, f"v{int(time.time()*1000)}.svg")
        shutil.copyfile(src, copy)
        profile = os.path.join(tmpd, "profile")
        os.makedirs(profile, exist_ok=True)
        cmd = [CHROME, "--headless=new", "--disable-gpu", "--no-sandbox",
               "--hide-scrollbars", f"--user-data-dir={profile}",
               f"--screenshot={dst}", f"--window-size={args.width},{args.height}",
               f"file://{copy}"]
        p = subprocess.Popen(cmd, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
        t0 = time.time()
        while time.time() - t0 < args.timeout:
            if os.path.exists(dst) and os.path.getsize(dst) > 0:
                break
            if p.poll() is not None:
                break
            time.sleep(0.5)
        try:
            p.kill()
        except Exception:
            pass
        if not os.path.exists(dst) or os.path.getsize(dst) == 0:
            sys.exit("Chrome 没能渲染出图。")
        from PIL import Image
        im = Image.open(dst)
        print(f"OK  {os.path.basename(src)} -> {os.path.basename(dst)}  {im.size[0]}x{im.size[1]}")
    finally:
        shutil.rmtree(tmpd, ignore_errors=True)


if __name__ == "__main__":
    main()
