#!/usr/bin/env python3
"""
svg2png.py —— 把 SVG 渲染成 PNG，用于**自检**。

为什么需要它：这套工具链的 LibreOffice 不支持 svg -> 图，
所以 SVG 一直是"我交出去但自己看不见"的东西，连着错了两次。
本机没有 rsvg-convert / inkscape / cairosvg，但 DSH 自带的 sharp 里
**附带了完整的 libvips，而且 librsvg 是静态链进去的**——
所以可以直接用 libvips 的 C 接口来渲染。

用法：  python3 svg2png.py in.svg out.png
"""

import ctypes
import os
import sys

LIB = ("/Applications/DeepSeek Harness.app/Contents/Resources/app.asar.unpacked/"
       "dsh/node_modules/@img/sharp-libvips-darwin-arm64/lib/libvips-cpp.8.18.7.dylib")


def main():
    if len(sys.argv) < 3:
        sys.exit("用法: svg2png.py <in.svg> <out.png>")
    src, dst = os.path.abspath(sys.argv[1]), os.path.abspath(sys.argv[2])
    if not os.path.exists(LIB):
        sys.exit("找不到 libvips: " + LIB)

    lib = ctypes.CDLL(LIB)
    lib.vips_init.argtypes = [ctypes.c_char_p]
    lib.vips_init.restype = ctypes.c_int
    if lib.vips_init(b"svg2png") != 0:
        sys.exit("vips_init 失败: " + lib.vips_error_buffer().decode())

    lib.vips_error_buffer.restype = ctypes.c_char_p

    lib.vips_image_new_from_file.argtypes = [ctypes.c_char_p, ctypes.c_void_p]
    lib.vips_image_new_from_file.restype = ctypes.c_void_p
    img = lib.vips_image_new_from_file(src.encode(), None)
    if not img:
        sys.exit("读入失败: " + lib.vips_error_buffer().decode())

    lib.vips_image_write_to_file.argtypes = [ctypes.c_void_p, ctypes.c_char_p,
                                             ctypes.c_void_p]
    lib.vips_image_write_to_file.restype = ctypes.c_int
    if lib.vips_image_write_to_file(img, dst.encode(), None) != 0:
        sys.exit("写出失败: " + lib.vips_error_buffer().decode())

    from PIL import Image
    im = Image.open(dst)
    print(f"OK  {os.path.basename(src)} -> {os.path.basename(dst)}  {im.size[0]}x{im.size[1]}")


if __name__ == "__main__":
    main()
