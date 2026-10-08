#!/usr/bin/env bash
# 端到端跑一遍示例：第 6 页（原生矢量）
# 用法： bash examples/run.sh [页码] [region]
set -e
HERE="$(cd "$(dirname "$0")" && pwd)"
ROOT="$(dirname "$HERE")"
S="$ROOT/skills/redraw-diagram/scripts"
PY="${PYTHON:-python3}"

SLIDE="${1:-6}"
REGION="${2:-}"
WORK="$HERE/_work"
mkdir -p "$WORK"

echo "== 1. 渲染原件 =="
if [ -f "$HERE/dryrun-01/pages/page-0001.png" ] && [ "$SLIDE" = "6" ]; then
  echo "   复用已渲染的 examples/dryrun-01/pages/page-0001.png"
  ORIG="$HERE/dryrun-01/pages/page-0001.png"
else
  "$PY" "$S/render_pptx.py" "$HERE/deck/测试架构图.pptx" --pages "$SLIDE" --outdir "$WORK/orig"
  ORIG="$WORK/orig/page-$(printf %04d "$SLIDE").png"
fi

echo "== 2. 读 -> spec =="
ARGS=(--slide "$SLIDE" --render "$ORIG" --out "$WORK/spec.json")
[ -n "$REGION" ] && ARGS+=(--region "$REGION")
"$PY" "$S/read_slide.py" "$HERE/deck/测试架构图.pptx" "${ARGS[@]}"

echo "== 3. 标号图（给用户核对的） =="
"$PY" "$S/mark_slide.py" "$WORK/spec.json" "$ORIG" --out "$WORK/marked.png"

echo "== 4. 画 =="
"$PY" "$S/build_deck.py" "$WORK/spec.json" --pptx "$WORK/out.pptx" --svg "$WORK/out.svg"

echo "== 5. 自检：渲染回去 + 和原图差分 =="
"$PY" "$S/render_pptx.py" "$WORK/out.pptx" --pages 1 --outdir "$WORK/redraw"
"$PY" "$S/verify.py" "$ORIG" "$WORK/redraw/page-0001.png" --outdir "$WORK/diff"
"$PY" "$S/svg2png.py" "$WORK/out.svg" "$WORK/svg-render.png" || \
  echo "   （SVG 渲染失败，跳过；SVG 仍然生成在 $WORK/out.svg）"

echo
echo "产物都在 $WORK/"
