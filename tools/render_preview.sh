#!/usr/bin/env bash
# 用 LibreOffice 把 pptx 渲染成 PDF 与逐页 PNG（仅用于预览/自检）
# 用法：tools/render_preview.sh deck.pptx outdir [dpi]
set -euo pipefail
PPTX="$(readlink -f "$1")"; OUT="$(mkdir -p "$2" && readlink -f "$2")"; DPI="${3:-60}"
mkdir -p "$OUT"
SOFFICE_PY="${SOFFICE_PY:-}"
if [ -n "$SOFFICE_PY" ]; then
  python3 "$SOFFICE_PY" --headless --convert-to pdf --outdir "$OUT" "$PPTX" >/dev/null
else
  soffice --headless --convert-to pdf --outdir "$OUT" "$PPTX" >/dev/null
fi
PDF="$OUT/$(basename "${PPTX%.*}").pdf"
rm -f "$OUT"/slide-*.png
pdftoppm -png -r "$DPI" "$PDF" "$OUT/slide"
ls -1 "$OUT"/slide-*.png
