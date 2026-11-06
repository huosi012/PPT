#!/bin/bash
# 用法：FILL_P1=.. FILL_P2=.. ./build.sh out   → out.docx / out.pdf
set -e
cd /tmp/lo1; rm -rf ub && cp -r un ub && python3 v2.py ub >/dev/null
rm -f $1.docx $1.pdf && (cd ub && zip -qXr ../$1.docx .) && soffice --headless --convert-to pdf $1.docx >/dev/null 2>&1
