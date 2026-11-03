#!/bin/bash
# 用法：run.sh <id> [render]   环境变量 FILL_PLAN / REFLECT_H 透传给 build.py
set -e
S=${JM_WORK:-$PWD/jm_work}
HERE=$(cd "$(dirname "$0")" && pwd)
i=$1
python3 -I $HERE/build.py $S/in/$i/x $S/work/$i $S/cfg/$i.py >/dev/null
rm -f $S/work/$i.docx $S/work/$i.pdf
(cd $S/work/$i && zip -qXr ../$i.docx .)
if [ "$2" = "render" ]; then
  (cd $S/work && soffice -env:UserInstallation=file:///tmp/lo_profile_$i --headless --convert-to pdf $i.docx >/dev/null 2>&1)
fi
