#!/bin/bash
# finalize.sh <id>：按 work/<id>.fill.json 的排版参数重建最终文件（无定位码）
S=${JM_WORK:-$PWD/jm_work}
HERE=$(cd "$(dirname "$0")" && pwd)
i=$1
PLAN=$(python3 -I -c "import json;print(json.dumps(json.load(open('$S/work/$i.fill.json'))['plan']))")
RH=$(python3 -I -c "import json;print(json.load(open('$S/work/$i.fill.json'))['reflect_h'])")
FILL_PLAN="$PLAN" REFLECT_H=$RH ROW_MARK=0 python3 -I $HERE/build.py $S/in/$i/x $S/work/$i $S/cfg/$i.py >/dev/null
rm -f $S/work/${i}_final.docx $S/work/${i}_final.pdf
(cd $S/work/$i && zip -qXr ../${i}_final.docx .)
(cd $S/work && soffice -env:UserInstallation=file:///tmp/lo_profile_$i --headless --convert-to pdf ${i}_final.docx >/dev/null 2>&1)
