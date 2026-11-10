#!/bin/bash
# strip.sh <pdf> <out.png> [dpi]
S=${JM_WORK:-$PWD/jm_work}
HERE=$(cd "$(dirname "$0")" && pwd)
T=$(mktemp -d -p $S); pdftoppm -r ${3:-40} -png $1 $T/p
python3 -I -c "
from PIL import Image
import glob,sys
fs=sorted(glob.glob('$T/p-*.png'), key=lambda f:int(f.split('-')[-1][:-4]))
ims=[Image.open(f) for f in fs]; w,h=ims[0].size
o=Image.new('RGB',(w*len(ims)+6*len(ims),h),'#888')
for k,im in enumerate(ims): o.paste(im,(k*(w+6),0))
o.save('$2')"
rm -rf $T
