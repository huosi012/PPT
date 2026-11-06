"""比较原文件与新文件主表第 r 行的文字（去标点空白），列出删除/新增片段。用法：python3 -I celldiff.py orig.docx new.docx row"""
import sys, re, zipfile, difflib
from lxml import etree
W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'; ns = {'w': W}
def rowtxt(path, r):
    doc = etree.fromstring(zipfile.ZipFile(path).read('word/document.xml'))
    body = doc.find('w:body', ns)
    t = max(body.findall('w:tbl', ns), key=lambda t: len(t.findall('w:tr', ns)))
    tr = t.findall('w:tr', ns)[r]
    s = ''.join(x.text or '' for x in tr.iter(f'{{{W}}}t'))
    return re.sub(r'[\W_]+', '', s)
a, b = rowtxt(sys.argv[1], int(sys.argv[3])), rowtxt(sys.argv[2], int(sys.argv[3]))
sm = difflib.SequenceMatcher(None, a, b, autojunk=False)
for op, i1, i2, j1, j2 in sm.get_opcodes():
    if op != 'equal': print(op, repr(a[i1:i2]), '->', repr(b[j1:j2]))
