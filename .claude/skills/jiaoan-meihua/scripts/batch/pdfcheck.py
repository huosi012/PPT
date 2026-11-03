"""检查 docx 中每段文字是否都出现在 PDF 中（防止渲染丢内容）。用法：python3 -I pdfcheck.py <unpacked_dir> <pdf>"""
import sys, re, subprocess
from lxml import etree
W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'; q = lambda n: f'{{{W}}}{n}'
t = etree.parse(sys.argv[1] + '/word/document.xml')
pdf = re.sub(r'[\W_]+', '', subprocess.run(['pdftotext', sys.argv[2], '-'], capture_output=True, text=True).stdout)
miss = []
for p in t.iter(q('p')):
    a, skip = p.getparent(), False
    while a is not None:
        if etree.QName(a).localname in ('txbxContent', 'AlternateContent', 'pict'): skip = True; break
        a = a.getparent()
    if skip: continue
    s = re.sub(r'[\W_]+', '', ''.join(x.text or '' for x in p.iter(q('t'))))
    if len(s) >= 6:
        sh = [s[i:i + 4] for i in range(0, len(s) - 3, 2)]
        if sum(x in pdf for x in sh) < 0.5 * len(sh): miss.append(s[:30])
n_pages = len(subprocess.run(['pdftotext', sys.argv[2], '-'], capture_output=True, text=True).stdout.split('\f')) - 1
print(f'pages={n_pages} missing={len(miss)}', miss[:5])
