"""逐页排满（通用版）。用法：python3 -I fill2.py <id>
- 用每行内容中“全文唯一”的文字片段，在 PDF 文本里定位行所在页
- 分页点只放在“整组结束”的行之后（不拆纵向合并组；分区栏/阶段行/表头行跟随下一行）
- 某页若以一个跨页的大行结尾，则该页本身已满，不加留白
- 末页：调教学反思行高把页面排满
"""
import sys, os, re, json, subprocess, shutil
from lxml import etree
W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'; ns = {'w': W}; q = lambda n: f'{{{W}}}{n}'
S = os.environ.get('JM_WORK', os.path.abspath('jm_work'))   # 工作目录：in/<id>/x、cfg/<id>.py、work/
HERE = os.path.dirname(os.path.abspath(__file__))
ID = sys.argv[1]
PROFILE = f'file:///tmp/lo_profile_{ID}'
LOG = open(f'{S}/work/{ID}.fill.log', 'w')

def log(*a):
    print(*a, file=LOG, flush=True)

def norm(s): return re.sub(r'[\W_]+', '', s)

def build(plan, reflect_h, mark=True):
    env = dict(os.environ, FILL_PLAN=json.dumps(plan), REFLECT_H=str(reflect_h), ROW_MARK='1' if mark else '0')
    subprocess.run(['python3', '-I', f'{HERE}/build.py', f'{S}/in/{ID}/x', f'{S}/work/{ID}', f'{S}/cfg/{ID}.py'],
                   env=env, check=True, capture_output=True)
    out = f'{S}/work/{ID}_fill'
    os.makedirs(out, exist_ok=True)
    for f in ('t.docx', 't.pdf'):
        if os.path.exists(f'{out}/{f}'): os.remove(f'{out}/{f}')
    subprocess.run(f'cd {S}/work/{ID} && zip -qXr {out}/t.docx .', shell=True, check=True)
    subprocess.run(['soffice', f'-env:UserInstallation={PROFILE}', '--headless', '--convert-to', 'pdf', '--outdir', out,
                    f'{out}/t.docx'], capture_output=True)
    x = subprocess.run(['pdftotext', f'{out}/t.pdf', '-'], capture_output=True, text=True).stdout
    return [norm(p) for p in x.split('\f')][:-1]

def row_snippets(P):
    t = etree.parse(f'{S}/work/{ID}/word/document.xml')
    body = t.getroot().find('w:body', ns)
    tbl = max(body.findall('w:tbl', ns), key=lambda x: len(x.findall('w:tr', ns)))
    rows = tbl.findall('w:tr', ns)
    def ptxt(p):
        out = []
        for x in p.iter(q('t')):
            a, inside = x.getparent(), False
            while a is not None and a is not p:
                if etree.QName(a).localname in ('drawing', 'pict', 'AlternateContent', 'txbxContent'): inside = True; break
                a = a.getparent()
            if not inside: out.append(x.text or '')
        return norm(''.join(out))
    alltxt = ''.join(ptxt(p) for r in rows for p in r.iter(q('p')))
    snips, vcont = [], []
    for r in rows:
        ss = []
        ps = [ptxt(p) for p in r.iter(q('p'))]
        for tx in ps:
            for L in (8, 12, 16):
                if len(tx) >= L:
                    for pos in range(0, len(tx) - L + 1, max(1, L // 2)):
                        c = tx[pos:pos + L]
                        if alltxt.count(c) == 1 and sum(pg.count(c) for pg in P) == 1: ss.append(c); break
                    else: continue
                    break
        ss.append(f'QZX{len(snips):03d}QZX')          # 行首定位码（必有）
        snips.append(ss)
        first = r.find('w:tc', ns)
        vm = first.find('w:tcPr/w:vMerge', ns)
        vcont.append(vm is not None and vm.get(q('val')) != 'restart')
    return snips, vcont

def pages_of(P, ss):
    return {i for s in ss for i, p in enumerate(P) if s in p}

info = json.load(open(f'{S}/work/{ID}.rows.json'))
N = info['n']
glue = set(info['sec']) | set(info['phase']) | set(info['head'])   # 这些行跟随下一行
refl = info['refl']
first_refl_block = min(refl) - 1        # 教学反思分区栏
plan, REF_H = [], 1400
P = build(plan, REF_H)
# 若教学反思行单独挤到末页，先把反思行压矮，争取少一页
P2 = build(plan, 700)
if len(P2) < len(P):
    REF_H, P = 700, P2
    log(f'reflection rows shrink: pages {len(P)}')
SN, VC = row_snippets(P)
log('rows without snippet:', [i for i, x in enumerate(SN) if not x])

def allowed_break(r):
    return r not in glue and (r + 1 >= N or not VC[r + 1])

def rp(P, r):
    if not SN[r] and r in glue:          # 分区栏/阶段行/表头行无唯一文字：跟随下一行所在页
        return rp_next(P, r + 1) or {0}
    k = r
    while k >= 0 and not SN[k]: k -= 1
    return pages_of(P, SN[k]) if k >= 0 else {0}

def rp_next(P, r):
    k = r
    while k < N and not SN[k]: k += 1
    return pages_of(P, SN[k]) if k < N else set()

start, page = 0, 0
while start < first_refl_block and page < len(P) - 1:
    # 本页完整容纳的最后一行
    e0 = start - 1
    for r in range(start, N):
        pg = rp(P, r)
        if pg == {page}: e0 = r
        else: break
    nxt = e0 + 1
    pn = rp(P, nxt) if nxt < N else set()
    if pn and min(pn) == page and len(pn) > 1:
        # 本页以跨页大行结尾：已满
        log(f'page {page + 1}: rows {start}-{e0} + split row {nxt} -> full')
        page = max(pn); start = nxt + 1
        # 跳过与大行同页的后续行由下一轮处理
        continue
    cands = [r for r in range(start, e0 + 1) if allowed_break(r)]
    if not cands:
        log(f'page {page + 1}: no break candidate in {start}-{e0}'); page += 1; start = e0 + 1; continue
    e = max(cands)
    def ok(P2, e=e, page=page):
        a, b = rp(P2, e), rp_next(P2, e + 1)
        return a == {page} and (not b or min(b) == page + 1)
    lo = 0 if (e == e0 and ok(P)) else None
    if lo is None:
        for E in range(100, 1300, 100):
            if ok(build(plan + [[start, e, E]], REF_H)): lo = E; break
    if lo is None:
        log(f'page {page + 1}: cannot end at {e}; keep natural'); start = e0 + 1; page += 1; continue
    hi = 1300
    while hi - lo > 25:
        mid = (lo + hi) // 2
        if ok(build(plan + [[start, e, mid]], REF_H)): lo = mid
        else: hi = mid
    E = max(0, lo - 25)
    if E > 700:
        log(f'page {page + 1}: rows {start}-{e} would need E={E} (>700), capped')
        E = 700 if ok(build(plan + [[start, e, 700]], REF_H)) else lo
    plan.append([start, e, E])
    log(f'page {page + 1}: rows {start}-{e}, E={E}')
    P = build(plan, REF_H)
    start, page = e + 1, page + 1

# 末页：教学反思行高
n_pages = len(P)
lo, hi = REF_H, 6000
while hi - lo > 50:
    mid = (lo + hi) // 2
    if len(build(plan, mid)) == n_pages: lo = mid
    else: hi = mid
REF_H = max(REF_H, lo - 60)
P = build(plan, REF_H, mark=False)
log(f'reflect_h={REF_H} pages={len(P)}')
json.dump({'plan': plan, 'reflect_h': REF_H, 'pages': len(P)}, open(f'{S}/work/{ID}.fill.json', 'w'))
shutil.copy(f'{S}/work/{ID}_fill/t.docx', f'{S}/work/{ID}_final.docx')
shutil.copy(f'{S}/work/{ID}_fill/t.pdf', f'{S}/work/{ID}_final.pdf')
log('done')
