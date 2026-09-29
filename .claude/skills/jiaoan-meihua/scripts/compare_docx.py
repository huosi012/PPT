import zipfile, re, hashlib
from lxml import etree
W='http://schemas.openxmlformats.org/wordprocessingml/2006/main';ns={'w':W};q=lambda n:f'{{{W}}}{n}'
R='http://schemas.openxmlformats.org/officeDocument/2006/relationships'
A='http://schemas.openxmlformats.org/drawingml/2006/main'
def load(path):
    z=zipfile.ZipFile(path)
    doc=etree.fromstring(z.read('word/document.xml'))
    rels=etree.fromstring(z.read('word/_rels/document.xml.rels'))
    rmap={r.get('Id'):r.get('Target') for r in rels}
    def md5(rid):
        t=rmap.get(rid); 
        try: return hashlib.md5(z.read('word/'+t)).hexdigest()[:8]+':'+t.split('/')[-1]
        except Exception: return 'MISSING:'+str(t)
    return doc, md5, z
def norm(s): return re.sub(r'\s+','',s)
def cellinfo(tc, md5):
    pr=tc.find('w:tcPr',ns)
    gs=pr.find('w:gridSpan',ns) if pr is not None else None
    vm=pr.find('w:vMerge',ns) if pr is not None else None
    span=int(gs.get(q('val'))) if gs is not None else 1
    v=None if vm is None else (vm.get(q('val')) or 'continue')
    imgs=[md5(b.get(f'{{{R}}}embed')) for b in tc.iter(f'{{{A}}}blip')]
    imgs+=['VML:'+md5(i.get(f'{{{R}}}id')) for i in tc.iter('{urn:schemas-microsoft-com:vml}imagedata')]
    return dict(span=span, vm=v, text=norm(''.join(tc.itertext())), imgs=imgs, nested=len(tc.findall('.//w:tbl',ns)))
def rows_of(doc, md5):
    body=doc.find('w:body',ns)
    tbls=body.findall('w:tbl',ns)
    main=max(tbls, key=lambda t: len(t.findall('w:tr',ns)))
    grid=[g.get(q('w')) for g in main.find('w:tblGrid',ns)]
    return grid, [[cellinfo(tc,md5) for tc in tr.findall('w:tc',ns)] for tr in main.findall('w:tr',ns)], tbls, body
