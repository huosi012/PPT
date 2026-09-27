#!/usr/bin/env python3
r"""Handwritten-notes -> Word (.docx) builder.

Page sources are small line-based text files (see DSL below). Formulas are written in
LaTeX and converted to native Word equations (OMML) with pandoc, then post-processed.

DSL (one paragraph per line, "TAG| content"; blank lines and lines starting with '#' are ignored)
  T|  chapter title                      H1| section heading (一、…)     H2| sub heading
  I1| ① | text     numbered item: number at 1 char, text from the item column
  I2| (1) | text   second-level item (one level deeper)
  B1| text  / B2| text                   bullet under an item / under a sub-item
  P0| text  / P1| text / P2| text        body text at level 0 / 1 / 2
  E0| latex / E1| latex / E2| latex      display equation (left aligned) at level 0 / 1 / 2
  PB|                                    page break
Inline markup inside text:
  $...$              inline formula (LaTeX)
  {flags|text}       styled text; may contain $...$ and nest.  flags (comma separated):
                     r red, pk pink, g gray, b bold, hei 黑体, hy/hg yellow/green highlighter,
                     u underline, uw wavy underline, ud double underline, s strikethrough, sm small,
                     ur/udr/uwr red underline (single/double/wavy), up/udp/uwp pink underline
Inside LaTeX:  \R{..} red   \PK{..} pink   \HY{..} yellow highlight   \HG{..} green highlight
               \UR{..} red + underline   \UD{..} red + double underline
"""
import hashlib
import json
import os
import re
import subprocess
import sys
import tempfile
import zipfile
from xml.sax.saxutils import escape as _esc

from lxml import etree

HERE = os.path.dirname(os.path.abspath(__file__))
M_NS = "http://schemas.openxmlformats.org/officeDocument/2006/math"
W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
NS = {"m": M_NS, "w": W_NS}

RED, PINK, GRAY, NAVY = "C00000", "E0007A", "595959", "1F3864"
HEI = "黑体"
ZWSP = "\u200b"

# text-level indent columns (twips): item number / level-1 text / level-2 text
COL0, COL1, COL2 = 240, 600, 960

# LaTeX colour macros -> a texmath font style used as a marker -> formatting
MARKERS = {
    "R": ("mathtt", "monospace", dict(color=RED)),
    "PK": ("mathcal", "script", dict(color=PINK)),
    "HY": ("mathsf", "sans-serif", dict(hl="yellow")),
    "HG": ("mathfrak", "fraktur", dict(hl="green")),
    "UR": ("mathbb", "double-struck", dict(color=RED, u="single")),
}
SCR2FMT = {scr: fmt for _, (_, scr, fmt) in MARKERS.items()}


def esc(s):
    return _esc(s, {'"': "&quot;"})


# ------------------------------------------------------------------ LaTeX -> OMML ---
_CACHE_FILE = os.path.join(HERE, ".omml_cache.json")
try:
    _CACHE = json.load(open(_CACHE_FILE, encoding="utf-8"))
except Exception:
    _CACHE = {}


def _pair_bars(tex):
    """|x| -> \\left|x\\right| (pairs bare bars at the same brace depth, left to right).
    Bare bars become plain runs, which LibreOffice's formula import misreads as 'or'."""
    out, stack = list(tex), [[]]
    i = 0
    while i < len(tex):
        c = tex[i]
        if c == "\\":
            m = re.match(r"\\(left|right|middle|big|Big|bigg|Bigg)[lr]?\s*\|", tex[i:])
            i += m.end() if m else 2
            continue
        if c == "{":
            stack.append([])
        elif c == "}":
            bars = stack.pop()
            if len(bars) % 2 == 0:
                for k, pos in enumerate(bars):
                    out[pos] = r"\left|" if k % 2 == 0 else r"\right|"
        elif c == "|":
            stack[-1].append(i)
        i += 1
    bars = stack[0]
    if len(bars) % 2 == 0:
        for k, pos in enumerate(bars):
            out[pos] = r"\left|" if k % 2 == 0 else r"\right|"
    return "".join(out)


def _macros(tex):
    for mac, (style, _, _) in MARKERS.items():
        tex = re.sub(r"\\%s\{" % mac, r"\\%s{" % style, tex)
    tex = re.sub(r"\\UD\{", r"\\boldsymbol{", tex)  # red double underline (see fix_omml bold_marker)
    return _pair_bars(tex)


def _pandoc(texs):
    """Convert each LaTeX string (own paragraph) with pandoc; returns inner OMML strings."""
    import pypandoc
    with tempfile.TemporaryDirectory() as td:
        md = os.path.join(td, "m.md")
        out = os.path.join(td, "m.docx")
        with open(md, "w", encoding="utf-8") as f:
            f.write("\n\n".join("$$%s$$" % t for t in texs) + "\n")
        pypandoc.convert_file(md, "docx", outputfile=out)
        xml = zipfile.ZipFile(out).read("word/document.xml")
    root = etree.fromstring(xml)
    paras = root.findall(".//w:body/w:p", NS)
    if len(paras) != len(texs):
        raise RuntimeError("pandoc paragraph count mismatch: %d vs %d" % (len(paras), len(texs)))
    res = []
    for tex, p in zip(texs, paras):
        om = p.find(".//m:oMath", NS)
        if om is None:
            raise ValueError("LaTeX not converted: %r" % tex)
        res.append("".join(etree.tostring(c, encoding="unicode") for c in om))
    return res


def tex_batch(texs):
    todo = sorted({t for t in texs if t not in _CACHE})
    if todo:
        for tex, xml in zip(todo, _pandoc([_macros(t) for t in todo])):
            _CACHE[tex] = xml
        # several builds may run at once: merge with what is on disk, write atomically
        try:
            disk = json.load(open(_CACHE_FILE, encoding="utf-8"))
        except Exception:
            disk = {}
        disk.update(_CACHE)
        tmp = "%s.%d.tmp" % (_CACHE_FILE, os.getpid())
        json.dump(disk, open(tmp, "w", encoding="utf-8"), ensure_ascii=False)
        os.replace(tmp, _CACHE_FILE)
    return [_CACHE[t] for t in texs]


def _wrpr_math(fmt):
    p = ['<w:rFonts w:ascii="Cambria Math" w:hAnsi="Cambria Math"/>']
    if fmt.get("color"):
        p.append('<w:color w:val="%s"/>' % fmt["color"])
    if fmt.get("size"):
        p.append('<w:sz w:val="%d"/><w:szCs w:val="%d"/>' % (fmt["size"], fmt["size"]))
    if fmt.get("hl"):
        p.append('<w:highlight w:val="%s"/>' % fmt["hl"])
    if fmt.get("u"):
        p.append('<w:u w:val="%s" w:color="%s"/>' % (fmt["u"], fmt.get("ucolor") or fmt.get("color") or "auto"))
    return "<w:rPr xmlns:w=\"%s\">%s</w:rPr>" % (W_NS, "".join(p))


LETTER = re.compile(r"^[A-Za-z\u0391-\u03A9\u03B1-\u03C9]$")
PR_OF = {"d": "dPr", "f": "fPr", "rad": "radPr", "nary": "naryPr", "m": "mPr",
         "sSup": "sSupPr", "sSub": "sSubPr", "sSubSup": "sSubSupPr", "func": "funcPr",
         "limLow": "limLowPr", "limUpp": "limUppPr", "acc": "accPr", "bar": "barPr",
         "groupChr": "groupChrPr", "eqArr": "eqArrPr", "box": "boxPr", "borderBox": "borderBoxPr"}


def fix_omml(inner, base_fmt=None, bold_marker=False):
    """Post-process pandoc OMML: colour markers -> real formatting, LibreOffice-safe fixes."""
    root = etree.fromstring('<root xmlns:m="%s" xmlns:w="%s">%s</root>' % (M_NS, W_NS, inner))
    m = lambda tag: "{%s}%s" % (M_NS, tag)
    run_fmt = {}
    for r in root.iter(m("r")):
        fmt = dict(base_fmt or {})
        rpr = r.find("m:rPr", NS)
        if rpr is not None:
            scr = rpr.find("m:scr", NS)
            if scr is not None and scr.get(m("val")) in SCR2FMT:
                fmt.update(SCR2FMT[scr.get(m("val"))])
                rpr.remove(scr)
                t = r.find("m:t", NS)
                if t is not None and LETTER.match(t.text or ""):
                    sty = rpr.find("m:sty", NS)
                    if sty is not None:
                        rpr.remove(sty)
            sty = rpr.find("m:sty", NS)
            if bold_marker and sty is not None and sty.get(m("val")) in ("b", "bi"):
                fmt.update(color=RED, u="double")  # \\UD{...} arrives as \\boldsymbol
                if sty.get(m("val")) == "bi":
                    rpr.remove(sty)
                else:
                    sty.set(m("val"), "p")
            if len(rpr) == 0:
                r.remove(rpr)
        if fmt:
            w = etree.fromstring(_wrpr_math(fmt))
            rpr = r.find("m:rPr", NS)
            r.insert(1 if rpr is not None else 0, w)
            run_fmt[r] = json.dumps(fmt, sort_keys=True)
    # colour delimiters / fraction bars etc. when everything inside shares one format
    for tag, prtag in PR_OF.items():
        for el in root.iter(m(tag)):
            runs = list(el.iter(m("r")))
            if not runs:
                continue
            fmts = {run_fmt.get(r) for r in runs}
            if len(fmts) != 1 or None in fmts:
                continue
            fmt = json.loads(fmts.pop())
            pr = el.find("m:" + prtag, NS)
            if pr is None:
                pr = etree.Element(m(prtag))
                el.insert(0, pr)
            old = pr.find("m:ctrlPr", NS)
            if old is not None:
                pr.remove(old)
            ctrl = etree.SubElement(pr, m("ctrlPr"))
            ctrl.append(etree.fromstring(_wrpr_math(fmt)))
    top = [c for c in root if c.tag == m("r")]
    loose = []
    if top and root[0] is top[0]:
        loose.append(top[0])
    if top and root[len(root) - 1] is top[-1]:
        loose.append(top[-1])
    for r in list(root.iter(m("r"))):
        t = r.find("m:t", NS)
        txt = (t.text or "").strip() if t is not None else ""
        par = r.getparent()
        first = None  # first visible child of a row/cell (skip zero-width filler runs)
        if par is not None and par.tag == m("e"):
            for c in par:
                ct = c.find("m:t", NS) if c.tag == m("r") else None
                if ct is not None and not (ct.text or "").replace(ZWSP, "").strip():
                    continue
                first = c
                break
        lead = first is r
        ops = ("=", "<", ">", "\u2264", "\u2265", "\u2260", "\u21d4", "\u21d2", "\u2192")
        if (txt == "/" and len(par) == 1) or ((r in loose or lead) and txt in ops):
            rpr = r.find("m:rPr", NS)
            if rpr is None:
                rpr = etree.Element(m("rPr"))
                r.insert(0, rpr)
            if rpr.find("m:nor", NS) is None:
                rpr.insert(0, etree.Element(m("nor")))
    for rpr in root.iter(m("rPr")):  # schema: m:nor excludes m:scr/m:sty (pandoc's \\text emits both)
        if rpr.find("m:nor", NS) is not None:
            for tag in ("scr", "sty"):
                el = rpr.find("m:" + tag, NS)
                if el is not None:
                    rpr.remove(el)
    for t in root.iter(m("t")):  # ASCII * reads as an operator in LibreOffice; U+2217 looks the same in Word
        if t.text and "*" in t.text:
            t.text = t.text.replace("*", "\u2217")
    for mcpr in root.iter(m("mcPr")):  # schema order is count, mcJc (pandoc writes the reverse)
        cnt = mcpr.find("m:count", NS)
        if cnt is not None and mcpr.index(cnt) != 0:
            mcpr.remove(cnt)
            mcpr.insert(0, cnt)
    for ch in root.iter(m("begChr"), m("endChr"), m("sepChr")):
        if ch.get(m("val")) == "∣":
            ch.set(m("val"), "|")
    for e in root.iter(m("e")):
        if len(e) == 0 and not (e.text or "").strip():
            r = etree.SubElement(e, m("r"))
            t = etree.SubElement(r, m("t"))
            t.text = ZWSP
            t.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
    s = "".join(etree.tostring(c, encoding="unicode") for c in root)
    return re.sub(r' xmlns:\w+="[^"]*"', "", s)  # prefixes are declared on the document root


# ------------------------------------------------------------------ inline parsing ---
FLAG_FMT = {
    "r": dict(color=RED), "pk": dict(color=PINK), "g": dict(color=GRAY), "b": dict(b=True),
    "hei": dict(hei=True), "hy": dict(hl="yellow"), "hg": dict(hl="green"),
    "u": dict(u="single"), "uw": dict(u="wave"), "ud": dict(u="double"), "s": dict(strike=True),
    "sm": dict(size=21),
    # underline in a pen colour while the text keeps its own colour
    "ur": dict(u="single", ucolor=RED), "udr": dict(u="double", ucolor=RED), "uwr": dict(u="wave", ucolor=RED),
    "up": dict(u="single", ucolor=PINK), "udp": dict(u="double", ucolor=PINK), "uwp": dict(u="wave", ucolor=PINK),
}
STYLED = re.compile(r"\{([a-z,]+)\|")


def _find_close(s, i):
    """s[i] is just after '{flags|'; return index of the matching '}' (skips $..$ math)."""
    depth, math = 1, False
    while i < len(s):
        c = s[i]
        if c == "\\" and math:
            i += 2
            continue
        if c == "$":
            math = not math
        elif not math and c == "{":
            depth += 1
        elif not math and c == "}":
            depth -= 1
            if depth == 0:
                return i
        i += 1
    raise ValueError("unclosed {…| in: " + s)


def parse_inline(s, fmt=None):
    """-> list of ('text', str, fmt) / ('math', tex, fmt)"""
    fmt = fmt or {}
    out, i, buf = [], 0, ""
    while i < len(s):
        mt = STYLED.match(s, i)
        if s[i] == "$":
            j = s.index("$", i + 1)
            if buf:
                out.append(("text", buf, fmt))
                buf = ""
            out.append(("math", s[i + 1:j], fmt))
            i = j + 1
        elif mt:
            j = _find_close(s, mt.end())
            if buf:
                out.append(("text", buf, fmt))
                buf = ""
            sub = dict(fmt)
            for fl in mt.group(1).split(","):
                sub.update(FLAG_FMT[fl])
            out.extend(parse_inline(s[mt.end():j], sub))
            i = j + 1
        else:
            buf += s[i]
            i += 1
    if buf:
        out.append(("text", buf, fmt))
    return out


def _math_key(tex, fmt):
    """wrap the whole formula in a colour macro when the surrounding text is coloured"""
    if fmt.get("hl") == "yellow":
        return r"\HY{%s}" % tex
    if fmt.get("hl") == "green":
        return r"\HG{%s}" % tex
    if fmt.get("color") == RED:
        return r"\UR{%s}" % tex if fmt.get("u") else r"\R{%s}" % tex
    if fmt.get("color") == PINK:
        return r"\PK{%s}" % tex
    return tex


def text_run(text, fmt):
    p = []
    if fmt.get("hei"):
        p.append('<w:rFonts w:eastAsia="%s"/>' % HEI)
    if fmt.get("b"):
        p.append("<w:b/><w:bCs/>")
    if fmt.get("strike"):
        p.append("<w:strike/>")
    if fmt.get("color"):
        p.append('<w:color w:val="%s"/>' % fmt["color"])
    if fmt.get("size"):
        p.append('<w:sz w:val="%d"/><w:szCs w:val="%d"/>' % (fmt["size"], fmt["size"]))
    if fmt.get("hl"):
        p.append('<w:highlight w:val="%s"/>' % fmt["hl"])
    if fmt.get("u"):
        uc = fmt.get("ucolor") or fmt.get("color")
        p.append('<w:u w:val="%s"%s/>' % (fmt["u"], ' w:color="%s"' % uc if uc else ""))
    rpr = "<w:rPr>%s</w:rPr>" % "".join(p) if p else ""
    return '<w:r>%s<w:t xml:space="preserve">%s</w:t></w:r>' % (rpr, esc(text))


TAB = "<w:r><w:tab/></w:r>"

# ------------------------------------------------------------------ page parsing ---
LEVEL_STYLE = {"P0": "Body0", "P1": "Body1", "P2": "Body2",
               "E0": "Equation0", "E1": "Equation1", "E2": "Equation2"}


def _split_top(rest):
    """split 'num | text' at the first '|' outside {…} (so the number itself may be styled, e.g. {r|④})"""
    depth = 0
    for i, c in enumerate(rest):
        if c == "{":
            depth += 1
        elif c == "}":
            depth -= 1
        elif c == "|" and depth == 0:
            return rest[:i], rest[i + 1:]
    raise ValueError("item line needs 'number | text': %r" % rest)


def parse_page(src):
    blocks = []
    for ln, raw in enumerate(src.splitlines(), 1):
        line = raw.rstrip()
        if not line.strip() or line.lstrip().startswith("#"):
            continue
        if "|" not in line:
            raise ValueError("line %d: missing TAG|: %r" % (ln, line))
        tag, rest = line.split("|", 1)
        tag, rest = tag.strip(), rest.strip()
        if tag in ("I1", "I2"):
            num, text = _split_top(rest)
            blocks.append((tag, num.strip(), text.strip()))
        elif tag in ("T", "H1", "H2", "P0", "P1", "P2", "B1", "B2", "E0", "E1", "E2", "PB"):
            blocks.append((tag, None, rest))
        else:
            raise ValueError("line %d: unknown tag %r" % (ln, tag))
    return blocks


def collect_math(blocks):
    texs = []
    for tag, num, text in blocks:
        if tag.startswith("E"):
            texs.append(text)
        elif tag != "PB":
            for kind, val, fmt in parse_inline(text):
                if kind == "math":
                    texs.append(val)
    return texs


DROP_STRUCK = True  # {s|…} = words the student crossed out; not part of the clean copy


def render_inline(text, omml):
    out = []
    for kind, val, fmt in parse_inline(text):
        if DROP_STRUCK and fmt.get("strike"):
            continue
        if kind == "text":
            out.append(text_run(val, fmt))
        else:
            mfmt = {k: fmt[k] for k in ("color", "hl", "u", "ucolor", "size") if k in fmt}
            out.append("<m:oMath>%s</m:oMath>" % fix_omml(omml[val], mfmt, "\\UD{" in val))
    return "".join(out)


def para(content, style, keep_next=False, numid=None, extra_ppr=""):
    pp = ['<w:pStyle w:val="%s"/>' % style]
    if keep_next:
        pp.append("<w:keepNext/>")
    if numid:
        pp.append('<w:numPr><w:ilvl w:val="0"/><w:numId w:val="%d"/></w:numPr>' % numid)
    pp.append(extra_ppr)
    return "<w:p><w:pPr>%s</w:pPr>%s</w:p>" % ("".join(pp), content)


def render_blocks(blocks):
    texs = collect_math(blocks)
    omml = {}
    if texs:
        for tex, xml in zip(texs, tex_batch(texs)):
            omml[tex] = xml
    out = []
    for k, (tag, num, text) in enumerate(blocks):
        nxt = blocks[k + 1][0] if k + 1 < len(blocks) else None
        keep = nxt is not None and nxt.startswith("E")
        if tag == "PB":
            out.append('<w:p><w:r><w:br w:type="page"/></w:r></w:p>')
        elif tag == "T":
            out.append(para(render_inline(text, omml), "Title"))
        elif tag == "H1":
            out.append(para(render_inline(text, omml), "Heading1"))
        elif tag == "H2":
            out.append(para(render_inline(text, omml), "Heading2", keep))
        elif tag in ("I1", "I2"):
            numrun = "".join(text_run(v, dict(f, hei=True)) for k, v, f in parse_inline(num) if k == "text")
            out.append(para(numrun + TAB + render_inline(text, omml),
                            "Item1" if tag == "I1" else "Item2", keep))
        elif tag in ("B1", "B2"):
            out.append(para(render_inline(text, omml), "Body1", keep, numid=1 if tag == "B1" else 2))
        elif tag.startswith("P"):
            out.append(para(render_inline(text, omml), LEVEL_STYLE[tag], keep))
        elif tag.startswith("E"):
            out.append(para('<m:oMathPara><m:oMathParaPr><m:jc m:val="left"/></m:oMathParaPr>'
                            "<m:oMath>%s</m:oMath></m:oMathPara>" % fix_omml(omml[text], None, "\\UD{" in text),
                            LEVEL_STYLE[tag], keep))
    return out


# ------------------------------------------------------------------ package ---
HDR = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'


def _body_style(sid, name, ind, eq=False):
    sp = ('<w:spacing w:before="60" w:after="60" w:line="240" w:lineRule="auto"/>' if eq
          else '<w:spacing w:before="20" w:after="60"/>')
    return ('<w:style w:type="paragraph" w:customStyle="1" w:styleId="%s"><w:name w:val="%s"/>'
            '<w:basedOn w:val="Normal"/><w:qFormat/><w:pPr>%s<w:ind w:left="%d"/>%s</w:pPr></w:style>'
            % (sid, name, sp, ind, '<w:jc w:val="left"/>' if eq else ""))


def _item_style(sid, name, text_col):
    return ('<w:style w:type="paragraph" w:customStyle="1" w:styleId="%s"><w:name w:val="%s"/>'
            '<w:basedOn w:val="Normal"/><w:qFormat/><w:pPr><w:tabs><w:tab w:val="left" w:pos="%d"/></w:tabs>'
            '<w:spacing w:before="40" w:after="60"/><w:ind w:left="%d" w:hanging="360"/></w:pPr></w:style>'
            % (sid, name, text_col, text_col))


STYLES = HDR + f"""<w:styles xmlns:w="{W_NS}">
<w:docDefaults>
 <w:rPrDefault><w:rPr><w:rFonts w:ascii="Times New Roman" w:eastAsia="宋体" w:hAnsi="Times New Roman" w:cs="Times New Roman"/><w:kern w:val="2"/><w:sz w:val="24"/><w:szCs w:val="24"/><w:lang w:val="en-US" w:eastAsia="zh-CN" w:bidi="ar-SA"/></w:rPr></w:rPrDefault>
 <w:pPrDefault><w:pPr><w:spacing w:after="60" w:line="300" w:lineRule="auto"/></w:pPr></w:pPrDefault>
</w:docDefaults>
<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/><w:qFormat/><w:pPr><w:widowControl w:val="0"/></w:pPr></w:style>
<w:style w:type="character" w:default="1" w:styleId="DefaultParagraphFont"><w:name w:val="Default Paragraph Font"/><w:uiPriority w:val="1"/><w:semiHidden/><w:unhideWhenUsed/></w:style>
<w:style w:type="paragraph" w:styleId="Title"><w:name w:val="Title"/><w:basedOn w:val="Normal"/><w:next w:val="Normal"/><w:uiPriority w:val="10"/><w:qFormat/>
 <w:pPr><w:keepNext/><w:pageBreakBefore/><w:pBdr><w:bottom w:val="single" w:sz="12" w:space="8" w:color="{NAVY}"/></w:pBdr><w:spacing w:before="120" w:after="240"/><w:jc w:val="center"/></w:pPr>
 <w:rPr><w:rFonts w:eastAsia="{HEI}"/><w:b/><w:bCs/><w:color w:val="{NAVY}"/><w:spacing w:val="20"/><w:sz w:val="40"/><w:szCs w:val="40"/></w:rPr></w:style>
<w:style w:type="paragraph" w:styleId="Heading1"><w:name w:val="heading 1"/><w:basedOn w:val="Normal"/><w:next w:val="Normal"/><w:uiPriority w:val="9"/><w:qFormat/>
 <w:pPr><w:keepNext/><w:keepLines/><w:spacing w:before="300" w:after="100"/><w:outlineLvl w:val="0"/></w:pPr>
 <w:rPr><w:rFonts w:eastAsia="{HEI}"/><w:b/><w:bCs/><w:color w:val="{NAVY}"/><w:sz w:val="30"/><w:szCs w:val="30"/></w:rPr></w:style>
<w:style w:type="paragraph" w:styleId="Heading2"><w:name w:val="heading 2"/><w:basedOn w:val="Normal"/><w:next w:val="Normal"/><w:uiPriority w:val="9"/><w:qFormat/>
 <w:pPr><w:keepNext/><w:keepLines/><w:spacing w:before="160" w:after="60"/><w:ind w:left="{COL0}"/><w:outlineLvl w:val="1"/></w:pPr>
 <w:rPr><w:rFonts w:eastAsia="{HEI}"/><w:b/><w:bCs/><w:sz w:val="26"/><w:szCs w:val="26"/></w:rPr></w:style>
{_item_style("Item1", "条目", COL1)}
{_item_style("Item2", "子条目", COL2)}
{_body_style("Body0", "正文0", COL0)}
{_body_style("Body1", "正文1", COL1)}
{_body_style("Body2", "正文2", COL2)}
{_body_style("Equation0", "公式0", COL0, eq=True)}
{_body_style("Equation1", "公式1", COL1, eq=True)}
{_body_style("Equation2", "公式2", COL2, eq=True)}
<w:style w:type="paragraph" w:styleId="Footer"><w:name w:val="footer"/><w:basedOn w:val="Normal"/><w:uiPriority w:val="99"/><w:unhideWhenUsed/>
 <w:pPr><w:spacing w:after="0"/><w:jc w:val="center"/></w:pPr><w:rPr><w:color w:val="7F7F7F"/><w:sz w:val="18"/><w:szCs w:val="18"/></w:rPr></w:style>
</w:styles>"""


def _bullet(aid, text_col):
    return (f'<w:abstractNum w:abstractNumId="{aid}"><w:multiLevelType w:val="hybridMultilevel"/>'
            f'<w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="bullet"/><w:lvlText w:val="\uf0b7"/>'
            f'<w:lvlJc w:val="left"/><w:pPr><w:ind w:left="{text_col}" w:hanging="240"/></w:pPr>'
            f'<w:rPr><w:rFonts w:ascii="Symbol" w:hAnsi="Symbol" w:hint="default"/></w:rPr></w:lvl></w:abstractNum>')


NUMBERING = HDR + (f'<w:numbering xmlns:w="{W_NS}">' + _bullet(0, COL1 + 240) + _bullet(1, COL2 + 240)
                   + '<w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num>'
                   + '<w:num w:numId="2"><w:abstractNumId w:val="1"/></w:num></w:numbering>')

SETTINGS = HDR + f"""<w:settings xmlns:w="{W_NS}" xmlns:m="{M_NS}">
<w:zoom w:percent="100"/><w:defaultTabStop w:val="420"/><w:characterSpacingControl w:val="compressPunctuation"/>
<w:compat><w:compatSetting w:name="compatibilityMode" w:uri="http://schemas.microsoft.com/office/word" w:val="15"/></w:compat>
<m:mathPr><m:mathFont m:val="Cambria Math"/><m:brkBin m:val="before"/><m:brkBinSub m:val="--"/><m:smallFrac m:val="0"/><m:dispDef m:val="0"/><m:lMargin m:val="0"/><m:rMargin m:val="0"/><m:defJc m:val="left"/><m:wrapIndent m:val="1440"/><m:intLim m:val="subSup"/><m:naryLim m:val="undOvr"/></m:mathPr>
<w:themeFontLang w:val="en-US" w:eastAsia="zh-CN"/><w:decimalSymbol w:val="."/><w:listSeparator w:val=","/>
</w:settings>"""

FONTS = HDR + f"""<w:fonts xmlns:w="{W_NS}">
<w:font w:name="Times New Roman"><w:panose1 w:val="02020603050405020304"/><w:charset w:val="00"/><w:family w:val="roman"/><w:pitch w:val="variable"/></w:font>
<w:font w:name="宋体"><w:altName w:val="SimSun"/><w:panose1 w:val="02010600030101010101"/><w:charset w:val="86"/><w:family w:val="auto"/><w:pitch w:val="variable"/></w:font>
<w:font w:name="黑体"><w:altName w:val="SimHei"/><w:panose1 w:val="02010609060101010101"/><w:charset w:val="86"/><w:family w:val="modern"/><w:pitch w:val="fixed"/></w:font>
<w:font w:name="Cambria Math"><w:panose1 w:val="02040503050406030204"/><w:charset w:val="00"/><w:family w:val="roman"/><w:pitch w:val="variable"/></w:font>
<w:font w:name="Symbol"><w:panose1 w:val="05050102010706020507"/><w:charset w:val="02"/><w:family w:val="roman"/><w:pitch w:val="variable"/></w:font>
</w:fonts>"""

FOOTER = HDR + f"""<w:ftr xmlns:w="{W_NS}" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
<w:p><w:pPr><w:pStyle w:val="Footer"/></w:pPr><w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText xml:space="preserve"> PAGE </w:instrText></w:r><w:r><w:fldChar w:fldCharType="separate"/></w:r><w:r><w:t>1</w:t></w:r><w:r><w:fldChar w:fldCharType="end"/></w:r></w:p>
</w:ftr>"""

CT = "application/vnd.openxmlformats-officedocument.wordprocessingml"
CONTENT_TYPES = HDR + f"""<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
<Default Extension="xml" ContentType="application/xml"/>
<Override PartName="/word/document.xml" ContentType="{CT}.document.main+xml"/>
<Override PartName="/word/styles.xml" ContentType="{CT}.styles+xml"/>
<Override PartName="/word/settings.xml" ContentType="{CT}.settings+xml"/>
<Override PartName="/word/numbering.xml" ContentType="{CT}.numbering+xml"/>
<Override PartName="/word/fontTable.xml" ContentType="{CT}.fontTable+xml"/>
<Override PartName="/word/footer1.xml" ContentType="{CT}.footer+xml"/>
<Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>
</Types>"""

REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
ROOT_RELS = HDR + f"""<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="{REL}/officeDocument" Target="word/document.xml"/>
<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/>
</Relationships>"""
DOC_RELS = HDR + f"""<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="{REL}/styles" Target="styles.xml"/>
<Relationship Id="rId2" Type="{REL}/settings" Target="settings.xml"/>
<Relationship Id="rId3" Type="{REL}/numbering" Target="numbering.xml"/>
<Relationship Id="rId4" Type="{REL}/fontTable" Target="fontTable.xml"/>
<Relationship Id="rId5" Type="{REL}/footer" Target="footer1.xml"/>
</Relationships>"""


def core_xml(title):
    return HDR + f"""<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" xmlns:dcmitype="http://purl.org/dc/dcmitype/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">
<dc:title>{esc(title)}</dc:title><dc:subject>线性代数笔记</dc:subject>
<dcterms:created xsi:type="dcterms:W3CDTF">2026-09-27T00:00:00Z</dcterms:created>
<dcterms:modified xsi:type="dcterms:W3CDTF">2026-09-27T00:00:00Z</dcterms:modified>
</cp:coreProperties>"""


SECT = ('<w:sectPr><w:footerReference w:type="default" r:id="rId5"/><w:pgSz w:w="11906" w:h="16838"/>'
        '<w:pgMar w:top="1440" w:right="1418" w:bottom="1440" w:left="1418" w:header="851" w:footer="851" w:gutter="0"/>'
        '<w:cols w:space="425"/><w:docGrid w:linePitch="312"/></w:sectPr>')


def write_docx(paragraphs, out, title="线性代数笔记"):
    doc = (HDR + f'<w:document xmlns:w="{W_NS}" xmlns:r="{REL}" xmlns:m="{M_NS}"><w:body>'
           + "".join(paragraphs) + SECT + "</w:body></w:document>")
    etree.fromstring(doc.encode("utf-8"))  # well-formedness check
    parts = [("[Content_Types].xml", CONTENT_TYPES), ("_rels/.rels", ROOT_RELS),
             ("docProps/core.xml", core_xml(title)), ("word/document.xml", doc),
             ("word/styles.xml", STYLES), ("word/settings.xml", SETTINGS),
             ("word/numbering.xml", NUMBERING), ("word/fontTable.xml", FONTS),
             ("word/footer1.xml", FOOTER), ("word/_rels/document.xml.rels", DOC_RELS)]
    with zipfile.ZipFile(out, "w", zipfile.ZIP_DEFLATED) as z:
        for name, data in parts:
            z.writestr(name, data.encode("utf-8"))


SOFFICE = "/root/.claude/skills/synced/7b9959fd-7175-4ac5-94aa-255a13398b12_8a3492c8-4283-4e89-993b-c1e1f6b45201/docx/scripts/office/soffice.py"


def render_preview(docx, out_prefix, dpi=110):
    """docx -> pdf (LibreOffice) -> one PNG per page; returns list of PNG paths"""
    import pymupdf
    outdir = os.path.dirname(os.path.abspath(docx))
    subprocess.run([sys.executable, SOFFICE, "--headless", "--convert-to", "pdf", "--outdir", outdir, docx],
                   check=True, capture_output=True, timeout=180)
    pdf = os.path.splitext(docx)[0] + ".pdf"
    pngs = []
    for i, pg in enumerate(pymupdf.open(pdf)):
        fn = "%s_%d.png" % (out_prefix, i + 1)
        pg.get_pixmap(dpi=dpi).save(fn)
        pngs.append(fn)
    return pngs
