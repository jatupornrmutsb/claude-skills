"""Split runs by script the way Word does: Thai text -> its own run with <w:cs/>.

Usage: python3 thai_runs.py raw.docx out.docx
"""
import re, sys, zipfile
src, dst = sys.argv[1], sys.argv[2]
THAI = re.compile(r"([฀-๿]+)")
RUN = re.compile(r"<w:r>(<w:rPr>.*?</w:rPr>)?<w:t(?: xml:space=\"preserve\")?>(.*?)</w:t></w:r>", re.S)

def thai_rpr(rpr):
    rpr = rpr or "<w:rPr></w:rPr>"
    # schema order: ... sz, szCs, ..., rtl, cs, em, lang ...  -> <w:cs/> must sit before <w:lang>
    if "<w:lang " in rpr:
        return rpr.replace("<w:lang ", "<w:cs/><w:lang ", 1)
    return rpr.replace("</w:rPr>", "<w:cs/></w:rPr>", 1)

def split_run(m):
    rpr, text = m.group(1) or "", m.group(2)
    out = []
    for p in [p for p in THAI.split(text) if p]:
        r = thai_rpr(rpr) if THAI.fullmatch(p) else rpr
        out.append(f'<w:r>{r}<w:t xml:space="preserve">{p}</w:t></w:r>')
    return "".join(out)

zin, zout = zipfile.ZipFile(src), zipfile.ZipFile(dst, "w", zipfile.ZIP_DEFLATED)
for it in zin.infolist():
    d = zin.read(it.filename)
    if it.filename in ("word/document.xml",) or re.match(r"word/(header|footer)\d*\.xml", it.filename):
        d = RUN.sub(split_run, d.decode("utf-8")).encode("utf-8")
    zout.writestr(it, d)
zout.close()
