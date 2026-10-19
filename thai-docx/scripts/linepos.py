import subprocess,re,sys

for _s in (sys.stdout, sys.stderr):  # Windows console ไม่ใช่ UTF-8 จะ print ภาษาไทยไม่ได้
    if hasattr(_s, "reconfigure"):
        _s.reconfigure(encoding="utf-8", errors="replace")
def lines(pdf,page=1):
    out=subprocess.run(['pdftotext','-f',str(page),'-l',str(page),'-bbox-layout',pdf,'-'],capture_output=True,text=True).stdout
    res=[]
    for m in re.finditer(r'<line xMin="([\d.]+)" yMin="([\d.]+)" xMax="([\d.]+)" yMax="([\d.]+)">(.*?)</line>',out,re.S):
        words=re.findall(r'>([^<]+)</word>',m.group(5))
        res.append((float(m.group(2)),float(m.group(1)),float(m.group(3)),float(m.group(4)),' '.join(words)[:46]))
    return res
for pdf in sys.argv[1:]:
    print('==',pdf)
    for y,x0,x1,y1,t in lines(pdf): print(f'y={y:6.1f}-{y1:6.1f} x={x0:6.1f}-{x1:6.1f} {t}')
