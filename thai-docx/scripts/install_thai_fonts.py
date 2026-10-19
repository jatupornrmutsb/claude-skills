#!/usr/bin/env python3
"""ถอดฟอนต์ที่ฝังอยู่ในไฟล์ .docx (word/fonts/*.odttf) แล้วติดตั้งลง ~/.fonts
เพื่อให้ LibreOffice render เอกสารไทยด้วยฟอนต์จริง (sandbox ไม่มี TH Sarabun)
ใช้: python3 install_thai_fonts.py ตัวอย่าง.docx
อ้างอิง: ECMA-376 — ฟอนต์ฝังถูก XOR 32 ไบต์แรกด้วย GUID (ไบต์กลับด้าน)
"""
import os, re, sys, zipfile, subprocess

for _s in (sys.stdout, sys.stderr):  # Windows console ไม่ใช่ UTF-8 จะ print ภาษาไทยไม่ได้
    if hasattr(_s, "reconfigure"):
        _s.reconfigure(encoding="utf-8", errors="replace")

def main(path):
    z = zipfile.ZipFile(path)
    names = z.namelist()
    if 'word/fontTable.xml' not in names:
        sys.exit('ไม่พบ fontTable.xml')
    ft = z.read('word/fontTable.xml').decode('utf8')
    rels = z.read('word/_rels/fontTable.xml.rels').decode('utf8') if 'word/_rels/fontTable.xml.rels' in names else ''
    rmap = {}
    for rel in re.findall(r'<Relationship [^>]*/>', rels):
        i = re.search(r'Id="([^"]+)"', rel); t = re.search(r'Target="([^"]+)"', rel)
        if i and t: rmap[i.group(1)] = t.group(1)
    out = os.path.expanduser('~/.fonts'); os.makedirs(out, exist_ok=True)
    n = 0
    for fm in re.finditer(r'<w:font w:name="([^"]+)">(.*?)</w:font>', ft, re.S):
        for kind, rid, key in re.findall(r'<w:embed(\w+) r:id="(\w+)" w:fontKey="([^"]+)"', fm.group(2)):
            kb = bytes.fromhex(key.strip('{}').replace('-', ''))[::-1]
            data = bytearray(z.read('word/' + rmap[rid]))
            for i in range(32): data[i] ^= kb[i % 16]
            fn = os.path.join(out, re.sub(r'\W', '_', fm.group(1)) + '_' + kind + '.ttf')
            open(fn, 'wb').write(data); n += 1; print('ติดตั้ง', fm.group(1), kind)
    subprocess.run(['fc-cache', '-f'], check=False)
    print('เสร็จ', n, 'ไฟล์ -> ตรวจด้วย: fc-list | grep -i sarabun')

if __name__ == '__main__':
    main(sys.argv[1]) if len(sys.argv) > 1 else sys.exit(__doc__)
