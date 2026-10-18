#!/usr/bin/env python3
"""เรนเดอร์ .docx เป็น PDF/PNG แล้วรายงานจำนวนหน้าและจุดเสี่ยงการจัดหน้า

  python render_check.py article.docx --outdir WORKDIR [--png] [--dpi 110]
         [--min-pages 8 --max-pages 10] [--count main|main+refs|all]

ต้องมี LibreOffice (soffice) และ poppler-utils (pdfinfo, pdftotext, pdffonts, pdftoppm)

รายงาน:
- จำนวนหน้าทั้งหมด และหน้าเริ่ม “เอกสารอ้างอิง” / “ภาคผนวก” (แยกนับได้)
- ฟอนต์ที่ฝังจริงใน PDF และเตือนเมื่อ TH Sarabun New ถูกแทนที่ (จำนวนหน้าไม่ใช่ค่ายืนยัน)
- คำที่ล้นขอบซ้าย/ขวาของพื้นที่พิมพ์ (ตาราง/รูป/ลิงก์ยาวข้ามขอบ)
- อักขระแฝงที่หลุดไปใน PDF
PNG ที่ได้ใช้ตรวจด้วยตา (ไม่ส่งให้ผู้ใช้โดยอัตโนมัติ)

ข้อสำคัญ: LibreOffice ไม่ใช่ Word — การตัดคำไทยและจำนวนหน้าอาจต่างจาก Word
หากต้องยืนยันการแสดงผลใน Word ให้เปิดตรวจใน Microsoft Word อีกครั้ง
"""
from __future__ import annotations

import argparse
import json
import os
import re
import shutil
import subprocess
import sys
import tempfile
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))
import docx_utils as U  # noqa: E402

PT_PER_CM = 28.3465


def run(cmd, **kw):
    return subprocess.run(cmd, capture_output=True, text=True, **kw)


def to_pdf(docx, outdir):
    soffice = shutil.which("soffice") or shutil.which("libreoffice")
    if not soffice:
        raise SystemExit("ไม่พบ LibreOffice (soffice) — เรนเดอร์ไม่ได้: ห้ามรับรองการจัดหน้า")
    prof = tempfile.mkdtemp(prefix="lo_profile_")
    env = dict(os.environ, HOME=prof)
    r = run([soffice, f"-env:UserInstallation=file://{prof}", "--headless", "--convert-to", "pdf",
             "--outdir", str(outdir), str(docx)], env=env, timeout=240)
    pdf = Path(outdir) / (Path(docx).stem + ".pdf")
    shutil.rmtree(prof, ignore_errors=True)
    if not pdf.exists():
        raise SystemExit(f"แปลงเป็น PDF ไม่สำเร็จ: {r.stderr.strip()[:300]}")
    return pdf


def page_texts(pdf, n):
    out = []
    for i in range(1, n + 1):
        # โหมดลำดับการอ่าน (ไม่ใช้ -layout) เพื่อให้หัวข้อในหน้า 2 คอลัมน์อยู่ต้นบรรทัด
        out.append(run(["pdftotext", "-f", str(i), "-l", str(i), str(pdf), "-"]).stdout)
    return out


def font_report(pdf):
    r = run(["pdffonts", str(pdf)]).stdout.splitlines()[2:]
    names = sorted({ln.split()[0].split("+")[-1] for ln in r if ln.strip()})
    want = U.SPEC["font"].replace(" ", "")
    fc = run(["fc-match", U.SPEC["font"]]).stdout.strip()
    installed = U.SPEC["font"].lower().replace(" ", "") in fc.lower().replace(" ", "")
    embedded = any(want.lower() in n.lower().replace("-", "") for n in names)
    return {"embedded_fonts": names, "fc_match": fc, "th_sarabun_installed": installed, "th_sarabun_embedded": embedded}


def overflow(pdf, n):
    """หาคำที่ล้นขอบพื้นที่พิมพ์ (ใช้ขอบตั้งต้น JDTAIA)"""
    xml = run(["pdftotext", "-bbox", str(pdf), "-"]).stdout
    left = U.SPEC["left"] / U.CM * PT_PER_CM
    right_m = U.SPEC["right"] / U.CM * PT_PER_CM
    hits = []
    page = 0
    pw = None
    for line in xml.splitlines():
        m = re.search(r'<page width="([\d.]+)"', line)
        if m:
            page += 1
            pw = float(m.group(1))
            continue
        m = re.search(r'<word xMin="([\d.]+)" yMin="[\d.]+" xMax="([\d.]+)" yMax="[\d.]+">(.*)</word>', line)
        if m and pw:
            x0, x1, w = float(m.group(1)), float(m.group(2)), m.group(3)
            if x0 < left - 4 or x1 > pw - right_m + 4:
                hits.append({"page": page, "word": w[:40], "xMin": round(x0, 1), "xMax": round(x1, 1)})
    return hits


def main(argv=None):
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("docx")
    ap.add_argument("--outdir", required=True, help="โฟลเดอร์ทำงาน (เช่น scratchpad) ไม่ใช่โฟลเดอร์ส่งมอบ")
    ap.add_argument("--png", action="store_true", help="สร้างภาพทุกหน้าเพื่อตรวจด้วยตา")
    ap.add_argument("--dpi", type=int, default=110)
    ap.add_argument("--min-pages", type=int, default=8)
    ap.add_argument("--max-pages", type=int, default=10)
    ap.add_argument("--count", choices=("main", "main+refs", "all"), default="main+refs",
                    help="ขอบเขตที่นับเทียบเกณฑ์ (ตั้งตามคู่มือ/คำสั่งผู้ใช้)")
    ap.add_argument("--json")
    a = ap.parse_args(argv)
    out = Path(a.outdir)
    out.mkdir(parents=True, exist_ok=True)
    pdf = to_pdf(a.docx, out)
    info = run(["pdfinfo", str(pdf)]).stdout
    n = int(re.search(r"Pages:\s+(\d+)", info).group(1))
    texts = page_texts(pdf, n)
    ref_page = next((i + 1 for i, t in enumerate(texts) if re.search(r"(?m)^\s*(\d+\.?\s*)?(เอกสารอ้างอิง|References|REFERENCES)\s*$", t)), None)
    app_page = next((i + 1 for i, t in enumerate(texts) if re.search(r"(?m)^\s*(ภาคผนวก|Appendix|APPENDIX)\b", t) and (ref_page is None or i + 1 >= ref_page)), None)
    last_main = (app_page - 1) if app_page else n
    seg = {"total": n, "references_start_page": ref_page, "appendix_start_page": app_page,
           "main_pages(before refs heading page incl.)": ref_page or last_main,
           "main+refs_pages": last_main,
           "appendix_pages": (n - app_page + 1) if app_page else 0}
    counted = {"main": ref_page or last_main, "main+refs": last_main, "all": n}[a.count]
    within = a.min_pages <= counted <= a.max_pages
    fonts = font_report(pdf)
    ov = overflow(pdf, n)
    hidden = {name: sum(t.count(ch) for t in texts) for ch, name in U.HIDDEN_CHARS.items() if any(ch in t for t in texts)}
    pngs = []
    if a.png:
        prefix = out / Path(a.docx).stem
        run(["pdftoppm", "-png", "-r", str(a.dpi), str(pdf), str(prefix)])
        pngs = sorted(str(p) for p in out.glob(Path(a.docx).stem + "-*.png"))
    res = {"pdf": str(pdf), "pages": seg, "counted_scope": a.count, "counted_pages": counted,
           "limit": [a.min_pages, a.max_pages], "within_limit": within, "fonts": fonts,
           "overflow_words": ov[:50], "overflow_count": len(ov), "hidden_chars_in_pdf": hidden, "png": pngs}
    if a.json:
        Path(a.json).write_text(json.dumps(res, ensure_ascii=False, indent=2), encoding="utf8")
    print(f"PDF: {pdf}")
    print(f"หน้าทั้งหมด {n} | เอกสารอ้างอิงเริ่มหน้า {ref_page} | ภาคผนวกเริ่มหน้า {app_page} | ภาคผนวก {seg['appendix_pages']} หน้า")
    print(f"นับแบบ {a.count} = {counted} หน้า เทียบเกณฑ์ {a.min_pages}–{a.max_pages}: {'ผ่าน' if within else 'ไม่ผ่าน'}")
    if not fonts["th_sarabun_installed"] or not fonts["th_sarabun_embedded"]:
        print(f"⚠ ฟอนต์ TH Sarabun New ไม่ได้ใช้จริงในการเรนเดอร์ (fc-match: {fonts['fc_match']}; ฝังใน PDF: {fonts['embedded_fonts']})")
        print("  → จำนวนหน้า/การตัดคำยังไม่ใช่ค่ายืนยัน ต้องตรวจซ้ำบนเครื่องที่มีฟอนต์หรือใน Microsoft Word")
    if ov:
        print(f"⚠ คำล้นขอบพื้นที่พิมพ์ {len(ov)} จุด เช่น {ov[:5]}")
    if hidden:
        print(f"⚠ พบอักขระแฝงใน PDF: {hidden}")
    if pngs:
        print(f"ภาพตรวจหน้า {len(pngs)} ไฟล์ใน {out} (ใช้ตรวจภายใน ไม่ส่งมอบโดยอัตโนมัติ)")
    print("หมายเหตุ: เรนเดอร์ด้วย LibreOffice — ยืนยันผลใน Word เมื่อต้องการความแม่นยำของการแสดงผลใน Word")
    return 0 if within and not ov and not hidden else 2


if __name__ == "__main__":
    sys.exit(main())
