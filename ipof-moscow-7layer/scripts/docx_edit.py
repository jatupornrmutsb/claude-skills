#!/usr/bin/env python3
"""แก้ไฟล์ .docx เดิมโดยคงรูปแบบ (ฟอนต์ ขนาด การขึ้นบรรทัดด้วยมือ ตาราง) แล้วบันทึกเป็นรุ่นใหม่

  python docx_edit.py ไฟล์ล่าสุด_V2.docx edits.json [--out ชื่อ.docx] [--dry-run]

edits.json = รายการคำสั่ง ทำตามลำดับ:
  {"find": "ข้อความเดิม", "replace": "ข้อความใหม่", "count": 1}      แทนข้อความข้าม run (count = จำนวนที่คาดว่าพบ)
  {"table": 3, "row": 2, "col": 4, "value": "4.20"}                ตั้งค่าเซลล์ (นับจาก 1) คงรูปแบบ run แรก
  {"paragraph_contains": "…", "new_text": "…"}                      เขียนย่อหน้าใหม่ทั้งย่อหน้า (รายงานว่าต้องตัดบรรทัดเอง)

กติกา
- ไม่เขียนทับต้นฉบับ ชื่อผลลัพธ์ตั้งเป็นรุ่นถัดไปอัตโนมัติ (_V2 → _V3; ไม่มีเลขรุ่น → _V1)
- แทนข้อความที่คร่อมจุดขึ้นบรรทัดด้วยมือ (<w:br/>) ไม่ได้ → รายงานให้แก้ทีละส่วน เพื่อไม่ทำลายการตัดบรรทัดเดิม
- จำนวนที่พบไม่ตรง count → ไม่บันทึก (ป้องกันแก้ผิดจุด)
"""
from __future__ import annotations

import argparse
import json
import re
import sys

for _s in (sys.stdout, sys.stderr):  # Windows console ไม่ใช่ UTF-8 จะ print ภาษาไทยไม่ได้
    if hasattr(_s, "reconfigure"):
        _s.reconfigure(encoding="utf-8", errors="replace")
from pathlib import Path

from docx import Document
from docx.oxml.ns import qn


def next_version(path: Path) -> Path:
    m = re.search(r"_V(\d+)$", path.stem)
    if m:
        stem = path.stem[: m.start()] + f"_V{int(m.group(1)) + 1}"
    else:
        stem = path.stem + "_V1"
    return path.with_name(stem + path.suffix)


def all_paragraphs(doc):
    def walk_tbl(tbl):
        for row in tbl.findall(qn("w:tr")):
            for tc in row.findall(qn("w:tc")):
                for el in tc.iterchildren():
                    if el.tag == qn("w:p"):
                        yield el
                    elif el.tag == qn("w:tbl"):
                        yield from walk_tbl(el)
    for el in doc.element.body.iterchildren():
        if el.tag == qn("w:p"):
            yield el
        elif el.tag == qn("w:tbl"):
            yield from walk_tbl(el)
    for sec in doc.sections:
        for part in (sec.header, sec.footer):
            if part is not None and not part.is_linked_to_previous:
                yield from part._element.iter(qn("w:p"))


def segments(p):
    """แบ่ง w:t ของย่อหน้าเป็นช่วงที่ไม่มี w:br คั่น: [[t,...], ...]"""
    segs, cur = [], []
    for el in p.iter():
        if el.tag == qn("w:t"):
            cur.append(el)
        elif el.tag in (qn("w:br"), qn("w:cr")) or (el.tag == qn("w:tab") and el.getparent().tag == qn("w:r")):
            if cur:
                segs.append(cur)
            cur = []
    if cur:
        segs.append(cur)
    return segs


def replace_in_seg(ts, find, repl):
    full = "".join(t.text or "" for t in ts)
    hits = [m.start() for m in re.finditer(re.escape(find), full)]
    if not hits:
        return 0
    bounds, pos = [], 0
    for t in ts:
        n = len(t.text or "")
        bounds.append((pos, pos + n))
        pos += n
    for s in reversed(hits):
        e = s + len(find)
        first = True
        for t, (a, b) in zip(ts, bounds):
            if b <= s or a >= e:
                continue
            txt = t.text or ""
            t.text = txt[: max(s, a) - a] + (repl if first else "") + txt[min(e, b) - a:]
            t.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
            first = False
    return len(hits)


def ptext(p):
    out = []
    for el in p.iter():
        if el.tag == qn("w:t"):
            out.append(el.text or "")
        elif el.tag in (qn("w:br"), qn("w:cr")):
            out.append("\n")
    return "".join(out)


def set_cell(doc, ti, ri, ci, value):
    tbls = doc.element.body.findall(".//" + qn("w:tbl"))
    tbl = tbls[ti - 1]
    tr = tbl.findall(qn("w:tr"))[ri - 1]
    tc = tr.findall(qn("w:tc"))[ci - 1]
    ts = list(tc.iter(qn("w:t")))
    old = "".join(t.text or "" for t in ts)
    if not ts:
        raise SystemExit(f"ตาราง {ti} แถว {ri} คอลัมน์ {ci} ไม่มี run ให้คงรูปแบบ — แก้ด้วยมือ")
    ts[0].text = value
    for t in ts[1:]:
        t.text = ""
    return old


def main(argv=None):
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("src")
    ap.add_argument("edits")
    ap.add_argument("--out")
    ap.add_argument("--dry-run", action="store_true")
    a = ap.parse_args(argv)
    src = Path(a.src)
    out = Path(a.out) if a.out else next_version(src)
    if out.resolve() == src.resolve():
        raise SystemExit("ห้ามเขียนทับต้นฉบับ")
    doc = Document(str(src))
    edits = json.loads(Path(a.edits).read_text(encoding="utf-8"))
    problems, report, rewritten = [], [], []
    for i, e in enumerate(edits, 1):
        if "find" in e:
            n = 0
            for p in all_paragraphs(doc):
                for seg in segments(p):
                    n += replace_in_seg(seg, e["find"], e["replace"])
            if n == 0:
                crossing = [ptext(p) for p in all_paragraphs(doc) if e["find"] in ptext(p).replace("\n", "")]
                problems.append(f"#{i} ไม่พบ “{e['find'][:40]}”" + (" (คร่อมจุดขึ้นบรรทัดด้วยมือ — แยกคำสั่งเป็นช่วง)" if crossing else ""))
            elif "count" in e and n != e["count"]:
                problems.append(f"#{i} พบ “{e['find'][:40]}” {n} จุด แต่คาด {e['count']}")
            report.append(f"#{i} แทน “{e['find'][:30]}” → “{e['replace'][:30]}” {n} จุด")
        elif "table" in e:
            old = set_cell(doc, e["table"], e["row"], e["col"], str(e["value"]))
            report.append(f"#{i} ตาราง {e['table']} R{e['row']}C{e['col']}: “{old[:30]}” → “{e['value']}”")
        elif "paragraph_contains" in e:
            hits = [p for p in all_paragraphs(doc) if e["paragraph_contains"] in ptext(p)]
            if len(hits) != 1:
                problems.append(f"#{i} พบย่อหน้าที่มี “{e['paragraph_contains'][:30]}” {len(hits)} ย่อหน้า (ต้อง 1)")
                continue
            p = hits[0]
            ts = list(p.iter(qn("w:t")))
            for br in list(p.iter(qn("w:br"))):
                br.getparent().remove(br)
            ts[0].text = e["new_text"]
            for t in ts[1:]:
                t.text = ""
            rewritten.append(e["new_text"][:60])
            report.append(f"#{i} เขียนย่อหน้าใหม่: “{e['new_text'][:40]}…”")
    for r in report:
        print(r)
    if problems:
        print("ไม่บันทึก — ต้องแก้คำสั่งก่อน:")
        for p in problems:
            print("  ⚠", p)
        return 1
    if a.dry_run:
        print("(dry-run ไม่บันทึก)")
        return 0
    doc.save(str(out))
    print(f"บันทึก: {out}")
    if rewritten:
        print("ย่อหน้าที่เขียนใหม่ — ต้องจัดการตัดบรรทัดเองใน Word:")
        for t in rewritten:
            print("  •", t)
    print("แนะนำให้บันทึกไฟล์นี้ไว้ในโฟลเดอร์โปรเจกต์ของคุณ")
    return 0


if __name__ == "__main__":
    sys.exit(main())
