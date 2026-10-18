#!/usr/bin/env python3
"""เปรียบเทียบข้อความทุกย่อหน้าและทุกช่องตาราง ระหว่างต้นฉบับกับไฟล์ผลลัพธ์

ใช้สำหรับยืนยันงาน "จัดรูปแบบอย่างเดียว" ว่าถ้อยคำและตัวเลขไม่เปลี่ยน

  python text_diff.py original.docx formatted.docx [--json out.json]

จัดกลุ่มความต่าง:
  BREAK_ONLY     ต่างเฉพาะจุดขึ้นบรรทัด (<w:br/>) — ไม่ใช่การเปลี่ยนถ้อยคำ
  BREAK_SPACE    ขึ้นบรรทัดถูกแทนด้วยช่องว่างหรือกลับกัน — ตรวจภาษาไทยว่าไม่เกิดวรรคใหม่
  WHITESPACE     ต่างเฉพาะช่องว่าง/แท็บซ้ำ
  HIDDEN_ONLY    ต่างเฉพาะอักขระแฝง (เช่น U+200B)
  RENUMBER       ต่างเฉพาะเลขใน [ ] หรือเลขหัวข้อ/ตาราง/รูป — รายงานแยกตามที่ได้รับอนุญาต
  EMPTY_PARA     เพิ่ม/ลบย่อหน้าว่าง
  SPLIT_MERGE    ย่อหน้าถูกแยก/รวม แต่ข้อความรวมเท่าเดิม
  REORDER        ย่อหน้าชุดเดิมถูกสลับลำดับ (เช่น เรียงรายการอ้างอิงใหม่) — ไม่ผ่านถ้าไม่ใส่ --allow-renumber
  WORDING        ถ้อยคำหรือตัวเลขเปลี่ยน  → งานจัดรูปแบบอย่างเดียว "ไม่ผ่าน"
ออกด้วยรหัส 1 เมื่อมี WORDING (หรือ RENUMBER เมื่อไม่ใส่ --allow-renumber)
"""
from __future__ import annotations

import argparse
import difflib
import json
import re
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))
import docx_utils as U  # noqa: E402

NUM_RE = re.compile(r"\d+(?:[.,]\d+)*")


def units(path):
    """ลำดับหน่วยข้อความ: ย่อหน้านอกตาราง และช่องตาราง (ทุกย่อหน้าในช่องรวมด้วย  )"""
    doc = U.load(path)
    out = []
    ti = 0
    for kind, obj in U.body_items(doc):
        if kind == "p":
            out.append(("P", U.normalize(U.para_text(obj))))
        else:
            ti += 1
            out.extend(table_units(obj, f"ตาราง#{ti}"))
    return out


def table_units(t, prefix):
    out = []
    seen = set()
    for r, row in enumerate(t.rows):
        for c, cell in enumerate(row.cells):
            if cell._tc in seen:
                continue
            seen.add(cell._tc)
            txt = " ".join(U.para_text(p) for p in cell.paragraphs)
            out.append((f"{prefix} R{r + 1}C{c + 1}", U.normalize(txt)))
            for k, inner in enumerate(cell.tables):
                out.extend(table_units(inner, f"{prefix} R{r + 1}C{c + 1}/ใน#{k + 1}"))
    return out


def strip_hidden(s):
    return "".join(ch for ch in s if ch not in U.HIDDEN_CHARS)


def classify(a, b):
    if a == b:
        return None
    if a.replace("\n", "") == b.replace("\n", ""):
        return "BREAK_ONLY"
    if strip_hidden(a) == strip_hidden(b):
        return "HIDDEN_ONLY"
    sq = lambda s: re.sub(r"[ \t ]+", " ", s).strip()  # noqa: E731
    if sq(a) == sq(b):
        return "WHITESPACE"
    if sq(strip_hidden(a).replace("\n", " ")) == sq(strip_hidden(b).replace("\n", " ")):
        return "BREAK_SPACE"
    renum = lambda s: re.sub(r"\[[\d,\s–\-–]+\]", "[#]", re.sub(r"^(\s*)(\d+(\.\d+)*)", r"\1#", re.sub(r"(ตารางที่|รูปที่|สมการที่)\s*\(?\d+\)?", r"\1 #", s)))  # noqa: E731
    if renum(sq(a)) == renum(sq(b)):
        return "RENUMBER"
    return "WORDING"


def norm_renum(s):
    s = re.sub(r"[ \t\u00a0]+", " ", strip_hidden(s).replace("\n", " ")).strip()
    s = re.sub(r"(ตารางที่|รูปที่|สมการที่)\s*\(?\d+\)?", r"\1 #", s)
    s = re.sub(r"^(\s*)(\d+(\.\d+)*)", r"\1#", s)
    return re.sub(r"\[[\d,\s\u2013\-–]+\]", "[#]", s)


def short(s, n=70):
    s = s.replace("\n", "⏎").replace(" ", "¶")
    return s if len(s) <= n else s[:n] + "…"


def diff(orig, new):
    A, B = units(orig), units(new)
    ka = [t for _, t in A]
    kb = [t for _, t in B]
    sm = difflib.SequenceMatcher(a=[strip_hidden(x).replace("\n", "") for x in ka],
                                 b=[strip_hidden(x).replace("\n", "") for x in kb], autojunk=False)
    changes = []
    for op, i1, i2, j1, j2 in sm.get_opcodes():
        if op == "equal":
            for i, j in zip(range(i1, i2), range(j1, j2)):
                c = classify(ka[i], kb[j])
                if c:
                    changes.append({"type": c, "where": A[i][0] + f"#{i}", "old": short(ka[i]), "new": short(kb[j])})
            continue
        olds, news = ka[i1:i2], kb[j1:j2]
        non_empty_old = [x for x in olds if x.strip()]
        non_empty_new = [x for x in news if x.strip()]
        if not non_empty_old and not non_empty_new:
            changes.append({"type": "EMPTY_PARA", "where": f"#{i1}", "old": f"{len(olds)} ย่อหน้าว่าง", "new": f"{len(news)} ย่อหน้าว่าง"})
            continue
        if op == "replace" and sorted(norm_renum(x) for x in olds) == sorted(norm_renum(y) for y in news) \
                and [norm_renum(x) for x in olds] != [norm_renum(y) for y in news]:
            changes.append({"type": "REORDER", "where": f"#{i1}-{i2}", "old": f"{len(olds)} หน่วยถูกสลับลำดับ",
                            "new": short(" | ".join(news[:3]))})
            if any(classify(x, y) == "RENUMBER" for x, y in zip(sorted(olds, key=norm_renum), sorted(news, key=norm_renum))):
                changes.append({"type": "RENUMBER", "where": f"#{i1}-{i2}", "old": "เลขรายการ", "new": "เลขรายการใหม่"})
            continue
        if op == "replace" and len(olds) == len(news):
            for i, (x, y) in enumerate(zip(olds, news)):
                c = classify(x, y)
                if c:
                    changes.append({"type": c, "where": A[i1 + i][0] + f"#{i1 + i}", "old": short(x), "new": short(y)})
            continue
        ja = "".join(non_empty_old).replace("\n", "")
        jb = "".join(non_empty_new).replace("\n", "")
        if strip_hidden(ja) == strip_hidden(jb) and (non_empty_old and non_empty_new):
            changes.append({"type": "SPLIT_MERGE", "where": f"#{i1}-{i2}", "old": short(" | ".join(non_empty_old)), "new": short(" | ".join(non_empty_new))})
            if len(olds) - len(non_empty_old) != len(news) - len(non_empty_new):
                changes.append({"type": "EMPTY_PARA", "where": f"#{i1}", "old": "", "new": ""})
            continue
        changes.append({"type": "WORDING", "where": (A[i1][0] if i1 < len(A) else "ท้ายเอกสาร") + f"#{i1}",
                        "old": short(" | ".join(olds)) or "(ไม่มี)", "new": short(" | ".join(news)) or "(ไม่มี)"})
    # ถ้าทั้งเอกสารมีหน่วยข้อความชุดเดิมทุกหน่วย (ไม่นับเลขอ้างอิง/หัวข้อ) ความต่างที่เหลือคือการย้ายลำดับ
    from collections import Counter as _C
    if any(c["type"] == "WORDING" for c in changes) and \
            _C(norm_renum(x) for x in ka if x.strip()) == _C(norm_renum(y) for y in kb if y.strip()):
        for c in changes:
            if c["type"] == "WORDING":
                c["type"] = "REORDER"
    # เทียบตัวเลขทั้งเอกสารแบบนับจำนวน (จับตัวเลขหายหรือเพิ่ม)
    na = sorted(NUM_RE.findall(" ".join(ka)))
    nb = sorted(NUM_RE.findall(" ".join(kb)))
    num_note = None
    if na != nb:
        from collections import Counter
        ca, cb = Counter(na), Counter(nb)
        num_note = {"missing": dict((ca - cb).most_common(20)), "added": dict((cb - ca).most_common(20))}
    stats = {"units_original": len(A), "units_new": len(B),
             "table_cells_original": sum(1 for w, _ in A if w != "P"),
             "table_cells_new": sum(1 for w, _ in B if w != "P")}
    return changes, num_note, stats


def main(argv=None):
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("original")
    ap.add_argument("new")
    ap.add_argument("--allow-renumber", action="store_true", help="ผู้ใช้อนุญาตให้เปลี่ยนเลขหัวข้อ/เลขอ้างอิง")
    ap.add_argument("--json")
    a = ap.parse_args(argv)
    changes, num_note, stats = diff(a.original, a.new)
    from collections import Counter
    by = Counter(c["type"] for c in changes)
    print("สถิติ:", json.dumps(stats, ensure_ascii=False))
    print("สรุปความต่าง:", dict(by) if by else "ไม่มี (ข้อความตรงกันทุกหน่วย)")
    for c in changes[:80]:
        print(f"  [{c['type']}] {c['where']}\n     เดิม: {c['old']}\n     ใหม่: {c['new']}")
    if len(changes) > 80:
        print(f"  ... อีก {len(changes) - 80} รายการ")
    if num_note:
        print("ตัวเลขในเอกสารไม่ตรงกัน:", json.dumps(num_note, ensure_ascii=False))
    if a.json:
        Path(a.json).write_text(json.dumps({"stats": stats, "changes": changes, "numbers": num_note}, ensure_ascii=False, indent=2), encoding="utf8")
    fail = by["WORDING"] > 0 or ((by["RENUMBER"] or by["REORDER"]) and not a.allow_renumber) or stats["table_cells_original"] != stats["table_cells_new"]
    if num_note and not (a.allow_renumber and by["RENUMBER"]):
        fail = True
    print("ผล:", "ไม่ผ่าน — มีการเปลี่ยนเนื้อหา/ตัวเลข/โครงตาราง" if fail else "ผ่าน — ไม่มีการเปลี่ยนถ้อยคำหรือตัวเลข")
    return 1 if fail else 0


if __name__ == "__main__":
    sys.exit(main())
