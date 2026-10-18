#!/usr/bin/env python3
"""ตรวจ (และเลือกเรียงใหม่) เลขอ้างอิง IEEE ในบทความ .docx

  ตรวจอย่างเดียว:  python ieee_refs.py article.docx [--json out.json]
  เรียงเลขใหม่:    python ieee_refs.py article.docx --renumber out.docx [--skip-appendix]

ตรวจ: ลำดับตามการอ้างครั้งแรก, อ้างแต่ไม่มีในรายการ, มีในรายการแต่ไม่ถูกอ้าง,
เลขรายการไม่ต่อเนื่อง, รายการซ้ำ, รูปแบบ DOI, รายการที่ไม่มีปี, วันที่เข้าถึงเว็บ

ข้อจำกัด (ต้องแจ้งผู้ใช้):
- ตรวจ "รูปแบบและความสอดคล้องของเลข" เท่านั้น ไม่ได้ยืนยันว่าแหล่งอ้างอิงมีจริง
  หรือ DOI ชี้ไปที่เอกสารนั้นจริง
- รองรับเลขในวงเล็บเหลี่ยม [1], [1, 3], [1–3], [1]–[3]
- รายการอ้างอิงที่ใช้ Word auto-numbering (ไม่มี [n] เป็นข้อความ) จะถูกนับตามลำดับย่อหน้า
- --renumber ไม่ลบรายการที่ไม่ถูกอ้าง แต่ย้ายไปท้ายและรายงาน
"""
from __future__ import annotations

import argparse
import json
import re
import sys
from collections import OrderedDict
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))
from docx.oxml.ns import qn  # noqa: E402

import docx_utils as U  # noqa: E402

# กลุ่มอ้างอิง: [1]–[3] (แยกวงเล็บ) หรือ [1, 2, 4–6] (วงเล็บเดียว)
SPLIT_RANGE_RE = re.compile(r"\[(\d{1,3})\]\s*([–\-–])\s*\[(\d{1,3})\]")
GROUP_RE = re.compile(r"\[(\d{1,3}(?:\s*(?:,|[–\-–])\s*\d{1,3})*)\]")
ENTRY_RE = re.compile(r"^\s*\[(\d{1,3})\]\s*")
DOI_RE = re.compile(r"\b(?:doi:\s*|https?://(?:dx\.)?doi\.org/)?(10\.\d{4,9}/[^\s,;]+)", re.I)
YEAR_RE = re.compile(r"\b(19|20)\d{2}\b|\b25[4-7]\d\b")
ACCESS_RE = re.compile(r"(Accessed[^.\]]*|สืบค้นเมื่อ[^.\]]*|เข้าถึงเมื่อ[^.\]]*)", re.I)


def expand(body: str):
    nums = []
    for part in re.split(r"\s*,\s*", body.strip()):
        m = re.match(r"(\d+)\s*[–\-–]\s*(\d+)$", part)
        if m:
            a, b = int(m.group(1)), int(m.group(2))
            nums.extend(range(a, b + 1) if a <= b else [a, b])
        elif part.isdigit():
            nums.append(int(part))
    return nums


def find_citations(text):
    """คืนรายการ (start, end, [nums], style) ไม่ทับกัน"""
    out = []
    taken = []
    for m in SPLIT_RANGE_RE.finditer(text):
        a, b = int(m.group(1)), int(m.group(3))
        out.append((m.start(), m.end(), list(range(a, b + 1)) if a <= b else [a, b], "split", m.group(2)))
        taken.append((m.start(), m.end()))
    for m in GROUP_RE.finditer(text):
        if any(s <= m.start() < e for s, e in taken):
            continue
        out.append((m.start(), m.end(), expand(m.group(1)), "group", "–" if re.search(r"[–\-–]", m.group(1)) else ""))
    return sorted(out)


def compress(nums):
    nums = sorted(set(nums))
    runs, start, prev = [], None, None
    for n in nums:
        if start is None:
            start = prev = n
        elif n == prev + 1:
            prev = n
        else:
            runs.append((start, prev))
            start = prev = n
    if start is not None:
        runs.append((start, prev))
    return runs


def render_group(nums, style, dash):
    runs = compress(nums)
    use_range = bool(dash)
    if style == "split":
        if len(runs) == 1 and runs[0][1] - runs[0][0] >= 1:
            return f"[{runs[0][0]}]{dash or '–'}[{runs[0][1]}]"
        return ", ".join(f"[{n}]" for n in sorted(set(nums)))
    parts = []
    for a, b in runs:
        if a == b:
            parts.append(str(a))
        elif use_range and b - a >= 2:
            parts.append(f"{a}{dash}{b}")
        else:
            parts.extend(str(x) for x in range(a, b + 1))
    return "[" + ", ".join(parts) + "]"


def scan(doc):
    paras = U.all_paragraphs(doc)
    region = "body"
    body_cites, appx_cites = [], []
    entries = []  # (index_in_paras, number or None, text)
    for idx, (loc, p, in_table) in enumerate(paras):
        text = U.para_text(p)
        t = text.strip()
        if not in_table and U.REF_HEADING_RE.match(t):
            region = "refs"
            continue
        if not in_table and U.APPENDIX_HEADING_RE.match(t) and len(t) <= 80:
            region = "appendix"
            continue
        if region == "appendix" and ENTRY_RE.match(t):
            continue  # รายการอ้างอิงเพิ่มเติมของภาคผนวก: ไม่นับเป็นการอ้าง และไม่รวมในรายการหลัก
        if region == "refs":
            if not t:
                continue
            m = ENTRY_RE.match(t)
            auto = p._p.pPr is not None and p._p.pPr.find(qn("w:numPr")) is not None
            entries.append({"idx": idx, "num": int(m.group(1)) if m else None, "auto_numbered": auto and not m, "text": t})
            continue
        for s, e, nums, style, dash in find_citations(text):
            (body_cites if region == "body" else appx_cites).append({"idx": idx, "loc": loc, "nums": nums, "raw": text[s:e]})
    # auto-numbered: กำหนดเลขตามลำดับ
    if entries and all(en["num"] is None for en in entries):
        for k, en in enumerate(entries, 1):
            en["num"] = k
    return paras, body_cites, appx_cites, entries


def analyse(doc):
    paras, body, appx, entries = scan(doc)
    first = list(OrderedDict.fromkeys(n for c in body for n in c["nums"]))
    first_appx = [n for n in OrderedDict.fromkeys(n for c in appx for n in c["nums"]) if n not in first]
    listed = [e["num"] for e in entries if e["num"] is not None]
    rep = {"citations_body": len(body), "citations_appendix": len(appx), "entries": len(entries),
           "first_appearance_order": first, "appendix_only_cited": first_appx}
    problems = []
    if not entries:
        problems.append(("error", "ไม่พบหัวข้อเอกสารอ้างอิงหรือรายการ [n]"))
    if listed and listed != list(range(1, len(listed) + 1)):
        problems.append(("error", f"เลขรายการท้ายเรื่องไม่เรียง 1..n: {listed}"))
    dup_nums = sorted({n for n in listed if listed.count(n) > 1})
    if dup_nums:
        problems.append(("error", f"เลขรายการซ้ำ: {dup_nums}"))
    cited = set(first) | set(first_appx)
    missing = sorted(cited - set(listed))
    unused = sorted(set(listed) - cited)
    if missing:
        problems.append(("error", f"อ้างในเนื้อหาแต่ไม่มีในรายการ: {missing}"))
    if unused:
        problems.append(("warn", f"มีในรายการแต่ไม่ถูกอ้าง: {unused}"))
    if first and first != list(range(1, len(first) + 1)):
        problems.append(("error", f"ลำดับการอ้างครั้งแรกไม่ใช่ 1,2,3,...: {first[:30]}"))
    # รายการซ้ำ (ข้อความหรือ DOI)
    seen_txt, seen_doi = {}, {}
    for e in entries:
        body_t = re.sub(r"\W+", "", ENTRY_RE.sub("", e["text"]).lower())
        if body_t in seen_txt:
            problems.append(("warn", f"รายการ [{e['num']}] อาจซ้ำกับ [{seen_txt[body_t]}]"))
        seen_txt.setdefault(body_t, e["num"])
        for d in DOI_RE.findall(e["text"]):
            d = d.rstrip(".").lower()
            if d in seen_doi:
                problems.append(("warn", f"DOI ซ้ำ [{seen_doi[d]}] และ [{e['num']}]: {d}"))
            seen_doi.setdefault(d, e["num"])
        if "doi" in e["text"].lower() and not DOI_RE.search(e["text"]):
            problems.append(("warn", f"[{e['num']}] มีคำว่า doi แต่รูปแบบไม่ใช่ 10.xxxx/…"))
        if not YEAR_RE.search(e["text"]):
            problems.append(("info", f"[{e['num']}] ไม่พบปีพิมพ์"))
        acc = ACCESS_RE.findall(e["text"])
        if acc:
            problems.append(("info", f"[{e['num']}] วันที่เข้าถึง “{acc[0].strip()}” — ต้องเป็นวันที่ตรวจจริง"))
        if re.search(r"https?://", e["text"]) and not acc:
            problems.append(("info", f"[{e['num']}] มี URL แต่ไม่พบวันที่เข้าถึง"))
    rep["problems"] = problems
    return rep, paras, body, appx, entries


def replace_in_paragraph(p, func):
    """แทนข้อความใน paragraph ข้าม run โดยคงรูปแบบ run แรกของแต่ละจุด คืนจำนวนจุดที่แก้"""
    ts = [t for t in p._p.iter(qn("w:t"))]
    if not ts:
        return 0
    full = "".join(t.text or "" for t in ts)
    hits = func(full)
    if not hits:
        return 0
    bounds, pos = [], 0
    for t in ts:
        n = len(t.text or "")
        bounds.append((pos, pos + n))
        pos += n
    for s, e, new in sorted(hits, reverse=True):
        first = True
        for t, (a, b) in zip(ts, bounds):
            if b <= s or a >= e:
                continue
            txt = t.text or ""
            ls, le = max(s, a) - a, min(e, b) - a
            t.text = txt[:ls] + (new if first else "") + txt[le:]
            first = False
            t.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
    return len(hits)


def renumber(path, out, skip_appendix=False):
    doc = U.load(path)
    rep, paras, body, appx, entries = analyse(doc)
    if any(e["auto_numbered"] for e in entries):
        raise SystemExit("รายการอ้างอิงใช้ auto-numbering ของ Word — ไม่เรียงใหม่อัตโนมัติ ให้แจ้งผู้ใช้")
    order = rep["first_appearance_order"] + rep["appendix_only_cited"]
    listed = [e["num"] for e in entries]
    unused = [n for n in listed if n not in order]
    order = [n for n in order if n in listed] + unused
    mapping = {old: new for new, old in enumerate(order, 1)}
    cite_idx = {c["idx"] for c in body} | (set() if skip_appendix else {c["idx"] for c in appx})
    changed = 0
    for idx in sorted(cite_idx):
        p = paras[idx][1]

        def f(text):
            res = []
            for s, e, nums, style, dash in find_citations(text):
                if not all(n in mapping for n in nums):
                    continue
                new = render_group([mapping[n] for n in nums], style, dash)
                if new != text[s:e]:
                    res.append((s, e, new))
            return res
        changed += replace_in_paragraph(p, f)
    # เรียงรายการท้ายเรื่อง
    els = [paras[e["idx"]][1]._p for e in entries]
    parent = els[0].getparent()
    anchor = els[0].getprevious()
    by_old = {e["num"]: el for e, el in zip(entries, els)}
    for el in els:
        parent.remove(el)
    from docx.text.paragraph import Paragraph
    for old in sorted(by_old, key=lambda o: mapping[o]):
        el = by_old[old]
        if anchor is None:
            parent.insert(0, el)
        else:
            anchor.addnext(el)
        anchor = el
        new_label = f"[{mapping[old]}]"
        replace_in_paragraph(Paragraph(el, doc), lambda t, nl=new_label: [(m.start(1) - 1, m.end(1) + 1, nl)] if (m := ENTRY_RE.match(t)) and t[m.start(1) - 1:m.end(1) + 1] != nl else [])
    doc.save(out)
    return mapping, unused, changed


def main(argv=None):
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("docx")
    ap.add_argument("--renumber", metavar="OUT.docx", help="บันทึกไฟล์ใหม่ที่เรียงเลขตามการอ้างครั้งแรก")
    ap.add_argument("--skip-appendix", action="store_true", help="ไม่แก้เลขอ้างในภาคผนวก")
    ap.add_argument("--json")
    a = ap.parse_args(argv)
    doc = U.load(a.docx)
    rep, *_ = analyse(doc)
    print(f"อ้างในเนื้อหา {rep['citations_body']} จุด, ในภาคผนวก {rep['citations_appendix']} จุด, รายการท้ายเรื่อง {rep['entries']} รายการ")
    print("ลำดับอ้างครั้งแรก:", rep["first_appearance_order"])
    for lvl, msg in rep["problems"]:
        print(f"  [{lvl}] {msg}")
    if not rep["problems"]:
        print("  ไม่พบปัญหาความสอดคล้องของเลข (ยังไม่ได้ตรวจความถูกต้องของแหล่งข้อมูล)")
    if a.renumber:
        if Path(a.renumber).resolve() == Path(a.docx).resolve():
            raise SystemExit("ห้ามเขียนทับต้นฉบับ — ระบุไฟล์ผลลัพธ์ใหม่")
        mapping, unused, changed = renumber(a.docx, a.renumber, a.skip_appendix)
        moved = {o: n for o, n in mapping.items() if o != n}
        print(f"บันทึก {a.renumber}: แก้กลุ่มอ้างอิง {changed} จุด, เปลี่ยนเลข {moved}")
        if unused:
            print(f"  รายการไม่ถูกอ้าง (ย้ายไปท้าย ไม่ลบ): เดิม {unused}")
        rep["renumber"] = {"mapping": mapping, "unused": unused, "changed_groups": changed}
    if a.json:
        Path(a.json).write_text(json.dumps(rep, ensure_ascii=False, indent=2, default=str), encoding="utf8")
    return 1 if any(l == "error" for l, _ in rep["problems"]) and not a.renumber else 0


if __name__ == "__main__":
    sys.exit(main())
