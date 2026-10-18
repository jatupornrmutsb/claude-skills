#!/usr/bin/env python3
"""อ่านคำตอบจากแบบสอบถาม/แบบประเมิน .docx (ช่อง ☑/☐) เป็นไฟล์ CSV

  python read_forms.py OUTDIR form1.docx form2.docx ...   (หรือโฟลเดอร์ที่มี .docx)
       [--respondent-from filename|R] [--codes "P|N|FR|NFR|A|E"]

รองรับ
1) ตารางที่หัวคอลัมน์เป็นระดับคะแนน (5 4 3 2 1 หรือ มากที่สุด…น้อยที่สุด) หรือ เห็นชอบ/ไม่เห็นชอบ
   หรือ M S C W — ช่องที่ติ๊กในแถวนั้นคือคำตอบของรหัสในคอลัมน์แรก
2) บรรทัด "คะแนนความเหมาะสม ☐ 5 ☑ 4 ☐ 3 ☐ 2 ☐ 1" — ใช้รหัสล่าสุดที่พบก่อนหน้า (เช่น FR01)
3) ช่องติ๊กอื่นๆ (ข้อมูลทั่วไป ฟังก์ชันที่ควรพัฒนาก่อน ความเห็นโดยรวม) → choices.csv พร้อมคำถามกำกับ
4) คำตอบปลายเปิดหลังรหัส OPxx หรือคอลัมน์ "เหตุผล/ข้อเสนอแนะ" → text.csv
รูปช่องติ๊กที่รู้จัก: ☑ ☒ ✓ ✔ ■ ▣ / ☐ □, w:sym (Wingdings/Wingdings 2) และ checkbox content control

ผลลัพธ์ใน OUTDIR: scores.csv, choices.csv, text.csv, issues.csv
issues.csv = แถวที่ไม่ติ๊ก หรือติ๊กมากกว่า 1 ช่องในคำถามเลือกตอบเดียว — ต้องตรวจกับต้นฉบับ ห้ามเดาค่า
"""
from __future__ import annotations

import argparse
import csv
import re
import sys

for _s in (sys.stdout, sys.stderr):  # Windows console ไม่ใช่ UTF-8 จะ print ภาษาไทยไม่ได้
    if hasattr(_s, "reconfigure"):
        _s.reconfigure(encoding="utf-8", errors="replace")
from pathlib import Path

from docx import Document
from docx.oxml.ns import qn

CHECKED_CH = set("☑☒✓✔■▣☒")
EMPTY_CH = set("☐□")
# Wingdings / Wingdings 2 code points
SYM_CHECKED = {"F0FE", "F0FD", "F0FC", "F078", "F0FB", "F052", "F054", "F050", "F04F"}
SYM_EMPTY = {"F0A8", "F06F", "F071", "F0A3", "F0A1", "F030", "F0A2"}
W14 = "http://schemas.microsoft.com/office/word/2010/wordml"

LEVEL_WORDS = {"มากที่สุด": 5, "มาก": 4, "ปานกลาง": 3, "น้อย": 2, "น้อยที่สุด": 1}
AGREE_WORDS = {"เห็นชอบ": "agree", "ไม่เห็นชอบ": "disagree", "เห็นด้วย": "agree", "ไม่เห็นด้วย": "disagree"}
MOSCOW_WORDS = {"M": "M", "S": "S", "C": "C", "W": "W", "Must": "M", "Should": "S", "Could": "C", "Won't": "W", "Won’t": "W"}


def tokens(el):
    """ลำดับ token ของ paragraph/cell: ('box', True/False) หรือ ('text', s)"""
    out = []
    for node in el.iter():
        tag = node.tag
        if tag == qn("w:t") and node.text:
            buf = ""
            for ch in node.text:
                if ch in CHECKED_CH or ch in EMPTY_CH:
                    if buf:
                        out.append(("text", buf))
                        buf = ""
                    out.append(("box", ch in CHECKED_CH))
                else:
                    buf += ch
            if buf:
                out.append(("text", buf))
        elif tag == qn("w:sym"):
            c = (node.get(qn("w:char")) or "").upper()
            if c in SYM_CHECKED:
                out.append(("box", True))
            elif c in SYM_EMPTY:
                out.append(("box", False))
        elif tag == f"{{{W14}}}checkbox":
            chk = node.find(f"{{{W14}}}checked")
            out.append(("box", chk is not None and chk.get(f"{{{W14}}}val") in ("1", "true")))
            # อักขระแสดงผลของ content control จะถูกข้าม: ทำเครื่องหมายไว้
            out.append(("skipnext", None))
        elif tag in (qn("w:br"), qn("w:cr")):
            out.append(("text", "\n"))
    # ลบอักขระแสดงผลของ content control ที่ซ้ำกับ checkbox
    res, skip = [], False
    for t in out:
        if t[0] == "skipnext":
            skip = True
            continue
        if skip and t[0] == "box":
            skip = False
            continue
        if skip and t[0] == "text":
            skip = False
        res.append(t)
    return res


def plain(el):
    return "".join(s for k, s in tokens(el) if k == "text").strip()


def labeled_boxes(toks):
    """[(checked, label)] โดย label = ข้อความหลังช่องจนถึงช่องถัดไป"""
    out = []
    for i, (k, v) in enumerate(toks):
        if k != "box":
            continue
        lab = ""
        for k2, v2 in toks[i + 1:]:
            if k2 == "box":
                break
            lab += v2
        out.append((v, lab.strip()))
    return out


def header_kind(cells_text):
    vals = {}
    for j, t in enumerate(cells_text):
        t = t.strip()
        if re.fullmatch(r"[1-5]", t):
            vals[j] = int(t)
        elif t in LEVEL_WORDS:
            vals[j] = LEVEL_WORDS[t]
        elif t in AGREE_WORDS:
            vals[j] = AGREE_WORDS[t]
        elif t in MOSCOW_WORDS:
            vals[j] = MOSCOW_WORDS[t]
    if len(vals) >= 2:
        return vals
    return None


def value_of_label(label):
    label = label.strip()
    m = re.match(r"^([1-5])\b", label)
    if m:
        return m.group(1)
    for w, v in sorted(AGREE_WORDS.items(), key=lambda x: -len(x[0])):
        if label.startswith(w):
            return v
    for w, v in MOSCOW_WORDS.items():
        if re.match(rf"^{re.escape(w)}\b", label):
            return v
    return None


def parse(path, respondent, code_re, out):
    doc = Document(path)
    body = doc.element.body
    last_code = None
    last_question = ""
    open_code = None

    def emit_choice(question, label, checked):
        out["choices"].append({"respondent": respondent, "question": question[:120], "option": label[:160], "checked": int(checked)})

    def find_code(text):
        m = code_re.search(text)
        return m.group(1) if m else None

    for child in body.iterchildren():
        if child.tag == qn("w:p"):
            toks = tokens(child)
            text = "".join(v for k, v in toks if k == "text").strip()
            boxes = labeled_boxes(toks)
            c = find_code(text) if text else None
            if c and not boxes:
                last_code = c
            if re.match(r"^\s*(OP\d{2})", text or ""):
                open_code = re.match(r"^\s*(OP\d{2})", text).group(1)
                rest = re.sub(r"^\s*OP\d{2}\s*[.)]?\s*", "", text)
                ans = re.split(r"(?:คำตอบ|ตอบ)\s*[:：]?\s*", rest, maxsplit=1)
                if len(ans) == 2 and ans[1].strip(" .…_"):
                    out["text"].append({"respondent": respondent, "item": open_code, "column": "", "text": ans[1].strip()})
                    open_code = None
                continue
            if open_code and text and not boxes:
                if not re.fullmatch(r"[._…\s-]+", text):
                    out["text"].append({"respondent": respondent, "item": open_code, "column": "", "text": text})
                continue
            if boxes:
                open_code = None
                lead = ""
                for k, v in toks:
                    if k == "box":
                        break
                    lead += v
                lead = lead.strip()
                if re.search(r"คะแนน", lead) and all(value_of_label(lb) for _, lb in boxes):
                    chosen = [value_of_label(lb) for ck, lb in boxes if ck]
                    item = c or last_code or f"?{lead[:20]}"
                    rec_score(out, respondent, item, chosen, f"{Path(path).name}: {lead}")
                else:
                    q = lead or last_question
                    for ck, lb in boxes:
                        emit_choice(q, lb, ck)
            elif text:
                last_question = text
        elif child.tag == qn("w:tbl"):
            parse_table(child, respondent, code_re, out, Path(path).name, last_question)
            open_code = None


def rec_score(out, respondent, item, chosen, where):
    if len(chosen) == 1:
        out["scores"].append({"respondent": respondent, "item": item, "value": chosen[0]})
    else:
        out["issues"].append({"respondent": respondent, "item": item,
                              "problem": "ไม่ได้ติ๊ก" if not chosen else f"ติ๊ก {len(chosen)} ช่อง: {chosen}", "where": where})


def parse_table(tbl, respondent, code_re, out, fname, context_q):
    rows = tbl.findall(qn("w:tr"))
    grid = []
    for tr in rows:
        cells = []
        for tc in tr.findall(qn("w:tc")):
            span = tc.find(qn("w:tcPr") + "/" + qn("w:gridSpan"))
            n = int(span.get(qn("w:val"))) if span is not None else 1
            cells.extend([tc] * n)
        grid.append(cells)
    header = None
    hdr_text = []
    for r, cells in enumerate(grid):
        texts = [plain(tc) for tc in cells]
        boxes_in_row = [labeled_boxes(tokens(tc)) for tc in cells]
        hk = header_kind(texts) if not any(boxes_in_row) else None
        if hk:
            # หัวหลายชั้น: รวมกับหัวเดิม
            header = {**(header or {}), **hk}
            hdr_text = [t or (hdr_text[j] if j < len(hdr_text) else "") for j, t in enumerate(texts)]
            continue
        if not any(boxes_in_row) and not any(code_re.search(t) for t in texts[:2]):
            hdr_text = [t or (hdr_text[j] if j < len(hdr_text) else "") for j, t in enumerate(texts)]
            continue
        code = None
        for t in texts[:2]:
            m = code_re.search(t)
            if m:
                code = m.group(1)
                break
        if header and any(boxes_in_row[j] for j in header if j < len(cells)):
            # จัดกลุ่มคอลัมน์ตามชนิดค่า (คะแนน / เห็นชอบ / MoSCoW)
            kinds = {}
            for j, v in header.items():
                kind = "score" if isinstance(v, int) else ("agree" if v in ("agree", "disagree") else "moscow")
                kinds.setdefault(kind, []).append(j)
            for kind, cols in kinds.items():
                chosen = []
                seen = set()
                for j in cols:
                    if j >= len(cells) or id(cells[j]) in seen:
                        continue
                    seen.add(id(cells[j]))
                    if any(ck for ck, _ in boxes_in_row[j]):
                        chosen.append(str(header[j]))
                item = code or f"ROW{r}:{texts[0][:30]}"
                key = item if kind == "score" else f"{item}#{kind}"
                if kind == "moscow" and not chosen:
                    continue  # ระดับที่เสนอ: กรอกเฉพาะเมื่อไม่เห็นชอบ
                rec_score(out, respondent, key, chosen, f"{fname}: ตาราง แถว {r}")
            # คอลัมน์ข้อความ (เหตุผล/ข้อเสนอแนะ)
            for j, t in enumerate(texts):
                if j in header or j == 0 or not t or (code and code in t):
                    continue
                if j < len(hdr_text) and re.search(r"เหตุผล|ข้อเสนอ|หมายเหตุ|ความคิดเห็น", hdr_text[j] or ""):
                    if not re.fullmatch(r"[._…\s-]+", t):
                        out["text"].append({"respondent": respondent, "item": code or f"ROW{r}", "column": hdr_text[j], "text": t})
            continue
        # ตารางข้อมูลทั่วไป: ช่องติ๊กในเซลล์
        q = texts[0] if texts and not boxes_in_row[0] else context_q
        seen = set()
        for j, bx in enumerate(boxes_in_row):
            if id(cells[j]) in seen:
                continue
            seen.add(id(cells[j]))
            for ck, lb in bx:
                out["choices"].append({"respondent": respondent, "question": (code or q)[:120], "option": lb[:160], "checked": int(ck)})


def main(argv=None):
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("outdir")
    ap.add_argument("inputs", nargs="+")
    ap.add_argument("--codes", default="P|N|FR|NFR|CON|A|AR|E|Q", help="คำนำหน้ารหัสรายการ")
    ap.add_argument("--respondent-from", choices=("filename", "R"), default="filename",
                    help="filename = ชื่อไฟล์, R = R01, R02… ตามลำดับไฟล์")
    a = ap.parse_args(argv)
    files = []
    for x in a.inputs:
        p = Path(x)
        files += sorted(p.glob("*.docx")) if p.is_dir() else [p]
    files = [f for f in files if not f.name.startswith("~$")]
    if not files:
        raise SystemExit("ไม่พบไฟล์ .docx")
    code_re = re.compile(rf"\b((?:{a.codes})\d{{1,2}}(?:\.\d{{1,2}})?)\b")
    out = {"scores": [], "choices": [], "text": [], "issues": []}
    for i, f in enumerate(files, 1):
        resp = f.stem if a.respondent_from == "filename" else f"R{i:02d}"
        parse(str(f), resp, code_re, out)
    od = Path(a.outdir)
    od.mkdir(parents=True, exist_ok=True)
    cols = {"scores": ["respondent", "item", "value"], "choices": ["respondent", "question", "option", "checked"],
            "text": ["respondent", "item", "column", "text"], "issues": ["respondent", "item", "problem", "where"]}
    for k, rows in out.items():
        with open(od / f"{k}.csv", "w", newline="", encoding="utf-8-sig") as fh:
            w = csv.DictWriter(fh, fieldnames=cols[k])
            w.writeheader()
            w.writerows(rows)
    items = sorted({r["item"] for r in out["scores"]})
    print(f"อ่าน {len(files)} ไฟล์: คะแนน {len(out['scores'])} ค่า ({len(items)} รายการ), ช่องเลือก {len(out['choices'])}, "
          f"ข้อความ {len(out['text'])}, ปัญหา {len(out['issues'])}")
    unk = [i for i in items if i.startswith(("ROW", "?"))]
    if unk:
        print("⚠ รายการที่ไม่พบรหัส (ตรวจคอลัมน์แรก):", unk[:10])
    if out["issues"]:
        print("⚠ แถวที่ต้องตรวจกับต้นฉบับ (ไม่ติ๊ก/ติ๊กเกิน):")
        for r in out["issues"][:20]:
            print("  ", r)
    return 1 if out["issues"] else 0


if __name__ == "__main__":
    sys.exit(main())
