#!/usr/bin/env python3
"""ตรวจรูปแบบ .docx เทียบข้อกำหนดตั้งต้น JDTAIA (อ่านอย่างเดียว ไม่แก้ไฟล์)

ใช้:
  python docx_audit.py article.docx [--columns 1|2] [--json out.json]

ตรวจ: หน้ากระดาษ/ขอบ/คอลัมน์ทุก section, ฟอนต์ทุกช่อง (ascii/hAnsi/eastAsia/cs),
ขนาด sz/szCs, ระยะตัวอักษร/สัดส่วน, อักขระแฝง, ช่องว่าง/แท็บต่อเนื่อง, ขึ้นบรรทัดเอง,
การจัดแนว (Thai Distribute), เลขหัวข้อ, ชื่อตาราง/รูป, คุณสมบัติตาราง, สมการ, ภาคผนวก
และนับสถานะก่อนแก้ (ตาราง รูป อ้างอิง)

ผลเป็น "ข้อสังเกต" ต้องตรวจภาพหน้าจริงประกอบเสมอ
"""
from __future__ import annotations

import argparse
import json
import re
import sys

for _s in (sys.stdout, sys.stderr):  # Windows console ไม่ใช่ UTF-8 จะ print ภาษาไทยไม่ได้
    if hasattr(_s, "reconfigure"):
        _s.reconfigure(encoding="utf-8", errors="replace")
from collections import Counter
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))
from docx.oxml.ns import qn  # noqa: E402

import docx_utils as U  # noqa: E402


def add(issues, level, code, loc, msg):
    issues.append({"level": level, "code": code, "loc": loc, "msg": msg})


def check_sections(doc, issues, want_cols):
    info = []
    for i, sec in enumerate(doc.sections):
        sp = sec._sectPr
        pg = sp.find(qn("w:pgSz"))
        mar = sp.find(qn("w:pgMar"))
        w = int(pg.get(qn("w:w"))) if pg is not None else None
        h = int(pg.get(qn("w:h"))) if pg is not None else None
        orient = pg.get(qn("w:orient")) if pg is not None else None
        m = {k: int(mar.get(qn("w:" + k))) for k in ("top", "bottom", "left", "right")} if mar is not None else {}
        ncols, space = U.section_cols(sec)
        sinfo = {"section": i, "w": w, "h": h, "orient": orient, "margins_cm": {k: U.twips_to_cm(v) for k, v in m.items()},
                 "columns": ncols, "col_space_cm": U.twips_to_cm(space) if space else None}
        info.append(sinfo)
        loc = f"section{i}"
        if w is None or abs(w - U.SPEC["page_w"]) > U.TOL or abs(h - U.SPEC["page_h"]) > U.TOL or orient == "landscape":
            add(issues, "error", "PAGE_SIZE", loc, f"ไม่ใช่ A4 แนวตั้ง (w={w}, h={h}, orient={orient})")
        for k in ("top", "bottom", "left", "right"):
            if k in m and abs(m[k] - U.SPEC[k]) > U.TOL:
                add(issues, "error", "MARGIN", loc, f"ขอบ{k} {U.twips_to_cm(m[k])} ซม. (ตั้งต้น {U.twips_to_cm(U.SPEC[k])} ซม.)")
        if ncols == 2 and space and abs(int(space) - U.SPEC["col_space"]) > U.TOL:
            add(issues, "warn", "COL_SPACE", loc, f"ช่องว่างคอลัมน์ {U.twips_to_cm(space)} ซม. (ตั้งต้น 0.5 ซม.)")
        if ncols > 2:
            add(issues, "error", "COLS", loc, f"{ncols} คอลัมน์")
    if want_cols:
        bad = [s["section"] for s in info if s["columns"] != want_cols]
        if want_cols == 1 and bad:
            add(issues, "error", "COLS_LEFTOVER", "document", f"สั่ง 1 คอลัมน์ แต่ section {bad} ยังไม่ใช่ 1 คอลัมน์")
        if want_cols == 2 and len(bad) == len(info):
            add(issues, "error", "COLS_MISSING", "document", "สั่ง 2 คอลัมน์ แต่ไม่มี section ใดเป็น 2 คอลัมน์")
        if want_cols == 2 and bad and len(bad) < len(info):
            add(issues, "info", "COLS_MIXED", "document", f"section {bad} เป็น 1 คอลัมน์ (ปกติสำหรับส่วนชื่อเรื่อง/ตารางเต็มหน้า — ตรวจว่าตั้งใจ)")
    return info


def run_props(r):
    rPr = r._r.rPr
    d = {}
    if rPr is None:
        return d
    f = rPr.find(qn("w:rFonts"))
    if f is not None:
        for k in ("ascii", "hAnsi", "eastAsia", "cs"):
            v = f.get(qn("w:" + k))
            if v:
                d[k] = v
    for tag in ("w:b", "w:bCs"):
        e = rPr.find(qn(tag))
        if e is not None:
            d[tag[2:]] = e.get(qn("w:val"), "true") not in ("0", "false")
    for tag, key in (("w:sz", "sz"), ("w:szCs", "szCs"), ("w:spacing", "charSpacing"), ("w:w", "scale"), ("w:position", "position")):
        e = rPr.find(qn(tag))
        if e is not None:
            d[key] = int(e.get(qn("w:val")))
    return d


def default_fonts(doc):
    st = doc.styles.element
    rd = st.find(qn("w:docDefaults"))
    out = {}
    if rd is not None:
        f = rd.find(".//" + qn("w:rFonts"))
        if f is not None:
            for k in ("ascii", "hAnsi", "eastAsia", "cs"):
                out[k] = f.get(qn("w:" + k)) or f.get(qn("w:" + k + "Theme"))
    return out


def check_paragraphs(doc, issues):
    font = U.SPEC["font"]
    counts = Counter()
    jc_by_class = Counter()
    heading_nums = []
    tables_caps, fig_caps = [], []
    for loc, p, in_table in U.all_paragraphs(doc):
        text = U.para_text(p)
        cls = U.classify(p, text, in_table)
        counts[cls] += 1
        jc = U.para_jc(p)
        jc_by_class[(cls, jc)] += 1
        t = text.strip()
        short = t[:50]
        for ch, name in U.HIDDEN_CHARS.items():
            n = text.count(ch)
            if n:
                add(issues, "error", "HIDDEN_CHAR", loc, f"{name} x{n}: “{short}”")
        if re.search(r" {3,}", text):
            add(issues, "warn", "MULTI_SPACE", loc, f"ช่องว่างต่อเนื่อง ≥3: “{short}”")
        if re.search(r"\t{2,}", text):
            add(issues, "warn", "MULTI_TAB", loc, f"แท็บต่อเนื่อง: “{short}”")
        nbr = text.count("\n")
        if nbr and cls in ("body", "reference") and U.has_thai(text):
            lvl = "warn" if nbr >= 2 else "info"
            add(issues, lvl, "MANUAL_BREAK", loc, f"ขึ้นบรรทัดเอง {nbr} จุด (ตรวจว่าไม่ตัดกลางคำ/ไม่ยืดบรรทัดสั้น): “{short}”")
        # การจัดแนว
        if cls in ("body", "reference") and U.has_thai(t) and len(t) > 60 and jc != "thaiDistribute":
            add(issues, "warn", "NOT_THAI_DIST", loc, f"เนื้อความไทยจัดแนว {jc or 'left(default)'}: “{short}”")
        if cls in ("body", "reference") and not U.has_thai(t) and U.has_latin(t) and len(t) > 60 and jc == "thaiDistribute":
            add(issues, "info", "EN_THAI_DIST", loc, f"ภาษาอังกฤษล้วนใช้ thaiDistribute (พิจารณา both): “{short}”")
        if cls in ("title", "heading", "caption_table", "caption_figure") and jc in ("thaiDistribute", "distribute"):
            add(issues, "warn", "DIST_ON_HEADING", loc, f"{cls} ถูกกระจาย: “{short}”")
        if cls == "caption_figure" and jc not in ("center",):
            add(issues, "warn", "FIG_CAPTION_ALIGN", loc, f"ชื่อรูปไม่กึ่งกลาง: “{short}”")
        if cls == "table_cell" and jc in ("thaiDistribute", "distribute") and len(t) < 15:
            add(issues, "warn", "DIST_SHORT_CELL", loc, f"ค่าสั้นในตารางถูกกระจาย: “{short}”")
        if cls == "bad_caption":
            add(issues, "error", "CAPTION_WORDING", loc, f"ใช้ “ตารางแสดง…” แทน “ตารางที่ …”: “{short}”")
        if cls == "heading":
            m = U.HEADING_RE.match(t)
            if m:
                heading_nums.append((loc, m.group(1)))
        if cls == "caption_table":
            tables_caps.append((loc, int(U.CAPTION_TABLE_RE.match(t).group(1))))
        if cls == "caption_figure" and U.CAPTION_FIG_RE.match(t):
            fig_caps.append((loc, int(U.CAPTION_FIG_RE.match(t).group(1))))
        # ระยะก่อน/หลังเนื้อความ
        if cls == "body":
            sp = p._p.pPr.find(qn("w:spacing")) if p._p.pPr is not None else None
            if sp is not None:
                b, a = sp.get(qn("w:before")), sp.get(qn("w:after"))
                line, rule = sp.get(qn("w:line")), sp.get(qn("w:lineRule"))
                if (b and int(b) > 0) or (a and int(a) > 0):
                    add(issues, "info", "PARA_SPACING", loc, f"ระยะก่อน/หลัง {b}/{a} twips: “{short}”")
                if line and not (rule in (None, "auto") and int(line) == 240):
                    add(issues, "info", "LINE_SPACING", loc, f"ระยะบรรทัด line={line} rule={rule} (ตั้งต้น Single): “{short}”")
        # run
        for r in p.runs:
            if not r.text.strip():
                continue
            rp = run_props(r)
            fonts = {k: v for k, v in rp.items() if k in ("ascii", "hAnsi", "eastAsia", "cs")}
            is_eq_font = cls == "equation"
            for k, v in fonts.items():
                if v != font and not is_eq_font:
                    add(issues, "warn", "FONT", loc, f"rFonts {k}={v}: “{r.text[:30]}”")
                    break
            if U.has_thai(r.text) and "cs" not in fonts and fonts:
                add(issues, "warn", "FONT_CS_MISSING", loc, f"run มีอักษรไทยแต่ไม่กำหนดฟอนต์ cs: “{r.text[:30]}”")
            if U.has_thai(r.text) and rp.get("b") and "bCs" not in rp:
                add(issues, "warn", "BCS_MISSING", loc, f"ตัวหนามีแค่ w:b ไม่มี w:bCs — อักษรไทยจะไม่หนา: “{r.text[:30]}”")
            if U.has_thai(r.text) and "sz" in rp and "szCs" not in rp:
                add(issues, "warn", "SZCS_MISSING", loc, f"ขนาดมีแค่ w:sz={rp['sz']/2}pt ไม่มี w:szCs — อักษรไทยใช้ขนาดอื่น: “{r.text[:30]}”")
            if "sz" in rp and "szCs" in rp and rp["sz"] != rp["szCs"]:
                add(issues, "warn", "SZ_MISMATCH", loc, f"sz={rp['sz']/2}pt szCs={rp['szCs']/2}pt (ไทยใช้ szCs): “{r.text[:30]}”")
            if rp.get("charSpacing"):
                add(issues, "warn", "CHAR_SPACING", loc, f"ระยะตัวอักษร {rp['charSpacing']} twips: “{r.text[:30]}”")
            if rp.get("scale") and rp["scale"] != 100:
                add(issues, "warn", "CHAR_SCALE", loc, f"สัดส่วนตัวอักษร {rp['scale']}%: “{r.text[:30]}”")
            if rp.get("position"):
                add(issues, "info", "CHAR_POSITION", loc, f"ยก/ลดตัวอักษร {rp['position']}: “{r.text[:30]}”")
    # เลขหัวข้อ
    check_heading_numbers(heading_nums, issues)
    for name, caps in (("ตาราง", tables_caps), ("รูป", fig_caps)):
        nums = [n for _, n in caps]
        if nums and nums != list(range(1, len(nums) + 1)):
            add(issues, "error", "CAPTION_NUMBERING", "document", f"เลข{name}ไม่เรียง 1..n: {nums}")
    return counts, jc_by_class, tables_caps, fig_caps


def check_heading_numbers(heads, issues):
    prev = None
    seen = set()
    for loc, num in heads:
        parts = [int(x) for x in num.split(".")]
        if len(parts) > 3:
            add(issues, "error", "HEADING_DEPTH", loc, f"หัวข้อเกิน 3 ระดับ: {num}")
        if num in seen:
            add(issues, "error", "HEADING_DUP", loc, f"เลขหัวข้อซ้ำ: {num}")
        seen.add(num)
        if prev is not None:
            ok = False
            # ถัดไปที่ถูกต้อง: ลึกลง 1 ระดับเริ่ม 1, หรือ เพิ่ม 1 ที่ระดับใดระดับหนึ่ง
            if len(parts) == len(prev) + 1 and parts[:-1] == prev and parts[-1] == 1:
                ok = True
            for lvl in range(min(len(parts), len(prev))):
                if parts[:lvl] == prev[:lvl] and parts[lvl] == prev[lvl] + 1 and len(parts) == lvl + 1:
                    ok = True
            if not ok:
                add(issues, "error", "HEADING_SEQ", loc, f"เลขหัวข้อข้ามหรือผิดลำดับ: {'.'.join(map(str, prev))} → {num}")
        elif parts != [1]:
            add(issues, "warn", "HEADING_START", loc, f"หัวข้อแรกเป็น {num}")
        prev = parts


def check_tables(doc, issues):
    out = []
    ti = 0
    for i, (kind, t) in enumerate(U.body_items(doc)):
        if kind != "t":
            continue
        ti += 1
        loc = f"T{i}"
        tbl = t._tbl
        tblPr = tbl.tblPr
        tw = tblPr.find(qn("w:tblW")) if tblPr is not None else None
        grid = [int(g.get(qn("w:w"))) for g in tbl.findall(qn("w:tblGrid") + "/" + qn("w:gridCol")) if g.get(qn("w:w"))]
        total = sum(grid) if grid else None
        rows = tbl.findall(qn("w:tr"))
        hdr = [r for r in rows if r.find(qn("w:trPr")) is not None and r.find(qn("w:trPr")).find(qn("w:tblHeader")) is not None]
        exact = [ri for ri, r in enumerate(rows) if r.find(qn("w:trPr")) is not None and (r.find(qn("w:trPr")).find(qn("w:trHeight")) is not None) and r.find(qn("w:trPr")).find(qn("w:trHeight")).get(qn("w:hRule")) == "exact"]
        merged_v = len(tbl.findall(".//" + qn("w:vMerge")))
        merged_h = len(tbl.findall(".//" + qn("w:gridSpan")))
        borders = tblPr.find(qn("w:tblBorders")) if tblPr is not None else None
        style = tblPr.find(qn("w:tblStyle")) if tblPr is not None else None
        tinfo = {"index": ti, "loc": loc, "rows": len(rows), "grid_twips": grid, "width_cm": U.twips_to_cm(total) if total else None,
                 "tblW": (tw.get(qn("w:w")), tw.get(qn("w:type"))) if tw is not None else None,
                 "header_rows": len(hdr), "exact_height_rows": exact, "vMerge": merged_v, "gridSpan": merged_h}
        out.append(tinfo)
        if total:
            if total > U.SPEC["text_width"] + U.TOL:
                add(issues, "error", "TABLE_WIDTH", loc, f"ตารางที่ลำดับ {ti} กว้าง {U.twips_to_cm(total)} ซม. เกิน 15.5 ซม.")
            elif U.SPEC["col_width"] + U.TOL < total < U.SPEC["text_width"] - 200:
                add(issues, "info", "TABLE_WIDTH_MID", loc, f"ตารางลำดับ {ti} กว้าง {U.twips_to_cm(total)} ซม. — ในหน้า 2 คอลัมน์ต้อง ≤7.5 ซม. หรือวางในส่วนเต็มหน้า")
            if len(grid) > 2 and len(set(grid)) == 1:
                add(issues, "info", "TABLE_EQUAL_COLS", loc, f"ตารางลำดับ {ti} ทุกคอลัมน์กว้างเท่ากัน — ตรวจความเหมาะสมตามชนิดข้อมูล")
        if tw is not None and tw.get(qn("w:type")) == "pct":
            add(issues, "warn", "TABLE_PCT", loc, f"ตารางลำดับ {ti} ใช้ความกว้างแบบ % (มักเลื่อนระหว่างโปรแกรม)")
        if len(rows) > 12 and not hdr:
            add(issues, "warn", "TABLE_NO_REPEAT_HEADER", loc, f"ตารางลำดับ {ti} มี {len(rows)} แถว แต่ไม่ตั้งทำซ้ำหัวตาราง")
        if exact:
            add(issues, "warn", "TABLE_EXACT_HEIGHT", loc, f"ตารางลำดับ {ti} แถว {exact[:10]} สูงตายตัว (exact) ข้อความอาจถูกตัด")
        if borders is None and style is None:
            add(issues, "warn", "TABLE_BORDERS", loc, f"ตารางลำดับ {ti} ไม่มีเส้นกรอบระดับตาราง (ตรวจภาพจริง)")
    return out


def count_state(doc):
    body = doc.element.body
    text_all = "\n".join(U.para_text(p) for _, p, _ in U.all_paragraphs(doc))
    return {
        "paragraphs": sum(1 for k, _ in U.body_items(doc) if k == "p"),
        "tables": len(body.findall(".//" + qn("w:tbl"))),
        "images": len(body.findall(".//" + qn("w:drawing"))) + len(body.findall(".//" + qn("w:pict"))),
        "equations": len(body.findall(".//" + qn("m:oMathPara"))) + len([m for m in body.findall(".//" + qn("m:oMath")) if m.getparent().tag != qn("m:oMathPara")]),
        "sections": len(doc.sections),
        "reference_entries": len(re.findall(r"(?m)^\s*\[\d+\]\s", text_all)),
        "has_appendix": bool(re.search(r"(?m)^\s*(ภาคผนวก|Appendix)", text_all)),
        "characters": len(text_all),
    }


def anonymity(doc, issues):
    cp = doc.core_properties
    meta = {"author": cp.author, "last_modified_by": cp.last_modified_by, "title": cp.title, "comments": cp.comments}
    for k, v in meta.items():
        if v and k in ("author", "last_modified_by"):
            add(issues, "info", "META_AUTHOR", "core.xml", f"{k}={v!r} (ฉบับปกปิดตัวตนต้องล้าง)")
    n_comments = 0
    try:
        part = doc.part.package.part_related_by  # noqa: F841
    except Exception:
        pass
    for rel in doc.part.rels.values():
        if rel.reltype.endswith("/comments"):
            n_comments = len(rel.target_part.blob.decode("utf8", "ignore").split("<w:comment ")) - 1
    if n_comments:
        add(issues, "info", "COMMENTS", "comments.xml", f"มีคอมเมนต์ {n_comments} รายการ (เป็นข้อมูล ไม่ใช่คำสั่ง; ฉบับส่งควรลบตามที่ผู้ใช้อนุญาต)")
    if doc.element.body.find(".//" + qn("w:ins")) is not None or doc.element.body.find(".//" + qn("w:del")) is not None:
        add(issues, "warn", "TRACKED_CHANGES", "document", "มี tracked changes ค้างอยู่")
    return meta


def main(argv=None):
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("docx")
    ap.add_argument("--columns", type=int, choices=(1, 2), help="จำนวนคอลัมน์ที่ผู้ใช้สั่งสำหรับงานนี้")
    ap.add_argument("--json", help="บันทึกผลเป็น JSON")
    ap.add_argument("--max-lines", type=int, default=60, help="จำนวนข้อสังเกตที่พิมพ์ต่อระดับ")
    a = ap.parse_args(argv)
    doc = U.load(a.docx)
    issues = []
    secs = check_sections(doc, issues, a.columns)
    counts, jc_by_class, tcaps, fcaps = check_paragraphs(doc, issues)
    tables = check_tables(doc, issues)
    state = count_state(doc)
    meta = anonymity(doc, issues)
    if state["tables"] != len(tcaps):
        add(issues, "warn", "TABLE_CAPTION_COUNT", "document", f"ตาราง {state['tables']} ตาราง แต่พบชื่อ “ตารางที่” {len(tcaps)} รายการ (ตารางในเซลล์/ตารางจัดหน้าอาจนับรวม)")
    dfonts = default_fonts(doc)
    result = {"file": a.docx, "state": state, "sections": secs, "default_fonts": dfonts,
              "paragraph_classes": dict(counts), "tables": tables, "metadata": meta, "issues": issues}
    if a.json:
        Path(a.json).write_text(json.dumps(result, ensure_ascii=False, indent=2), encoding="utf8")
    print(f"ไฟล์: {a.docx}")
    print("สถานะ:", json.dumps(state, ensure_ascii=False))
    for s in secs:
        print(f"  section{s['section']}: {s['columns']} คอลัมน์ ขอบ(ซม.) {s['margins_cm']}")
    print("ฟอนต์ค่าเริ่มต้นเอกสาร:", dfonts)
    by = Counter(i["level"] for i in issues)
    print(f"ข้อสังเกต: error={by['error']} warn={by['warn']} info={by['info']}")
    for lvl in ("error", "warn", "info"):
        items = [i for i in issues if i["level"] == lvl]
        for i in items[: a.max_lines]:
            print(f"  [{lvl}] {i['code']} @{i['loc']}: {i['msg']}")
        if len(items) > a.max_lines:
            print(f"  ... อีก {len(items) - a.max_lines} รายการ (ดู --json)")
    return 1 if by["error"] else 0


if __name__ == "__main__":
    sys.exit(main())
