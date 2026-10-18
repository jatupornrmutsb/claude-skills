#!/usr/bin/env python3
"""จัดรูปแบบ .docx ตามข้อกำหนดตั้งต้น JDTAIA บนสำเนา (ไม่แตะข้อความใน <w:t>)

  python apply_layout.py in.docx out.docx --page --columns 2 --split-before "1. บทนำ" \
      --fonts --align --spacing --reset-char --tables [--sizes] [--equations] [--anonymize-metadata]

ทุกตัวเลือกปิดไว้จนกว่าจะสั่ง เพื่อให้แก้เฉพาะสิ่งที่ผู้ใช้อนุญาต
  --page              A4 แนวตั้ง ขอบ 3.5/3.0/3.0/2.5 ซม. ทุก section
  --columns N         1 หรือ 2 คอลัมน์ (7.5 ซม. ช่องว่าง 0.5 ซม.) ทุก section
                      ถ้าใช้ --split-before หรือ --keep-first-single section แรกจะเป็น 1 คอลัมน์
  --split-before TXT  แทรก section break แบบต่อเนื่องก่อนย่อหน้าแรกที่ขึ้นต้นด้วย TXT
                      (แยกส่วนชื่อเรื่อง/บทคัดย่อ 1 คอลัมน์ ออกจากเนื้อหา 2 คอลัมน์)
  --fonts             ฟอนต์ TH Sarabun New ครบ ascii/hAnsi/eastAsia/cs ในทุก run, style และค่าเริ่มต้น
                      และเติม bCs/iCs/szCs ที่ขาด (ไม่เช่นนั้นตัวหนา/ขนาดไม่มีผลกับอักษรไทย)
  --align             เนื้อความไทย=thaiDistribute, อังกฤษล้วน=both, หัวข้อ/ชื่อตาราง=left,
                      ชื่อรูป=center, ค่าสั้นในตารางที่ถูกกระจาย=คืนเป็น left
  --spacing           เนื้อความ: Single, ก่อน/หลัง 0 (ไม่ลบย่อหน้าว่าง แค่รายงานจำนวน)
  --reset-char        ลบระยะตัวอักษร (w:spacing) ≠0 และสัดส่วน (w:w) ≠100 ที่ผิดปกติ
  --tables            ทำซ้ำหัวตารางแถวแรกเมื่อ >N แถว, เปลี่ยนความสูง exact → atLeast
  --sizes             ขนาดตามบทบาท: เนื้อความ/อ้างอิง 14 pt, ในตาราง 12 pt, ชื่อตาราง/รูป 14 pt
                      (ไม่เปลี่ยนชื่อบทความ/ผู้แต่ง/สังกัด — ตรวจด้วยตาเพราะจำแนกอัตโนมัติไม่แม่น)
  --equations         run ในสมการ (OMML) ใช้ Times New Roman 10 pt
  --anonymize-metadata ล้าง author/last_modified_by ในคุณสมบัติไฟล์ (ฉบับปกปิดตัวตน)
  --strip-hidden      ลบอักขระแฝงเดิม (U+200B ฯลฯ) ออกจาก <w:t> — ข้อยกเว้นเดียวที่แตะข้อความ
                      ไม่เปลี่ยนตัวอักษรที่มองเห็น แต่ต้องรายงานผู้ใช้ (text_diff จัดเป็น HIDDEN_ONLY)

ไม่ทำ: เปลี่ยนความกว้างคอลัมน์ตาราง, ย้าย/ลบย่อหน้า, ใส่อักขระแฝง, ย่อขนาดเพื่อลดหน้า
ตรวจผลด้วย text_diff.py, docx_audit.py และ render_check.py ทุกครั้ง
"""
from __future__ import annotations

import argparse
import copy
import json
import sys

for _s in (sys.stdout, sys.stderr):  # Windows console ไม่ใช่ UTF-8 จะ print ภาษาไทยไม่ได้
    if hasattr(_s, "reconfigure"):
        _s.reconfigure(encoding="utf-8", errors="replace")
from collections import Counter
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))
from docx.oxml import OxmlElement  # noqa: E402
from docx.oxml.ns import qn  # noqa: E402

import docx_utils as U  # noqa: E402

FONT = U.SPEC["font"]


def set_attr(el, name, val):
    el.set(qn(name), str(val))


def get_or_add(parent, tag, before=None):
    el = parent.find(qn(tag))
    if el is None:
        el = OxmlElement(tag)
        parent.append(el)
    return el


def set_rfonts(rPr, font):
    f = rPr.find(qn("w:rFonts"))
    if f is None:
        f = OxmlElement("w:rFonts")
        rPr.insert(0, f)
    for k in ("asciiTheme", "hAnsiTheme", "eastAsiaTheme", "cstheme"):
        if f.get(qn("w:" + k)) is not None:
            del f.attrib[qn("w:" + k)]
    for k in ("ascii", "hAnsi", "eastAsia", "cs"):
        f.set(qn("w:" + k), font)


def ensure_rPr(r_el):
    rPr = r_el.find(qn("w:rPr"))
    if rPr is None:
        rPr = OxmlElement("w:rPr")
        r_el.insert(0, rPr)
    return rPr


def set_size(rPr, pt):
    for tag in ("w:sz", "w:szCs"):
        e = rPr.find(qn(tag))
        if e is None:
            e = OxmlElement(tag)
            rPr.append(e)
        e.set(qn("w:val"), str(int(pt * 2)))


def mirror_cs(rPr):
    """ข้อความไทยใช้คุณสมบัติ complex script: b→bCs, i→iCs, sz→szCs (เพิ่มเฉพาะที่ขาด)"""
    n = 0
    for latin, cs in (("w:b", "w:bCs"), ("w:i", "w:iCs"), ("w:sz", "w:szCs")):
        le = rPr.find(qn(latin))
        if le is not None and rPr.find(qn(cs)) is None:
            ce = OxmlElement(cs)
            for k, v in le.attrib.items():
                ce.set(k, v)
            le.addnext(ce)
            n += 1
    return n


def page_setup(doc, log):
    for sec in doc.sections:
        sp = sec._sectPr
        pg = sp.find(qn("w:pgSz"))
        if pg is None:
            pg = OxmlElement("w:pgSz")
            sp.append(pg)
        pg.set(qn("w:w"), str(U.SPEC["page_w"]))
        pg.set(qn("w:h"), str(U.SPEC["page_h"]))
        if pg.get(qn("w:orient")):
            del pg.attrib[qn("w:orient")]
        mar = sp.find(qn("w:pgMar"))
        if mar is None:
            mar = OxmlElement("w:pgMar")
            sp.append(mar)
        for k in ("top", "bottom", "left", "right"):
            mar.set(qn("w:" + k), str(U.SPEC[k]))
        for k, v in (("header", 709), ("footer", 709), ("gutter", 0)):
            if mar.get(qn("w:" + k)) is None:
                mar.set(qn("w:" + k), str(v))
        log["page_sections"] += 1


def set_cols(sectPr, n):
    cols = sectPr.find(qn("w:cols"))
    if cols is None:
        cols = OxmlElement("w:cols")
        # ลำดับใน CT_SectPr: cols อยู่หลัง pgMar/pgBorders/lnNumType/pgNumType
        anchor = None
        for tag in ("w:pgNumType", "w:lnNumType", "w:pgBorders", "w:pgMar", "w:pgSz"):
            anchor = sectPr.find(qn(tag))
            if anchor is not None:
                break
        if anchor is not None:
            anchor.addnext(cols)
        else:
            sectPr.append(cols)
    for child in list(cols):
        cols.remove(child)
    for a in list(cols.attrib):
        del cols.attrib[a]
    cols.set(qn("w:num"), str(n))
    cols.set(qn("w:space"), str(U.SPEC["col_space"] if n == 2 else 720))
    if n == 2:
        cols.set(qn("w:equalWidth"), "1")


def set_section_type(sectPr, val):
    t = sectPr.find(qn("w:type"))
    if t is None:
        t = OxmlElement("w:type")
        sectPr.insert(0, t)
    t.set(qn("w:val"), val)


def split_before(doc, prefix, log):
    for kind, p in U.body_items(doc):
        if kind == "p" and U.para_text(p).strip().startswith(prefix):
            prev = p._p.getprevious()
            while prev is not None and prev.tag != qn("w:p"):
                prev = prev.getprevious()
            if prev is None:
                raise SystemExit(f"--split-before: “{prefix}” เป็นย่อหน้าแรก แยก section ไม่ได้")
            pPr = prev.find(qn("w:pPr"))
            if pPr is not None and pPr.find(qn("w:sectPr")) is not None:
                log["split"] = "มี section break อยู่แล้ว ณ จุดนี้"
                return
            if pPr is None:
                pPr = OxmlElement("w:pPr")
                prev.insert(0, pPr)
            final = doc.element.body.find(qn("w:sectPr"))
            new = copy.deepcopy(final)
            for tag in ("w:headerReference", "w:footerReference", "w:titlePg"):
                for e in new.findall(qn(tag)):
                    new.remove(e)
            pPr.append(new)
            set_section_type(final, "continuous")
            log["split"] = f"แทรก section break ก่อน “{prefix}”"
            return
    raise SystemExit(f"--split-before: ไม่พบย่อหน้าที่ขึ้นต้นด้วย “{prefix}”")


def columns(doc, n, keep_first_single, log):
    secs = doc.sections
    for i, sec in enumerate(secs):
        want = n
        if n == 2 and keep_first_single and i == 0 and len(secs) > 1:
            want = 1
        set_cols(sec._sectPr, want)
        log["columns"].append(want)


def fonts(doc, log):
    st = doc.styles.element
    dd = st.find(qn("w:docDefaults"))
    if dd is not None:
        rpd = dd.find(qn("w:rPrDefault"))
        if rpd is None:
            rpd = OxmlElement("w:rPrDefault")
            dd.insert(0, rpd)
        rPr = rpd.find(qn("w:rPr"))
        if rPr is None:
            rPr = OxmlElement("w:rPr")
            rpd.append(rPr)
        set_rfonts(rPr, FONT)
        lang = rPr.find(qn("w:lang"))
        if lang is None:
            lang = OxmlElement("w:lang")
            rPr.append(lang)
        lang.set(qn("w:bidi"), "th-TH")
    for s in st.findall(qn("w:style")):
        rPr = s.find(qn("w:rPr"))
        if rPr is not None:
            if rPr.find(qn("w:rFonts")) is not None:
                set_rfonts(rPr, FONT)
                log["styles_font"] += 1
            log["cs_props_added"] += mirror_cs(rPr)
    body = doc.element.body
    for r in body.iter(qn("w:r")):
        if r.getparent() is not None and r.getparent().tag.startswith("{http://schemas.openxmlformats.org/officeDocument/2006/math}"):
            continue
        rPr = ensure_rPr(r)
        set_rfonts(rPr, FONT)
        log["cs_props_added"] += mirror_cs(rPr)
        log["runs_font"] += 1
    # หัว/ท้ายกระดาษ
    for sec in doc.sections:
        for part in (sec.header, sec.footer):
            if part is not None and not part.is_linked_to_previous:
                for r in part._element.iter(qn("w:r")):
                    set_rfonts(ensure_rPr(r), FONT)


def jc_set(p, val):
    pPr = p._p.get_or_add_pPr()
    jc = pPr.find(qn("w:jc"))
    if jc is None:
        jc = OxmlElement("w:jc")
        pPr.append(jc)
    jc.set(qn("w:val"), val)


def reorder_pPr(pPr):
    """เรียงลูกของ pPr ตาม schema คร่าวๆ (jc/spacing อาจถูก append ท้าย)"""
    order = ["pStyle", "keepNext", "keepLines", "pageBreakBefore", "framePr", "widowControl", "numPr",
             "suppressLineNumbers", "pBdr", "shd", "tabs", "suppressAutoHyphens", "kinsoku", "wordWrap",
             "overflowPunct", "topLinePunct", "autoSpaceDE", "autoSpaceDN", "bidi", "adjustRightInd",
             "snapToGrid", "spacing", "ind", "contextualSpacing", "mirrorIndents", "suppressOverlap", "jc",
             "textDirection", "textAlignment", "textboxTightWrap", "outlineLvl", "divId", "cnfStyle", "rPr",
             "sectPr", "pPrChange"]
    rank = {qn("w:" + k): i for i, k in enumerate(order)}
    kids = list(pPr)
    kids.sort(key=lambda e: rank.get(e.tag, 999))
    for k in kids:
        pPr.remove(k)
    for k in kids:
        pPr.append(k)


def align(doc, log):
    for loc, p, in_table in U.all_paragraphs(doc):
        text = U.para_text(p)
        t = text.strip()
        cls = U.classify(p, text, in_table)
        jc = U.para_jc(p)
        new = None
        if cls in ("body", "reference"):
            if U.has_thai(t):
                new = "thaiDistribute"
            elif U.has_latin(t) and len(t) > 60:
                new = "both"
        elif cls in ("title", "heading", "caption_table", "bad_caption"):
            if jc in ("thaiDistribute", "distribute", "both", None):
                new = "left"
        elif cls == "caption_figure":
            new = "center"
        elif cls == "table_cell" and jc in ("thaiDistribute", "distribute") and len(t) < 15:
            new = "left"
        if new and new != jc:
            jc_set(p, new)
            reorder_pPr(p._p.pPr)
            log["align"][f"{cls}->{new}"] += 1


def spacing(doc, log):
    empties = 0
    for loc, p, in_table in U.all_paragraphs(doc):
        text = U.para_text(p)
        cls = U.classify(p, text, in_table)
        if cls == "empty" and not in_table:
            empties += 1
        if cls not in ("body", "reference"):
            continue
        pPr = p._p.get_or_add_pPr()
        sp = pPr.find(qn("w:spacing"))
        if sp is None:
            sp = OxmlElement("w:spacing")
            pPr.append(sp)
        for a in list(sp.attrib):
            del sp.attrib[a]
        sp.set(qn("w:before"), "0")
        sp.set(qn("w:after"), "0")
        sp.set(qn("w:line"), "240")
        sp.set(qn("w:lineRule"), "auto")
        reorder_pPr(pPr)
        log["spacing_paras"] += 1
    log["empty_paragraphs_kept"] = empties


def reset_char(doc, log):
    for r in doc.element.body.iter(qn("w:r")):
        rPr = r.find(qn("w:rPr"))
        if rPr is None:
            continue
        sp = rPr.find(qn("w:spacing"))
        if sp is not None and int(sp.get(qn("w:val"), "0")) != 0:
            rPr.remove(sp)
            log["char_spacing_reset"] += 1
        w = rPr.find(qn("w:w"))
        if w is not None and int(w.get(qn("w:val"), "100")) != 100:
            rPr.remove(w)
            log["char_scale_reset"] += 1


def tables(doc, log, min_rows):
    for tbl in doc.element.body.iter(qn("w:tbl")):
        rows = tbl.findall(qn("w:tr"))
        if len(rows) > min_rows:
            first = rows[0]
            trPr = first.find(qn("w:trPr"))
            if trPr is None:
                trPr = OxmlElement("w:trPr")
                tblPrEx = first.find(qn("w:tblPrEx"))
                (tblPrEx.addnext(trPr) if tblPrEx is not None else first.insert(0, trPr))
            if trPr.find(qn("w:tblHeader")) is None:
                trPr.append(OxmlElement("w:tblHeader"))
                log["repeat_header_set"] += 1
        for tr in rows:
            trPr = tr.find(qn("w:trPr"))
            if trPr is None:
                continue
            h = trPr.find(qn("w:trHeight"))
            if h is not None and h.get(qn("w:hRule")) == "exact":
                h.set(qn("w:hRule"), "atLeast")
                log["exact_height_relaxed"] += 1


def sizes(doc, log):
    for loc, p, in_table in U.all_paragraphs(doc):
        text = U.para_text(p)
        cls = U.classify(p, text, in_table)
        pt = {"body": 14, "reference": 14, "caption_table": 14, "caption_figure": 14, "bad_caption": 14, "table_cell": 12}.get(cls)
        if not pt:
            continue
        for r in p._p.iter(qn("w:r")):
            if r.getparent().tag.startswith("{http://schemas.openxmlformats.org/officeDocument/2006/math}"):
                continue
            set_size(ensure_rPr(r), pt)
        log["sizes"][f"{cls}={pt}"] += 1


def equations(doc, log):
    M = "{http://schemas.openxmlformats.org/officeDocument/2006/math}"
    for r in doc.element.body.iter(M + "r"):
        rPr = r.find(qn("w:rPr"))
        if rPr is None:
            rPr = OxmlElement("w:rPr")
            r.append(rPr)
        set_rfonts(rPr, "Times New Roman")
        set_size(rPr, 10)
        log["equation_runs"] += 1


def main(argv=None):
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("src")
    ap.add_argument("out")
    ap.add_argument("--page", action="store_true")
    ap.add_argument("--columns", type=int, choices=(1, 2))
    ap.add_argument("--split-before")
    ap.add_argument("--keep-first-single", action="store_true")
    ap.add_argument("--fonts", action="store_true")
    ap.add_argument("--align", action="store_true")
    ap.add_argument("--spacing", action="store_true")
    ap.add_argument("--reset-char", action="store_true")
    ap.add_argument("--tables", action="store_true")
    ap.add_argument("--repeat-header-min-rows", type=int, default=8)
    ap.add_argument("--sizes", action="store_true")
    ap.add_argument("--equations", action="store_true")
    ap.add_argument("--anonymize-metadata", action="store_true")
    ap.add_argument("--strip-hidden", action="store_true", help="ลบอักขระแฝงที่มีอยู่เดิม (เช่น U+200B) — แจ้งผู้ใช้เสมอ")
    a = ap.parse_args(argv)
    if Path(a.src).resolve() == Path(a.out).resolve():
        raise SystemExit("ห้ามเขียนทับต้นฉบับ — ระบุไฟล์ผลลัพธ์ใหม่")
    doc = U.load(a.src)
    log = {"page_sections": 0, "columns": [], "styles_font": 0, "runs_font": 0, "align": Counter(),
           "spacing_paras": 0, "cs_props_added": 0, "char_spacing_reset": 0, "char_scale_reset": 0, "repeat_header_set": 0,
           "exact_height_relaxed": 0, "sizes": Counter(), "equation_runs": 0}
    if a.page:
        page_setup(doc, log)
    if a.split_before:
        split_before(doc, a.split_before, log)
    if a.columns:
        columns(doc, a.columns, a.keep_first_single or bool(a.split_before), log)
    if a.fonts:
        fonts(doc, log)
    if a.reset_char:
        reset_char(doc, log)
    if a.align:
        align(doc, log)
    if a.spacing:
        spacing(doc, log)
    if a.tables:
        tables(doc, log, a.repeat_header_min_rows)
    if a.sizes:
        sizes(doc, log)
    if a.equations:
        equations(doc, log)
    if a.strip_hidden:
        n = 0
        for t in doc.element.body.iter(qn("w:t")):
            if t.text and any(ch in t.text for ch in U.HIDDEN_CHARS):
                n += sum(t.text.count(ch) for ch in U.HIDDEN_CHARS)
                t.text = "".join(ch for ch in t.text if ch not in U.HIDDEN_CHARS)
        log["hidden_chars_removed"] = n
    if a.anonymize_metadata:
        doc.core_properties.author = ""
        doc.core_properties.last_modified_by = ""
        log["metadata"] = "ล้าง author/last_modified_by แล้ว (เนื้อหาในเอกสารต้องตรวจแยก)"
    doc.save(a.out)
    print(json.dumps({k: (dict(v) if isinstance(v, Counter) else v) for k, v in log.items()}, ensure_ascii=False, indent=1))
    print(f"บันทึก: {a.out}  → ตรวจต่อด้วย text_diff.py / docx_audit.py / render_check.py")


if __name__ == "__main__":
    main()
