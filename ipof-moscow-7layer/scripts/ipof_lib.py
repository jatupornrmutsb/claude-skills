"""ไลบรารีกลางของ Skill ipof-moscow-7layer

- โหลด/ตรวจทะเบียนโครงการ (project.json) ซึ่งเป็นแหล่งข้อมูลเดียวของรหัส P/N/FR/NFR/CON
- สถิติ: X̄ และ S.D. แบบตัวอย่าง (n − 1) จากคะแนนดิบ, ค่ารายด้าน/ภาพรวม/ข้อกำหนดรวมจากคะแนนดิบทุกค่า
- แปลผล 5 ระดับ, MoSCoW ตามจุดแบ่ง, เกณฑ์เห็นชอบร้อยละ 80 ของผู้เชี่ยวชาญ
"""
from __future__ import annotations

import csv
import sys

for _s in (sys.stdout, sys.stderr):  # Windows console ไม่ใช่ UTF-8 จะ print ภาษาไทยไม่ได้
    if hasattr(_s, "reconfigure"):
        _s.reconfigure(encoding="utf-8", errors="replace")
import json
import math
import re
from collections import defaultdict
from pathlib import Path

CODE_PATTERNS = {
    "problems": r"P\d{2}", "needs": r"N\d{2}", "constraints": r"CON\d{2}",
    "open_questions": r"OP\d{2}", "functions": r"F\d{2}",
}
REQ_RE = re.compile(r"^(FR|NFR)\d{2}$")
GROUPS = {"G1": "จัดการข้อมูล", "G2": "ประมวลผล", "G3": "สืบค้นและคัดเลือก", "G4": "นำเสนอและสื่อสาร"}
LAYERS = {"L1": "ผู้ใช้งาน", "L2": "การนำเสนอ", "L3": "บริการแอปพลิเคชัน", "L4": "วิเคราะห์ข้อมูล",
          "L5": "ข้อมูล", "L6": "ความมั่นคงและธรรมาภิบาลข้อมูล", "L7": "โครงสร้างพื้นฐานและการเชื่อมต่อ"}
GROUP_LAYER = {"G1": "L5", "G2": "L4", "G3": "L3", "G4": "L2"}
NFR_LAYERS = {"L2", "L6", "L7"}
MOSCOW = {"M": "Must", "S": "Should", "C": "Could", "W": "Won't Have This Time"}
VERIFY = {"Test", "Analysis", "Inspection", "Demonstration"}
LEVELS = [(1.80, "น้อยที่สุด"), (2.60, "น้อย"), (3.40, "ปานกลาง"), (4.20, "มาก"), (5.00, "มากที่สุด")]


# ---------- สถิติ ----------
def mean(xs):
    xs = [float(x) for x in xs]
    return sum(xs) / len(xs) if xs else float("nan")


def sd(xs):
    """S.D. แบบตัวอย่าง (n − 1)"""
    xs = [float(x) for x in xs]
    if len(xs) < 2:
        return float("nan")
    m = mean(xs)
    return math.sqrt(sum((x - m) ** 2 for x in xs) / (len(xs) - 1))


def interpret(m):
    """แปลผลจากค่าเฉลี่ยที่ปัดเป็นทศนิยม 2 ตำแหน่ง (ตรงกับค่าที่รายงาน)"""
    if m is None or math.isnan(m):
        return ""
    r = round(m + 1e-9, 2)
    for hi, label in LEVELS:
        if r <= hi:
            return label
    return LEVELS[-1][1]


def fmt2(x):
    return "" if x is None or (isinstance(x, float) and math.isnan(x)) else f"{x:.2f}"


def pass_threshold(k, pct=0.80):
    """จำนวนผู้เห็นชอบขั้นต่ำ = ceil(ร้อยละ 80 ของผู้เชี่ยวชาญ)"""
    return math.ceil(pct * k - 1e-9)


def agreement_label(agree, k, pct=0.80):
    need = pass_threshold(k, pct)
    if agree >= need:
        return True, f"ถึงเกณฑ์ ({agree}/{k})"
    return False, f"ปรับตามเสียงข้างมาก ({agree}/{k})"


# ---------- ข้อมูลคะแนน (long CSV: respondent,item,value) ----------
def load_scores(path):
    data = defaultdict(dict)  # item -> {respondent: value}
    with open(path, encoding="utf-8-sig") as f:
        for row in csv.DictReader(f):
            v = row["value"].strip()
            if v == "":
                continue
            data[row["item"].strip()][row["respondent"].strip()] = v
    return data


def numeric(values):
    out = []
    for v in values:
        try:
            out.append(float(v))
        except (TypeError, ValueError):
            pass
    return out


def item_stats(scores, items):
    res = {}
    for it in items:
        xs = numeric(scores.get(it, {}).values())
        m = mean(xs) if xs else float("nan")
        res[it] = {"n": len(xs), "mean": m, "sd": sd(xs), "level": interpret(m)}
    return res


def pooled(scores, items):
    """ค่าเฉลี่ย/S.D. จากคะแนนดิบทุกค่าของหลายข้อ (ไม่เฉลี่ยค่าเฉลี่ยรายข้อ)"""
    xs = []
    for it in items:
        xs += numeric(scores.get(it, {}).values())
    m = mean(xs) if xs else float("nan")
    return {"n_values": len(xs), "mean": m, "sd": sd(xs), "level": interpret(m)}


# ---------- ทะเบียนโครงการ ----------
def load_project(path):
    return json.loads(Path(path).read_text(encoding="utf-8"))


def save_project(proj, path):
    Path(path).write_text(json.dumps(proj, ensure_ascii=False, indent=2), encoding="utf-8")


def validate_project(p):
    """ตรวจรหัส การอ้างอิงข้าม และกฎของ Skill คืนรายการ (level, msg)"""
    out = []
    E = lambda m: out.append(("error", m))  # noqa: E731
    W = lambda m: out.append(("warn", m))  # noqa: E731
    for key, pat in CODE_PATTERNS.items():
        codes = [x["code"] for x in p.get(key, [])]
        for c in codes:
            if not re.fullmatch(pat, c):
                E(f"{key}: รหัส {c} ไม่ตรงรูปแบบ {pat}")
        if len(codes) != len(set(codes)):
            E(f"{key}: รหัสซ้ำ")
        nums = [int(re.sub(r"\D", "", c)) for c in codes]
        if nums and nums != list(range(1, len(nums) + 1)):
            W(f"{key}: เลขรหัสไม่ต่อเนื่อง {codes}")
    pj = p.get("project", {})
    for key, label in (("title", "ชื่อเรื่องวิจัย"), ("main_question", "คำถามหลักของระบบ"), ("n_users", "จำนวนผู้ตอบ (n)"),
                       ("n_experts", "จำนวนผู้เชี่ยวชาญ (k)"), ("journal", "วารสารเป้าหมาย")):
        if pj.get(key) in (None, ""):
            W(f"ยังไม่มี{label} — ถามผู้ใช้ ห้ามใส่ค่าสมมติ")
    P = {x["code"] for x in p.get("problems", [])}
    N = {x["code"]: x for x in p.get("needs", [])}
    dims = {d["code"] for d in p.get("dimensions", [])}
    if not 5 <= len(P) <= 8:
        W(f"ปัญหา (P) มี {len(P)} ข้อ — แนวทาง 5–8 ข้อ")
    if not 10 <= len(N) <= 15:
        W(f"ความต้องการ (N) มี {len(N)} ข้อ — แนวทาง 10–15 ข้อ")
    for n in N.values():
        if n.get("dimension") not in dims:
            E(f"{n['code']}: ด้าน {n.get('dimension')} ไม่อยู่ใน dimensions")
        for pc in n.get("problems", []):
            if pc not in P:
                E(f"{n['code']}: อ้างปัญหา {pc} ที่ไม่มี")
    reqs = p.get("requirements", [])
    codes = [r["code"] for r in reqs]
    if len(codes) != len(set(codes)):
        E("requirements: รหัสซ้ำ")
    for t in ("FR", "NFR"):
        nums = [int(c[len(t):]) for c in codes if re.fullmatch(t + r"\d{2}", c)]
        if nums and sorted(nums) != list(range(1, len(nums) + 1)):
            W(f"{t}: เลขไม่ต่อเนื่อง {sorted(nums)}")
    if len(reqs) > 20:
        W(f"ข้อกำหนด {len(reqs)} ข้อ เกิน 20 — พิจารณารวมข้อที่ใช้ข้อมูล/หน้าที่เดียวกัน")
    musts = [r for r in reqs if r.get("moscow_final") == "M"]
    if len(musts) > 8:
        W(f"ข้อ Must {len(musts)} ข้อ เกิน 8")
    used_n = set()
    for r in reqs:
        c = r["code"]
        if not REQ_RE.match(c):
            E(f"{c}: รหัสข้อกำหนดต้องเป็น FR01…/NFR01…")
        if r.get("type") != c.rstrip("0123456789"):
            E(f"{c}: type {r.get('type')} ไม่ตรงกับรหัส")
        if r.get("type") == "FR" and r.get("group") not in GROUPS:
            E(f"{c}: FR ต้องมีกลุ่ม G1–G4")
        if r.get("type") == "NFR" and r.get("group"):
            E(f"{c}: NFR ไม่จัดกลุ่มหน้าที่")
        for nc in r.get("needs", []):
            if nc not in N:
                E(f"{c}: อ้าง {nc} ที่ไม่มี")
            used_n.add(nc)
        for pc in r.get("problems", []):
            if pc not in P:
                E(f"{c}: อ้าง {pc} ที่ไม่มี")
        if r.get("moscow_final") not in MOSCOW:
            E(f"{c}: moscow_final ต้องเป็น M/S/C/W")
        if r.get("moscow_score") and r.get("moscow_final") != r.get("moscow_score") and not r.get("moscow_reason"):
            E(f"{c}: ระดับฉบับสุดท้ายต่างจากลำดับคะแนน ต้องบันทึก moscow_reason")
        if r.get("moscow_final") == "W" and not r.get("moscow_reason"):
            E(f"{c}: Won't ต้องระบุเหตุผลด้านขอบเขต")
        lay = r.get("layer")
        if lay not in LAYERS:
            E(f"{c}: ต้องระบุชั้น L1–L7")
        elif r.get("type") == "FR" and r.get("group") in GROUP_LAYER and lay != GROUP_LAYER[r["group"]]:
            W(f"{c}: {r['group']} ควรอยู่ {GROUP_LAYER[r['group']]} (พบ {lay}) — บันทึกเหตุผลถ้าตั้งใจ")
        elif r.get("type") == "NFR" and lay not in NFR_LAYERS:
            W(f"{c}: NFR ควรอยู่ L2/L6/L7 (พบ {lay})")
        if r.get("moscow_final") == "M":
            if not r.get("component"):
                E(f"{c}: Must ต้องมีองค์ประกอบรับผิดชอบ (component)")
            if r.get("acceptance") and r.get("verification") not in VERIFY:
                E(f"{c}: วิธีตรวจสอบต้องเป็น Test/Analysis/Inspection/Demonstration")
            if not r.get("tc"):
                W(f"{c}: Must ยังไม่มีกรณีทดสอบ TC-xx")
        for d in r.get("depends_on", []):
            if d not in codes:
                E(f"{c}: depends_on {d} ไม่มีในทะเบียน")
        for t in r.get("tc", []):
            if not re.fullmatch(r"TC-\d{2}", t):
                E(f"{c}: รหัสกรณีทดสอบ {t} ต้องเป็น TC-xx")
    # ข้อมูลที่ข้อ Must พึ่งพาต้องเป็น Must ด้วย
    byc = {r["code"]: r for r in reqs}
    for r in musts:
        for d in r.get("depends_on", []):
            if d in byc and byc[d].get("moscow_final") != "M":
                E(f"{r['code']} (Must) พึ่งพา {d} ที่เป็น {byc[d].get('moscow_final')} — ข้อที่ Must ต้องใช้ต้องเป็น Must")
    unused = sorted(set(N) - used_n)
    if unused:
        W(f"ความต้องการที่ยังไม่ถูกสังเคราะห์เป็นข้อกำหนด: {unused}")
    for c in p.get("constraints", []):
        if not c.get("source"):
            W(f"{c['code']}: ระบุที่มา (คำตอบปลายเปิด/ข้อเสนอแนะผู้เชี่ยวชาญ)")
    cut = p.get("moscow_cutoffs") or {}
    if not {"must_min", "should_min"} <= set(cut):
        W("ยังไม่ระบุจุดแบ่ง MoSCoW (moscow_cutoffs.must_min / should_min)")
    return out


def requirement_scores(p, scores):
    """ค่าเฉลี่ยข้อกำหนด = pooled คะแนนดิบทุกค่าของ N ที่รวมมา"""
    return {r["code"]: pooled(scores, r.get("needs", [])) for r in p.get("requirements", []) if r.get("needs")}


def moscow_by_score(m, cut):
    r = round(m + 1e-9, 2)
    if r >= cut["must_min"]:
        return "M"
    if r >= cut["should_min"]:
        return "S"
    return "C"
