#!/usr/bin/env python3
"""วิเคราะห์ผลและตรวจทะเบียนข้อกำหนด

  python analyze.py validate --project project.json
  python analyze.py user     --project project.json --data DIR [--json results.json]
  python analyze.py agree    --data DIR --items FR01,FR02 [--k 5]
  python analyze.py compare  OLD_DIR NEW_DIR

user    : X̄/S.D. รายข้อ รายด้าน ภาพรวม (จากคะแนนดิบ), ค่าเฉลี่ยข้อกำหนด (pooled จาก N ที่รวม),
          MoSCoW ตามจุดแบ่งเทียบระดับฉบับสุดท้าย พร้อมเหตุผล
agree   : จำนวนเห็นชอบ/ร้อยละ/ผล "ถึงเกณฑ์ (x/k)" หรือ "ปรับตามเสียงข้างมาก (x/k)" และระดับที่ถูกเสนอมากที่สุด
compare : เทียบแบบประเมินชุดใหม่กับชุดเดิมรายข้อ รายงานข้อที่เปลี่ยน
"""
from __future__ import annotations

import argparse
import json
import sys

for _s in (sys.stdout, sys.stderr):  # Windows console ไม่ใช่ UTF-8 จะ print ภาษาไทยไม่ได้
    if hasattr(_s, "reconfigure"):
        _s.reconfigure(encoding="utf-8", errors="replace")
from collections import Counter
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))
import ipof_lib as L  # noqa: E402


def cmd_validate(a):
    p = L.load_project(a.project)
    res = L.validate_project(p)
    for lvl, m in res:
        print(f"[{lvl}] {m}")
    reqs = p.get("requirements", [])
    cnt = Counter((r["type"], r.get("group") or "-", r.get("moscow_final")) for r in reqs)
    print(f"ข้อกำหนด {len(reqs)} ข้อ: FR {sum(1 for r in reqs if r['type']=='FR')}, NFR {sum(1 for r in reqs if r['type']=='NFR')}, "
          f"Must {sum(1 for r in reqs if r.get('moscow_final')=='M')}")
    print("ไม่พบปัญหา" if not res else f"error {sum(1 for l, _ in res if l == 'error')}, warn {sum(1 for l, _ in res if l == 'warn')}")
    return 1 if any(l == "error" for l, _ in res) else 0


def cmd_user(a):
    p = L.load_project(a.project)
    sc = L.load_scores(Path(a.data) / "scores.csv")
    out = {"items": {}, "dimensions": {}, "requirements": {}, "moscow": {}}
    P = [x["code"] for x in p["problems"]]
    st = L.item_stats(sc, P)
    out["items"].update(st)
    out["dimensions"]["ปัญหา (P)"] = L.pooled(sc, P)
    dims = {d["code"]: d["name"] for d in p["dimensions"]}
    allN = []
    for dc, dn in dims.items():
        codes = [n["code"] for n in p["needs"] if n["dimension"] == dc]
        allN += codes
        out["items"].update(L.item_stats(sc, codes))
        out["dimensions"][dn] = L.pooled(sc, codes)
    out["dimensions"]["ความต้องการภาพรวม"] = L.pooled(sc, allN)
    n_resp = len({r for it in sc.values() for r in it})
    print(f"ผู้ตอบ n = {n_resp}")
    print(f"{'รหัส':6} {'n':>3} {'X̄':>5} {'S.D.':>5} แปลผล")
    for c, s in out["items"].items():
        print(f"{c:6} {s['n']:>3} {L.fmt2(s['mean']):>5} {L.fmt2(s['sd']):>5} {s['level']}")
    for d, s in out["dimensions"].items():
        print(f"[{d}] X̄ {L.fmt2(s['mean'])} S.D. {L.fmt2(s['sd'])} {s['level']} (คะแนนดิบ {s['n_values']} ค่า)")
    rs = L.requirement_scores(p, sc)
    cut = p.get("moscow_cutoffs")
    print("\nข้อกำหนด (ค่าเฉลี่ยจากคะแนนดิบทุกค่าของ N ที่รวม):")
    order = sorted(rs, key=lambda c: -rs[c]["mean"])
    byc = {r["code"]: r for r in p["requirements"]}
    for c in order:
        s = rs[c]
        r = byc[c]
        per = ", ".join(f"{n} {L.fmt2(out['items'].get(n, {}).get('mean', float('nan')))}" for n in r["needs"])
        by_score = L.moscow_by_score(s["mean"], cut) if cut else "?"
        fin = r.get("moscow_final")
        flag = "" if fin == by_score else f"  ← ฉบับสุดท้าย {fin}: {r.get('moscow_reason') or '⚠ ยังไม่มีเหตุผล'}"
        print(f"  {c:6} X̄ {L.fmt2(s['mean'])} S.D. {L.fmt2(s['sd'])} ตามคะแนน {by_score}{flag}   [{per}]")
        out["requirements"][c] = s
        out["moscow"][c] = {"by_score": by_score, "final": fin, "reason": r.get("moscow_reason", "")}
    if cut:
        print(f"จุดแบ่ง: Must ≥ {cut['must_min']:.2f}, Should ≥ {cut['should_min']:.2f}, ต่ำกว่านั้น Could; Won't กำหนดจากขอบเขต")
    if a.json:
        Path(a.json).write_text(json.dumps(out, ensure_ascii=False, indent=1, default=float), encoding="utf-8")
    return 0


def cmd_agree(a):
    sc = L.load_scores(Path(a.data) / "scores.csv")
    items = a.items.split(",")
    rows = []
    for c in items:
        votes = sc.get(f"{c}#agree", {})
        k = a.k or len(votes)
        ag = sum(1 for v in votes.values() if v == "agree")
        ok, lab = L.agreement_label(ag, k)
        prop = Counter(sc.get(f"{c}#moscow", {}).values())
        top = prop.most_common()
        tie = len(top) > 1 and top[0][1] == top[1][1]
        sug = "" if not top else ("เสมอ — ผู้วิจัยตัดสินพร้อมเหตุผล" if tie else top[0][0])
        rows.append((c, ag, k, ok, lab, dict(prop), sug))
        print(f"{c}: เห็นชอบ {ag}/{k} ({ag / k * 100 if k else 0:.2f}%) → {lab}" + ("" if ok else f"; เสนอ {dict(prop)} → {sug}"))
    return 0


def cmd_compare(a):
    old = L.load_scores(Path(a.old) / "scores.csv")
    new = L.load_scores(Path(a.new) / "scores.csv")
    changed = []
    for it in sorted(set(old) | set(new)):
        o, n = old.get(it, {}), new.get(it, {})
        for r in sorted(set(o) | set(n)):
            if o.get(r) != n.get(r):
                changed.append((it, r, o.get(r, "(ไม่มี)"), n.get(r, "(ไม่มี)")))
    if not changed:
        print("ชุดใหม่ตรงกับชุดเดิมทุกข้อ")
    for it, r, x, y in changed:
        print(f"  {it} ผู้ตอบ {r}: {x} → {y}")
    new_items = sorted(set(new) - set(old))
    gone = sorted(set(old) - set(new))
    if new_items:
        print("รายการใหม่:", new_items)
    if gone:
        print("รายการที่หายไป:", gone)
    print(f"เปลี่ยน {len(changed)} ค่า")
    return 0


def main(argv=None):
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    sp = ap.add_subparsers(dest="cmd", required=True)
    v = sp.add_parser("validate"); v.add_argument("--project", required=True)  # noqa: E702
    u = sp.add_parser("user"); u.add_argument("--project", required=True); u.add_argument("--data", required=True); u.add_argument("--json")  # noqa: E702
    g = sp.add_parser("agree"); g.add_argument("--data", required=True); g.add_argument("--items", required=True); g.add_argument("--k", type=int)  # noqa: E702
    c = sp.add_parser("compare"); c.add_argument("old"); c.add_argument("new")  # noqa: E702
    a = ap.parse_args(argv)
    return {"validate": cmd_validate, "user": cmd_user, "agree": cmd_agree, "compare": cmd_compare}[a.cmd](a)


if __name__ == "__main__":
    sys.exit(main())
