#!/usr/bin/env python3
"""สร้างแผนภาพจาก project.json เป็น HTML (แก้ไขได้) แล้ว render เป็น PNG ความละเอียดสูง

  python render_diagrams.py --project project.json --outdir DIR [--only ipof,layers,context,dataflow,trace]
         [--scale 3] [--no-png]

ได้ไฟล์: ipof.html/png, layers7.html/png, context.html/png, dataflow.html/png, traceability.html/.csv
และรายงานความครบถ้วน: ข้อกำหนดทุกข้ออยู่ในภาพ 7 ชั้น, ข้อ Must ทุกข้อมีองค์ประกอบรับผิดชอบ,
ทุกกล่องมีรหัสและระดับ (ระดับฉบับสุดท้าย)

การจัดวาง IPOF (กำหนดเองได้ด้วย requirement.ipof = {"part": "...", "sub": "..."})
  Input    = FR กลุ่ม G1 จัดกลุ่มตาม data_type
  Process  = การรับข้อมูลอัตโนมัติ / การตรวจคุณภาพและที่มา (NFR) / การวิเคราะห์ (FR G2)
  Output   = FR ระดับ Must กลุ่ม G3 และ G4
  Feedback = การแจ้งปัญหา / การตรวจหลักฐานและแก้ข้อมูล / การบันทึกประวัติ
  แถบกำกับ = NFR ด้านสิทธิ์และการปกป้องข้อมูล (ipof.part = "Governance" หรือ NFR ชั้น L6)
"""
from __future__ import annotations

import argparse
import csv
import html
import re
import shutil
import subprocess
import sys

for _s in (sys.stdout, sys.stderr):  # Windows console ไม่ใช่ UTF-8 จะ print ภาษาไทยไม่ได้
    if hasattr(_s, "reconfigure"):
        _s.reconfigure(encoding="utf-8", errors="replace")
from collections import OrderedDict
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))
import ipof_lib as L  # noqa: E402

CSS = """
:root{--ink:#1f2937;--muted:#6b7280;--line:#94a3b8;--M:#b91c1c;--S:#a16207;--C:#1d4ed8;--W:#6b7280;
--bg:#ffffff;--panel:#f8fafc;--band:#ecfdf5;--bandline:#059669;}
*{box-sizing:border-box} body{margin:0;background:var(--bg);color:var(--ink);
font-family:'TH Sarabun New','TH SarabunPSK','Sarabun','Loma',sans-serif;font-size:20px;line-height:1.15}
.wrap{padding:24px;width:max-content;min-width:1200px}
h1{font-size:30px;margin:0 0 12px} .sub{color:var(--muted);font-size:18px;margin-bottom:14px}
.chip{display:inline-block;border:1.5px solid currentColor;border-radius:6px;padding:0 6px;margin:2px 3px 2px 0;font-size:18px;white-space:nowrap}
.chip b{font-weight:700}
.M{color:var(--M);background:#fee2e2} .S{color:var(--S)} .C{color:var(--C)} .W{color:var(--W);border-style:dashed}
.box{background:var(--panel);border:1.5px solid var(--line);border-radius:10px;padding:10px 12px}
.box h3{margin:0 0 6px;font-size:22px} .box h4{margin:8px 0 2px;font-size:19px;color:var(--muted)}
.legend{margin-top:14px;font-size:18px;color:var(--muted)} .legend .chip{font-size:17px}
.req{display:block;margin:3px 0;font-size:18px} .req .t{color:var(--ink)}
"""


def chip(r):
    lv = r.get("moscow_final", "?")
    return f'<span class="chip {lv}"><b>{html.escape(r["code"])}</b> [{lv}]</span>'


def req_line(r, show_text=True):
    t = f' <span class="t">{html.escape(r.get("short") or r["text"])[:70]}</span>' if show_text else ""
    return f'<span class="req">{chip(r)}{t}</span>'


def legend():
    items = "".join(f'<span class="chip {k}"><b>{k}</b></span> {v} &nbsp; ' for k, v in L.MOSCOW.items())
    return f'<div class="legend">สัญลักษณ์ระดับ MoSCoW (ระดับฉบับสุดท้าย): {items}</div>'


def page(title, sub, body, extra_css=""):
    return (f"<!doctype html><html lang='th'><head><meta charset='utf-8'><title>{html.escape(title)}</title>"
            f"<style>{CSS}{extra_css}</style></head><body><div class='wrap'><h1>{html.escape(title)}</h1>"
            f"<div class='sub'>{html.escape(sub)}</div>{body}</div></body></html>")


def ipof_part(r):
    ip = r.get("ipof")
    if isinstance(ip, dict) and ip.get("part"):
        return ip["part"], ip.get("sub", "")
    if r["type"] == "FR":
        g = r.get("group")
        if g == "G1":
            return "Input", r.get("data_type", "ชุดข้อมูล")
        if g == "G2":
            return "Process", "การวิเคราะห์"
        if g in ("G3", "G4") and r.get("moscow_final") == "M":
            return "Output", L.GROUPS[g]
    if r["type"] == "NFR" and r.get("layer") == "L6":
        return "Governance", ""
    return None, None


def build_ipof(p):
    reqs = p["requirements"]
    parts = OrderedDict((k, OrderedDict()) for k in ("Input", "Process", "Output", "Feedback"))
    gov, notin = [], []
    proc_default = ["การรับข้อมูลอัตโนมัติ", "การตรวจคุณภาพและที่มา", "การวิเคราะห์"]
    fb_default = ["การแจ้งปัญหาจากผู้ใช้", "การตรวจหลักฐานและแก้ข้อมูล", "การบันทึกประวัติ"]
    for s in proc_default:
        parts["Process"][s] = []
    for s in fb_default:
        parts["Feedback"][s] = []
    for r in reqs:
        part, sub = ipof_part(r)
        if part == "Governance":
            gov.append(r)
        elif part in parts:
            parts[part].setdefault(sub or "", []).append(r)
        else:
            notin.append(r["code"])
    cols = ""
    for name, subs in parts.items():
        inner = ""
        for sub, rs in subs.items():
            items = "".join(req_line(r) for r in rs) or '<span class="req" style="color:var(--muted)">–</span>'
            inner += (f"<h4>{html.escape(sub)}</h4>" if sub else "") + items
        cols += f'<div class="box ipcol"><h3>{name}</h3>{inner}</div>'
    flow = cols
    govh = "".join(chip(r) + html.escape(r.get("short") or r["text"])[:40] + " &nbsp; " for r in gov) or "–"
    body = (f'<div class="gov"><div class="govlab">การกำกับสิทธิ์และการปกป้องข้อมูล (NFR) ครอบทุกองค์ประกอบ: {govh}</div>'
            f'<div class="ipgrid">{flow}</div>'
            f'<div class="fbline">Feedback ↺ ย้อนกลับไปปรับ Input และ Process</div></div>{legend()}')
    css = (".gov{border:3px solid var(--bandline);background:var(--band);border-radius:14px;padding:12px}"
           ".govlab{font-weight:700;color:var(--bandline);margin-bottom:10px}"
           ".ipgrid{display:grid;grid-template-columns:repeat(4,minmax(300px,1fr));gap:22px;position:relative}"
           ".ipcol{background:#fff;position:relative}"
           ".ipcol:not(:last-child)::after{content:'→';position:absolute;right:-21px;top:40%;font-size:30px;color:var(--line)}"
           ".fbline{margin-top:10px;text-align:center;color:var(--muted);border-top:2px dashed var(--line);padding-top:6px}")
    sysname = p["project"].get("system_name", "ระบบ")
    return page(f"กรอบ IPOF ของ{sysname}", "ทุกกล่องระบุรหัสข้อกำหนดและระดับ MoSCoW ฉบับสุดท้าย", body, css), notin


def build_layers(p):
    reqs = p["requirements"]
    by = OrderedDict((k, OrderedDict()) for k in L.LAYERS)
    for r in reqs:
        lay = r.get("layer")
        if lay in by:
            by[lay].setdefault(r.get("component") or "(ยังไม่ระบุองค์ประกอบ)", []).append(r)
    users = p.get("context", {}).get("users", [])
    rows = ""
    for lay in ("L1", "L2", "L3", "L4", "L5", "L7"):
        comps = by[lay]
        inner = ""
        if lay == "L1" and users:
            inner += "".join(f'<div class="box comp"><h4>{html.escape(u)}</h4></div>' for u in users)
        for comp, rs in comps.items():
            inner += f'<div class="box comp"><h4>{html.escape(comp)}</h4>{"".join(chip(r) for r in rs)}</div>'
        rows += (f'<div class="layer"><div class="lab">{lay}<br><span>{L.LAYERS[lay]}</span></div>'
                 f'<div class="comps">{inner or "<span style=color:var(--muted)>–</span>"}</div></div>')
    l6 = "".join(f'<div class="box comp"><h4>{html.escape(c)}</h4>{"".join(chip(r) for r in rs)}</div>' for c, rs in by["L6"].items())
    body = (f'<div class="stack"><div class="layers">{rows}</div>'
            f'<div class="l6"><div class="lab">L6<br><span>{L.LAYERS["L6"]}</span><br><small>กำกับทุกชั้น</small></div>{l6}</div></div>{legend()}')
    css = (".stack{display:grid;grid-template-columns:1fr 300px;gap:14px}"
           ".layer{display:grid;grid-template-columns:170px 1fr;border:1.5px solid var(--line);border-radius:10px;margin-bottom:10px;background:#fff}"
           ".lab{background:#e2e8f0;padding:10px;font-weight:700;font-size:24px;border-radius:9px 0 0 9px}.lab span{font-size:18px;font-weight:400}"
           ".comps{display:flex;flex-wrap:wrap;gap:8px;padding:10px}.comp{min-width:170px;max-width:330px}"
           ".l6{border:3px solid var(--bandline);background:var(--band);border-radius:12px;padding:8px;display:flex;flex-direction:column;gap:8px}"
           ".l6 .lab{background:transparent;color:var(--bandline);border-radius:0}")
    sysname = p["project"].get("system_name", "ระบบ")
    missing = [r["code"] for r in reqs if r.get("layer") not in L.LAYERS]
    no_comp = [r["code"] for r in reqs if r.get("moscow_final") == "M" and not r.get("component")]
    return page(f"สถาปัตยกรรมเชิงแนวคิด 7 ชั้นของ{sysname}", "ข้อกำหนดทุกข้อแสดงพร้อมรหัสและระดับฉบับสุดท้าย", body, css), missing, no_comp


def build_context(p):
    cx = p.get("context", {})
    sysname = p["project"].get("system_name", "ระบบ")
    left = "".join(f'<div class="box">{html.escape(u)}</div>' for u in cx.get("users", []))
    right = "".join(f'<div class="box">{html.escape(s)}</div>' for s in cx.get("sources", []))
    ex = "".join(f'<li>{html.escape(e.get("from",""))} → {html.escape(e.get("to",""))}: {html.escape(e.get("data",""))}</li>' for e in cx.get("exchanges", []))
    ins = "".join(f"<li>{html.escape(x)}</li>" for x in cx.get("in_scope", []))
    outs = "".join(f"<li>{html.escape(x)}</li>" for x in cx.get("out_of_scope", []))
    body = (f'<div class="ctx"><div class="col"><h3>ผู้ใช้งาน</h3>{left}</div><div class="ar">⇄</div>'
            f'<div class="sys box"><h3>{html.escape(sysname)}</h3><div class="sub">ขอบเขตระบบ (ระดับแนวคิด/ต้นแบบ)</div></div>'
            f'<div class="ar">⇄</div><div class="col"><h3>แหล่งข้อมูลภายนอก</h3>{right}</div></div>'
            f'<div class="grid2"><div class="box"><h3>ข้อมูลที่แลกเปลี่ยน</h3><ul>{ex or "<li>–</li>"}</ul></div>'
            f'<div class="box"><h3>ในขอบเขต / นอกขอบเขต</h3><b>ในขอบเขต</b><ul>{ins or "<li>–</li>"}</ul><b>นอกขอบเขต</b><ul>{outs or "<li>–</li>"}</ul></div></div>')
    css = (".ctx{display:grid;grid-template-columns:280px 50px 1fr 50px 280px;align-items:center;gap:6px;margin-bottom:16px}"
           ".col{display:flex;flex-direction:column;gap:8px}.ar{font-size:34px;text-align:center;color:var(--line)}"
           ".sys{border:3px solid var(--ink);min-height:200px;display:flex;flex-direction:column;justify-content:center;text-align:center}"
           ".grid2{display:grid;grid-template-columns:1fr 1fr;gap:14px}ul{margin:4px 0 8px 18px;padding:0}")
    return page(f"ภาพบริบทของ{sysname}", "ผู้ใช้ แหล่งข้อมูลภายนอก ข้อมูลที่แลกเปลี่ยน และขอบเขต", body, css)


DEFAULT_FLOW = ["Source", "ตัวรวบรวมรายศูนย์พร้อมบัฟเฟอร์", "Validation", "Transformation", "Store", "Processing", "Output", "Feedback/Logging"]


def build_dataflow(p):
    reqs = p["requirements"]
    stages = p.get("dataflow", {}).get("stages", DEFAULT_FLOW)
    notes = p.get("dataflow", {}).get("notes", {})
    by = {s: [r for r in reqs if r.get("flow_stage") == s] for s in stages}
    cells = ""
    for i, s in enumerate(stages):
        cells += (f'<div class="box st"><h3>{i + 1}. {html.escape(s)}</h3><div class="sub">{html.escape(notes.get(s, ""))}</div>'
                  f'{"".join(chip(r) for r in by[s])}</div>')
        if i < len(stages) - 1:
            cells += '<div class="ar">→</div>'
    css = ".flow{display:flex;align-items:stretch;gap:4px}.st{width:190px}.ar{align-self:center;font-size:30px;color:var(--line)}"
    sysname = p["project"].get("system_name", "ระบบ")
    return page(f"การไหลของข้อมูลใน{sysname}", "ลำดับการไหลจากแหล่งข้อมูลถึงการบันทึกประวัติ", f'<div class="flow">{cells}</div>{legend()}', css)


def build_trace(p, outdir):
    reqs = {r["code"]: r for r in p["requirements"]}
    rows = []
    for r in p["requirements"]:
        if r.get("moscow_final") != "M":
            continue
        rows.append([r["code"], r.get("component", ""), f'{r.get("layer", "")} {L.LAYERS.get(r.get("layer"), "")}',
                     ", ".join(f'{d} [{reqs[d].get("moscow_final", "?")}]' for d in r.get("depends_on", []) if d in reqs) or "–",
                     ", ".join(r.get("tc", [])) or "–"])
    hdr = ["ข้อกำหนด Must", "องค์ประกอบ", "ชั้น", "ข้อกำหนดที่พึ่งพา", "กรณีทดสอบ"]
    with open(Path(outdir) / "traceability.csv", "w", newline="", encoding="utf-8-sig") as f:
        w = csv.writer(f)
        w.writerow(hdr)
        w.writerows(rows)
    t = "<table><tr>" + "".join(f"<th>{h}</th>" for h in hdr) + "</tr>" + "".join(
        "<tr>" + "".join(f"<td>{html.escape(x)}</td>" for x in row) + "</tr>" for row in rows) + "</table>"
    css = "table{border-collapse:collapse}td,th{border:1px solid var(--line);padding:4px 10px;text-align:left}th{background:#e2e8f0}"
    return page("ตารางตรวจสอบย้อนกลับ", "Must → องค์ประกอบ/ชั้น → ข้อกำหนดที่พึ่งพา → กรณีทดสอบ", t, css)


def to_png(html_path, png_path, scale):
    try:
        from playwright.sync_api import sync_playwright
        with sync_playwright() as pw:
            b = pw.chromium.launch()
            pg = b.new_page(device_scale_factor=scale, viewport={"width": 1400, "height": 900})
            pg.goto(Path(html_path).resolve().as_uri())
            pg.wait_for_timeout(300)
            pg.locator(".wrap").screenshot(path=str(png_path))
            b.close()
        return "playwright"
    except Exception as e:  # noqa: BLE001
        exe = shutil.which("wkhtmltoimage")
        if exe:
            subprocess.run([exe, "--zoom", str(scale), "--quiet", str(html_path), str(png_path)], timeout=120)
            return "wkhtmltoimage"
        print(f"⚠ render PNG ไม่ได้ ({e}) — ส่ง HTML แล้วให้ผู้ใช้เปิดในเบราว์เซอร์")
        return None


def main(argv=None):
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("--project", required=True)
    ap.add_argument("--outdir", required=True)
    ap.add_argument("--only", default="ipof,layers,context,dataflow,trace")
    ap.add_argument("--scale", type=int, default=3)
    ap.add_argument("--no-png", action="store_true")
    a = ap.parse_args(argv)
    p = L.load_project(a.project)
    od = Path(a.outdir)
    od.mkdir(parents=True, exist_ok=True)
    want = set(a.only.split(","))
    out = {}
    problems = []
    if "ipof" in want:
        h, notin = build_ipof(p)
        out["ipof"] = h
        if notin:
            print(f"ข้อกำหนดที่ไม่อยู่ในกรอบ IPOF (ตามกติกา เช่น G3/G4 ที่ไม่ใช่ Must): {notin}")
    if "layers" in want:
        h, missing, no_comp = build_layers(p)
        out["layers7"] = h
        if missing:
            problems.append(f"ข้อกำหนดที่ไม่อยู่ในภาพ 7 ชั้น: {missing}")
        if no_comp:
            problems.append(f"ข้อ Must ที่ไม่มีองค์ประกอบรับผิดชอบ: {no_comp}")
    if "context" in want:
        out["context"] = build_context(p)
    if "dataflow" in want:
        out["dataflow"] = build_dataflow(p)
    if "trace" in want:
        out["traceability"] = build_trace(p, od)
    for name, h in out.items():
        h = h.replace(" - ", " – ")
        hp = od / f"{name}.html"
        hp.write_text(h, encoding="utf-8")
        eng = None if a.no_png else to_png(hp, od / f"{name}.png", a.scale)
        print(f"{name}: {hp}" + (f" + PNG ({eng})" if eng else ""))
    for pr in problems:
        print("⚠", pr)
    return 1 if problems else 0


if __name__ == "__main__":
    sys.exit(main())
