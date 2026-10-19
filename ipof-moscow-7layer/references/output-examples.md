# ตัวอย่างโครงสร้างไฟล์ผลลัพธ์

ไฟล์ตัวอย่างจริงอยู่ใน `assets/examples/` (ข้อมูลสมมุติ ห้ามใช้เป็นผลวิจัย)

## 1. ผลอ่านแบบประเมิน (`read_forms.py OUTDIR forms/`)

`scores.csv`
```
respondent,item,value
U01,P01,4
U01,N01,5
E01,FR01,5            ← บรรทัด "คะแนนความเหมาะสม"
E01,FR06#agree,disagree
E01,FR06#moscow,M     ← ระดับที่เสนอ (เฉพาะผู้ไม่เห็นชอบ)
```
`choices.csv` (respondent, question, option, checked) · `text.csv` (respondent, item, column, text) · `issues.csv` (แถวไม่ติ๊ก/ติ๊กเกิน — ต้องตรวจต้นฉบับ)

## 2. ทะเบียนโครงการ

`assets/examples/project.example.json` — project, dimensions, problems, needs, functions, open_questions, moscow_cutoffs, requirements, constraints, context, round1, arch_items, eval_items, confirm_items, round2

## 3. Excel

`assets/examples/summary_user_example.xlsx` — แผ่น "ค่าเฉลี่ยรายด้าน":

| รหัส | รายการ | n | X̄ | S.D. | แปลผล |
|---|---|---|---|---|---|
| N01 | … | `=COUNT(ช่วง)` | `=AVERAGE(ช่วง)` | `=STDEV.S(ช่วง)` | `=IF(…)` |
| | **รวมด้านข้อมูล** (พื้นสี) | `=COUNT(ช่วง1,ช่วง2,…)` | `=AVERAGE(ช่วง1,ช่วง2,…)` | `=STDEV.S(…)` | |

คู่กับ `<ไฟล์>.manifest.json` สำหรับ `verify_excel.py`

## 4. แผนภาพ

`assets/examples/ipof.png`, `assets/examples/layers7.png` (+ `.html` แก้ได้ ใน OUTDIR ของงานจริง), `context`, `dataflow`, `traceability.csv`

## 5. คำสั่งแก้ .docx (`edits.json`)

```json
[
  {"find": "ค่าเฉลี่ย 4.21", "replace": "ค่าเฉลี่ย 4.25", "count": 1},
  {"table": 2, "row": 3, "col": 4, "value": "4.25"},
  {"paragraph_contains": "สรุปผลการวิจัย", "new_text": "…ย่อหน้าใหม่…"}
]
```
ผลลัพธ์ `บทความ_V3.docx` จาก `บทความ_V2.docx`

## 6. ชุดไฟล์ส่งมอบต่อรุ่น

```
บทความ_V3.docx                     ← บทความ (ตาราง 1–5)
เอกสารประกอบเล่ม_V3.docx           ← ตาราง 7 ตาราง + รูป IPOF + ผลรายครั้ง + เกณฑ์ยอมรับทุกระดับ
สรุปผลผู้ใช้_V2.xlsx
สรุปผลผู้เชี่ยวชาญครั้งที่1_V2.xlsx
สรุปผลผู้เชี่ยวชาญครั้งที่2_V2.xlsx
ตรวจเทียบเอกสารอ้างอิง_V3.md
แผนภาพ/ipof.png layers7.png context.png dataflow.png (+ .html)
```
ก่อนสร้างไฟล์ .docx ให้โหลด `jdtaia-ieee-document` / `thai-docx` / `docx` และไฟล์ .xlsx โหลด `xlsx` ตามที่ติดตั้ง (SKILL.md หัวข้อ "ส่งออกเอกสาร")
แจ้งผู้ใช้ให้บันทึกทุกไฟล์ไว้ในโฟลเดอร์โปรเจกต์
