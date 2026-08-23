---
description: Apply thai-docx formatting rules when creating/editing a Thai .docx file
---

ใช้ skill `thai-docx` (~/.claude/skills/thai-docx/SKILL.md) เมื่อสร้าง/แก้ไฟล์ .docx ที่มีภาษาไทย โดยเฉพาะ:

1. ตั้งฟอนต์ให้ชัดเจนที่ document defaults: TH Sarabun New (หรือ TH Sarabun PSK / Angsana New) — ห้ามปล่อยให้เป็นค่า default ของ docx-js
2. ห้ามใช้ `\n` ในข้อความ — แยกเป็น Paragraph แต่ละอันเสมอ
3. ระบุ page size และ margin อย่างชัดเจน (DXA) ตามเทมเพลตที่กำหนด ไม่ปล่อยให้เป็นค่า default
4. ตารางต้องมี columnWidths ที่ตาราง + width แบบ WidthType.DXA ที่ทุกเซลล์ ผลรวม columnWidths ต้องเท่ากับความกว้างของตาราง ห้ามใช้ WidthType.PERCENTAGE
5. กำหนด spacing (line/after) ของ Paragraph อย่างชัดเจนสำหรับเอกสารไทย เพื่อป้องกันสระ/วรรณยุกต์ชนกับบรรทัดข้างบน
6. ก่อนส่งไฟล์ให้ผู้ใช้ทุกครั้ง: render เป็น PDF แล้วแปลงเป็น JPEG ด้วย
   ```
   python scripts/office/soffice.py --headless --convert-to pdf output.docx
   pdftoppm -jpeg -r 100 output.pdf page
   ```
   แล้วตรวจดูภาพจริงเพื่อตรวจสอบ: ตำแหน่งสระ/วรรณยุกต์, ความกว้างตาราง, ระยะขอบกระดาษทุกหน้า, ระยะห่างบรรทัด — ก่อนส่งไฟล์จริง

ถ้าพบความผิดเพี้ยน ให้ตรวจ 3 จุดที่มักเป็นต้นตอก่อน: (1) ฟอนต์ไม่ได้ตั้งค่าชัดเจน (2) มี `\n` หลงเหลืออยู่ (3) ตารางใช้ PERCENTAGE แทน DXA
