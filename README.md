# claude-skills

ชุด skill สำหรับ [Claude Code](https://code.claude.com) ของ jatupornrmutsb

- skill งานวิศวกรรมซอฟต์แวร์ในโฟลเดอร์ [`skills/`](./skills) นำมาจาก [utarn/engineer-skills](https://github.com/utarn/engineer-skills) (ดู[เครดิต](#เครดิตและสัญญาอนุญาต))
- skill ที่เขียนเอง: [`thai-docx/`](./thai-docx) และ slash command ใน [`commands/`](./commands)

## วิธีใช้

พิมพ์ `/ชื่อคำสั่ง` ใน Claude Code แล้วตามด้วยสิ่งที่ต้องการ เช่น `/tdd เพิ่มฟังก์ชันคำนวณเกรด`

- **⚡ อัตโนมัติ** = Claude เลือกใช้ skill นี้เองได้เมื่อเห็นว่าเข้ากับงาน (สั่งเองด้วย `/` ก็ได้)
- skill ที่ไม่มีเครื่องหมายต้องพิมพ์ `/คำสั่ง` เอง
- **🧪 ทดลอง** = ต้นฉบับยังพัฒนาอยู่ (โฟลเดอร์ `skills/in-progress/`) อาจเปลี่ยนบ่อย
- คำสั่งที่ลงท้ายด้วย `-yes` ทำงานเหมือนคำสั่งหลัก แต่ไม่หยุดถามยืนยัน ทำต่อจนจบเลย

> ไม่แน่ใจว่าจะใช้คำสั่งไหน ให้เริ่มจาก `/ask-utarn` แล้วเล่าว่ากำลังจะทำอะไร

### 1. เริ่มงาน

| คำสั่ง | ใช้ทำอะไร | ตัวอย่างการสั่ง |
|---|---|---|
| `/ask-utarn` | ช่วยแนะนำว่าควรใช้คำสั่งไหน | `/ask-utarn อยากเพิ่มระบบล็อกอิน ควรเริ่มยังไง` |
| `/setup-utarn-skills` | ตั้งค่า repo ครั้งแรก (ที่เก็บ issue, ป้ายกำกับ, โฟลเดอร์เอกสาร) ก่อนใช้คำสั่งอื่น | `/setup-utarn-skills` |
| `/setup-utarn-skills-yes` | เหมือนข้างบน แต่ใช้ค่าแนะนำโดยไม่ถาม | `/setup-utarn-skills-yes` |
| `/init-bug-tracker` | สร้างชุดป้ายกำกับบน GitHub repo ที่ใช้รับแจ้งบั๊ก | `/init-bug-tracker myorg/myapp-bugs` |
| `/git-guardrails-claude-code` ⚡ | ใส่ hook กันไม่ให้ Claude รันคำสั่ง git อันตราย (push, reset --hard ฯลฯ) | `/git-guardrails-claude-code` |
| `/setup-pre-commit` ⚡ | ติดตั้ง Husky ให้จัดรูปแบบโค้ด ตรวจชนิดข้อมูล และรันเทสต์ก่อน commit | `/setup-pre-commit` |
| `/setup-ts-deep-modules` 🧪 | ตั้งค่า dependency-cruiser ให้แต่ละแพ็กเกจ TypeScript เปิดให้เรียกได้แค่ทางไฟล์หลัก | `/setup-ts-deep-modules` |

### 2. วางแผน

| คำสั่ง | ใช้ทำอะไร | ตัวอย่างการสั่ง |
|---|---|---|
| `/grill-me` | ให้ Claude ซักถามแผนของเราจนกว่าจะชัด | `/grill-me ฉันจะทำระบบจองห้องประชุม` |
| `/grilling` ⚡ | เหมือน grill-me แต่ Claude เรียกใช้เองได้ | `ช่วย grill แนวคิดนี้หน่อย` |
| `/grill-novice` | ถามเฉพาะความต้องการด้วยภาษาง่ายๆ ส่วนเรื่องเทคนิค Claude ตัดสินใจเองแล้วสรุปให้เราตรวจ | `/grill-novice อยากได้เว็บเช็กชื่อนักศึกษา` |
| `/grill-with-docs` | ซักถามแผนพร้อมบันทึกการตัดสินใจ (ADR) และอภิธานศัพท์ไปด้วย | `/grill-with-docs ออกแบบระบบสิทธิ์ผู้ใช้` |
| `/domain-modeling` ⚡ | กำหนดคำศัพท์และแนวคิดหลักของระบบให้ใช้ตรงกันทั้งโปรเจกต์ | `/domain-modeling คำว่า "รายวิชา" กับ "กลุ่มเรียน" ต่างกันยังไง` |
| `/codebase-design` ⚡ | ออกแบบโมดูลให้ใช้ง่าย ซ่อนรายละเอียด และทดสอบง่าย | `/codebase-design ออกแบบโมดูลส่งอีเมล` |
| `/prototype` ⚡ | ทำต้นแบบทิ้งได้ เพื่อลองว่าแนวคิดหรือหน้าจอใช้ได้จริงไหม | `/prototype ลองหน้าจอแดชบอร์ดผลการเรียน 3 แบบ` |
| `/research` ⚡ | ค้นข้อมูลจากแหล่งที่เชื่อถือได้ แล้วสรุปเป็นไฟล์ Markdown | `/research วิธีใช้ OAuth กับ Google Workspace` |
| `/to-prd` | สรุปสิ่งที่คุยกันเป็นเอกสารความต้องการ (PRD) แล้วโพสต์ลงที่เก็บ issue | `/to-prd` |
| `/to-spec` | สรุปสิ่งที่คุยกันเป็นสเปกงาน แล้วโพสต์ลงที่เก็บ issue | `/to-spec` |
| `/to-spec-yes` | เหมือน to-spec แต่ไม่ถามยืนยัน | `/to-spec-yes` |
| `/to-issues` | แตกแผนหรือสเปกเป็น issue ย่อยที่หยิบไปทำแยกกันได้ | `/to-issues จากสเปก #12` |
| `/to-issues-yes` | เหมือน to-issues แต่ไม่ถามยืนยัน | `/to-issues-yes จากสเปก #12` |
| `/to-tickets` | แตกงานเป็น ticket พร้อมระบุว่า ticket ไหนต้องรออันไหน | `/to-tickets` |
| `/to-tickets-yes` | เหมือน to-tickets แต่ไม่ถามยืนยัน | `/to-tickets-yes` |
| `/wayfinder` | วางแผนงานใหญ่ที่ทำใน session เดียวไม่จบ เป็นแผนที่ของเรื่องที่ต้องตัดสินใจ | `/wayfinder ย้ายระบบจาก PHP ไป .NET` |
| `/to-questionnaire` | เปลี่ยนเรื่องที่เราตอบเองไม่ได้ เป็นแบบสอบถามให้คนอื่นกรอก | `/to-questionnaire ถามฝ่ายทะเบียนเรื่องเงื่อนไขการลงทะเบียน` |
| `/loop-me` 🧪 | ซักถามเพื่อออกแบบ workflow ที่อยากสร้าง | `/loop-me` |

### 3. เขียนโค้ดและทดสอบ

| คำสั่ง | ใช้ทำอะไร | ตัวอย่างการสั่ง |
|---|---|---|
| `/implement` | ลงมือเขียนโค้ดตามสเปกหรือ ticket | `/implement #15` |
| `/tdd` ⚡ | เขียนเทสต์ก่อนแล้วค่อยเขียนโค้ด (red-green-refactor) | `/tdd เพิ่มฟังก์ชันคำนวณ GPA` |
| `/work-on-issues` ⚡ | ดึง issue จาก GitHub/GitLab มาทำทีละเรื่องจนปิด | `/work-on-issues` |
| `/improve-codebase-architecture` | สแกนโค้ดหาจุดปรับโครงสร้าง ทำรายงาน HTML แล้วซักถามจุดที่เราเลือก | `/improve-codebase-architecture` |
| `/migrate-to-shoehorn` ⚡ | เปลี่ยน `as` ในไฟล์เทสต์ TypeScript ไปใช้ shoehorn | `/migrate-to-shoehorn src/**/*.test.ts` |
| `/scaffold-exercises` ⚡ | สร้างโครงโฟลเดอร์แบบฝึกหัด (โจทย์ เฉลย คำอธิบาย) | `/scaffold-exercises บทที่ 3 มี 5 ข้อ` |
| `/wizard` 🧪 | สร้างสคริปต์ bash ที่พาคนทำขั้นตอนตั้งค่าทีละขั้น และเขียนไฟล์ .env ให้ | `/wizard ตั้งค่า Google OAuth client` |

### 4. ตรวจและแก้บั๊ก

| คำสั่ง | ใช้ทำอะไร | ตัวอย่างการสั่ง |
|---|---|---|
| `/code-review` ⚡ | ตรวจโค้ดที่เปลี่ยนว่าตรงมาตรฐานของ repo และตรงสเปกไหม | `/code-review main` |
| `/find-mismatch` ⚡ | หาบั๊กที่จะพังตอนรัน เช่น ข้อมูลส่งกันไม่ตรง หรือ async ผิด | `/find-mismatch` |
| `/diagnosing-bugs` ⚡ | ไล่หาสาเหตุบั๊กยากหรือระบบช้าอย่างเป็นขั้นตอน | `/diagnosing-bugs หน้า login ค้างหลัง deploy` |
| `/resolving-merge-conflicts` ⚡ | แก้ conflict ตอน merge หรือ rebase | `/resolving-merge-conflicts` |
| `/triage` | คัดแยก issue และ PR จากภายนอก แล้วเขียนสรุปให้พร้อมส่งต่อ | `/triage` |
| `/bug-tracker` | ไล่ทำบั๊กใน GitHub bug tracker ตั้งแต่ตรวจสอบจนปิด | `/bug-tracker` |

> ⚠️ Claude Code มีคำสั่ง `/code-review` ในตัวอยู่แล้ว ถ้าติดตั้ง skill นี้ ชื่ออาจทับกัน ให้ตรวจดูว่าตัวไหนทำงานเมื่อพิมพ์คำสั่ง

### 5. งานทั่วไป

| คำสั่ง | ใช้ทำอะไร | ตัวอย่างการสั่ง |
|---|---|---|
| `/handoff` | สรุปบทสนทนาเป็นเอกสารส่งต่อให้ session ถัดไป | `/handoff จะทำต่อเรื่องหน้ารายงาน` |
| `/claude-handoff` 🧪 | ส่งงานต่อให้ agent ใหม่ที่รันเบื้องหลังทำต่อทันที | `/claude-handoff เขียนเทสต์ที่เหลือ` |
| `/teach` | ให้ Claude สอนเรื่องใหม่โดยใช้โค้ดใน workspace เป็นตัวอย่าง | `/teach Entity Framework migration` |
| `/wait-what` | ข้อความล่าสุดของ Claude อ่านไม่เข้าใจ ให้อธิบายใหม่ | `/wait-what` |
| `/writing-for-agents` ⚡ | หลักการเขียนเอกสารให้ AI อ่าน เช่น SKILL.md, CLAUDE.md | `/writing-for-agents ช่วยปรับ CLAUDE.md นี้` |
| `/edit-article` | ปรับบทความให้อ่านง่าย จัดโครงสร้าง และกระชับขึ้น | `/edit-article draft.md` |
| `/obsidian-vault` ⚡ | ค้น สร้าง และจัดโน้ตใน Obsidian | `/obsidian-vault หาโน้ตเรื่อง IDSS` |
| `/writing-fragments` 🧪 | เขียนขั้นที่ 1: เก็บไอเดียเป็นชิ้นๆ ยังไม่จัดโครง | `/writing-fragments บทความเรื่อง AI ในมหาวิทยาลัย` |
| `/writing-beats` 🧪 | เขียนขั้นที่ 2: เรียงไอเดียเป็นลำดับเรื่อง | `/writing-beats` |
| `/writing-shape` 🧪 | เขียนขั้นที่ 3: เรียบเรียงเป็นบทความทีละย่อหน้า | `/writing-shape` |
| `/thai-docx` ⚡ | สร้างหรือแก้ไฟล์ Word ภาษาไทยให้ฟอนต์ สระ และตารางไม่เพี้ยน | `/thai-docx ทำบันทึกข้อความขออนุมัติโครงการ` |

## ติดตั้งบนเครื่องอื่น

สคริปต์จะคัดลอก skill ทุกตัวไปไว้ที่ `~/.claude/skills/` (แบบแบน ไม่มีโฟลเดอร์หมวด) และ slash command ไปไว้ที่ `~/.claude/commands/`
skill ชื่อเดียวกันที่มีอยู่แล้วในเครื่องจะถูกแทนที่ด้วยเวอร์ชันจาก repo นี้ ส่วน skill อื่นที่ไม่ได้มาจาก repo นี้จะไม่ถูกแตะ

สิ่งที่ต้องมี: [Git](https://git-scm.com) และ [Claude Code](https://code.claude.com)

### Windows (PowerShell)

```powershell
git clone https://github.com/jatupornrmutsb/claude-skills.git
cd claude-skills
powershell -ExecutionPolicy Bypass -File .\install.ps1
```

หรือติดตั้งโดยไม่ต้อง clone เอง (repo เป็น Private จึงต้องใช้ [GitHub CLI](https://cli.github.com/) ที่ `gh auth login` แล้ว):

```powershell
$t = Join-Path $env:TEMP ("cs-" + [guid]::NewGuid()); gh repo clone jatupornrmutsb/claude-skills $t -- --depth 1 --quiet; & "$t\install.ps1"; Remove-Item -Recurse -Force $t
```

### Linux / WSL

```bash
git clone https://github.com/jatupornrmutsb/claude-skills.git
cd claude-skills
./install.sh
```

หรือติดตั้งโดยไม่ต้อง clone เอง (repo เป็น Private จึงต้องใช้ [GitHub CLI](https://cli.github.com/) ที่ `gh auth login` แล้ว):

```bash
t=$(mktemp -d) && gh repo clone jatupornrmutsb/claude-skills "$t" -- --depth 1 --quiet && "$t/install.sh"; rm -rf "$t"
```

> **WSL:** WSL กับ Windows มี `~/.claude` แยกกัน ถ้าใช้ Claude Code ทั้งสองฝั่ง ต้องติดตั้งทั้งสองฝั่ง

ถ้าไม่ต้องการ skill กลุ่มทดลอง (🧪) ให้ใส่ `-SkipInProgress` (Windows) หรือ `--skip-in-progress` (Linux/WSL)
ติดตั้งเสร็จแล้วให้เปิด Claude Code ใหม่ แล้วพิมพ์ `/` เพื่อดูว่ามีคำสั่งขึ้นมาแล้ว

### อัปเดตเครื่องที่ติดตั้งไว้แล้ว

```bash
cd claude-skills
git pull
./install.sh          # Windows: .\install.ps1
```

## อัปเดตจากต้นฉบับ (utarn/engineer-skills)

repo นี้มี remote ชื่อ `upstream` ชี้ไปที่ต้นฉบับ ถ้า clone ใหม่ต้องเพิ่มก่อน 1 ครั้ง:

```bash
git remote add upstream https://github.com/utarn/engineer-skills.git
```

ดึงเวอร์ชันใหม่มาแทนโฟลเดอร์ `skills/` (ไฟล์อื่นของเราไม่ถูกแตะ):

```bash
git fetch upstream
git diff --stat HEAD upstream/main -- skills LICENSE   # ดูก่อนว่าอะไรเปลี่ยน
git rm -r -q skills                                     # ลบของเก่าออก เพื่อให้ skill ที่ต้นฉบับลบไปแล้วหายไปด้วย
git checkout upstream/main -- skills LICENSE
git commit -m "อัปเดต skills จาก utarn/engineer-skills"
git push
```

จากนั้นรันสคริปต์ติดตั้งอีกครั้งเพื่ออัปเดต skill ในเครื่อง
ถ้าต้นฉบับเพิ่ม skill ใหม่หรือย้ายหมวด ให้ปรับตารางใน README นี้ให้ตรงด้วย

> อย่าแก้ไฟล์ใน `skills/` โดยตรง เพราะการอัปเดตแบบข้างบนจะเขียนทับ ถ้าจะปรับ skill ของต้นฉบับ ให้คัดลอกออกมาเป็นโฟลเดอร์ใหม่ที่ root ของ repo (แบบ `thai-docx/`) แล้วตั้งชื่อใหม่

## เครดิตและสัญญาอนุญาต

- skill ใน `skills/` เขียนโดย **Utharn Buranasaksee** — [utarn/engineer-skills](https://github.com/utarn/engineer-skills)
- ซึ่ง fork มาจาก **Matt Pocock** — [mattpocock/skills](https://github.com/mattpocock/skills)
- ทั้งสอง repo ใช้สัญญาอนุญาต MIT ไฟล์ [LICENSE](./LICENSE) ใน repo นี้คัดลอกมาจากต้นฉบับโดยไม่แก้ไข
