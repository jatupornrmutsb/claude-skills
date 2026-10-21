#!/usr/bin/env node
// make_memo.js — สร้าง "บันทึกข้อความ" (หนังสือภายใน) .docx ตามแบบหน่วยงาน ด้วย docx-js
// ใช้: node make_memo.js memo.json out.docx
// หลักการ: ไม่พึ่งรูปทรงลอย/ช่องว่าง/ปุ่ม Tab — ใช้ tab stop + indent + dotted underline เท่านั้น
const fs = require('fs');
const path = require('path');
const {
  Document, Packer, Paragraph, TextRun, ImageRun, Footer, AlignmentType, TabStopType,
  UnderlineType, HorizontalPositionRelativeFrom, VerticalPositionRelativeFrom,
  TextWrappingType, Tab, LeaderType, LineRuleType,
} = require('docx');

const FONT = 'TH SarabunPSK';
const FONT_ALL = { ascii: FONT, hAnsi: FONT, cs: FONT, eastAsia: FONT };
const PAGE = { width: 11907, height: 16840 };            // A4
const MARGIN = 1440;                                       // 1 นิ้ว ตามแบบหน่วยงาน
const TEXT_W = PAGE.width - 2 * MARGIN;                    // 9027 DXA
// รูปอยู่ที่ <skill>/assets (ข้างๆ scripts/) — รองรับทั้ง ../assets และ ./assets
const ASSETS = [path.join(__dirname, '..', 'assets'), path.join(__dirname, 'assets')]
  .find((d) => fs.existsSync(path.join(d, 'garuda.png'))) || path.join(__dirname, '..', 'assets');

// ---------- helpers: ป้องกันตัวอักษรไทยเพี้ยน ----------
const THAI_DIGITS = '๐๑๒๓๔๕๖๗๘๙';
const toThaiDigits = (s) => String(s).replace(/[0-9]/g, (d) => THAI_DIGITS[d]);
function normalizeThai(s) {
  return String(s)
    .normalize('NFC')
    .replace(/ํา/g, 'ำ')   // นิคหิต+สระอา (จากการ copy PDF) -> สระอำ
    .replace(/[ ​‌‍﻿]/g, ' ') // NBSP/zero-width -> เว้นวรรคปกติ
    .replace(/[ \t]{2,}/g, ' ')            // เว้นวรรคซ้อน -> 1
    .trim();
}
const clean = (s) => toThaiDigits(normalizeThai(s));

// ---------- run / paragraph builders ----------
const run = (text, o = {}) =>
  new TextRun({ text, font: FONT_ALL, size: o.size || 32, bold: !!o.bold, underline: o.underline });
const tabRun = () => new TextRun({ children: [new Tab()], font: FONT_ALL, size: 32 });

const label = (t) => run(t, { size: 40, bold: true });          // ส่วนราชการ/ที่/วันที่/เรื่อง = 20pt หนา
// keep=true: คงช่องว่างหน้า-หลังไว้ (ใช้กับ "วันที่" ที่เว้นเลขวันไว้เขียนมือ)
const value = (t, keep) => run(keep
  ? toThaiDigits(String(t).normalize('NFC').replace(/\u0e4d\u0e32/g, '\u0e33'))
  : clean(t));

const DOT_END = { type: TabStopType.RIGHT, position: TEXT_W, leader: LeaderType.DOT };   // จุดไข่ปลาถึงขอบขวา
const DOT_MID = (pos) => ({ type: TabStopType.LEFT, position: pos, leader: LeaderType.DOT });

function headerRows(m) {
  return [
    // ส่วนราชการ ........
    new Paragraph({
      tabStops: [{ type: TabStopType.LEFT, position: 1474 }, DOT_END],
      children: [label('ส่วนราชการ'), tabRun(), value(m.dept), tabRun()],
    }),
    // ที่ ........ วันที่ ........
    new Paragraph({
      tabStops: [
        { type: TabStopType.LEFT, position: 400 }, DOT_MID(4253),
        { type: TabStopType.LEFT, position: 5130 }, DOT_END,
      ],
      children: [
        label('ที่'), tabRun(), value(m.no), tabRun(),
        label('วันที่'), tabRun(), value(m.date, true), tabRun(),
      ],
    }),
    // เรื่อง ........ (บรรทัดต่อไปชิดตำแหน่งเดียวกับค่า ด้วย hanging indent)
    new Paragraph({
      alignment: AlignmentType.THAI_DISTRIBUTE,
      indent: { left: 680, hanging: 680 },
      tabStops: [{ type: TabStopType.LEFT, position: 680 }, DOT_END],
      children: [label('เรื่อง'), tabRun(), value(m.subject), tabRun()],
    }),
  ];
}

function bodyParas(m, first) {
  const out = [];
  // เรียน
  out.push(new Paragraph({
    spacing: { before: 120 },
    tabStops: [{ type: TabStopType.LEFT, position: 600 }],
    children: [run('เรียน'), tabRun(), run(clean(m.to))],
  }));
  // ภาคเหตุ / ภาคความประสงค์ (ปกติ 2 ย่อหน้า) — ย่อหน้าด้วย firstLine ไม่ใช้ปุ่ม Tab
  (m.paras || []).forEach((p) =>
    out.push(new Paragraph({
      alignment: AlignmentType.THAI_DISTRIBUTE,
      spacing: { before: 120 },
      indent: { firstLine: m.indent || 1440 },
      children: [run(clean(p))],
    })));
  // รายการเอกสารแนบ (ถ้ามี) : เลขไทย + hanging indent
  (m.items || []).forEach((it, i) =>
    out.push(new Paragraph({
      alignment: AlignmentType.THAI_DISTRIBUTE,
      indent: { left: 1985, hanging: 545 },
      tabStops: [{ type: TabStopType.LEFT, position: 1985 }],
      children: [run(toThaiDigits(i + 1) + '.'), tabRun(), run(clean(it))],
    })));
  // คำลงท้าย (keepNext กันลายเซ็นหลุดไปหน้าใหม่ตามลำพัง)
  out.push(new Paragraph({
    spacing: { before: 120 }, indent: { firstLine: m.indent || 1440 }, keepNext: true, keepLines: true,
    children: [run(clean(m.closing))],
  }));
  // ลายเซ็น: เว้น 3 บรรทัด แล้วจัดชื่อ/ตำแหน่งกึ่งกลางบล็อกเดียวกัน (ไม่ใช้เว้นวรรค)
  out.push(new Paragraph({
    alignment: AlignmentType.CENTER, spacing: { before: 1200 }, indent: { left: 1700 },
    keepNext: true, keepLines: true, children: [run(clean(m.name))],
  }));
  out.push(new Paragraph({
    alignment: AlignmentType.CENTER, indent: { left: 1700 }, keepLines: true,
    children: [run(clean(m.position))],
  }));
  return out;
}

function memoBlock(m, idx) {
  const garuda = fs.readFileSync(path.join(ASSETS, 'garuda.png'));
  const title = new Paragraph({
    alignment: AlignmentType.CENTER,
    pageBreakBefore: idx > 0,
    children: [
      // ตราครุฑ 2.19 x 1.98 ซม. ลอยแบบตำแหน่งสัมบูรณ์เทียบกับหน้า (เหมือนต้นแบบ) — 1 รูปเท่านั้น
      new ImageRun({
        type: 'png', data: garuda,
        transformation: { width: 83, height: 75 },
        altText: { title: 'ตราครุฑ', description: 'ตราครุฑ', name: 'garuda' },
        floating: {
          horizontalPosition: { relative: HorizontalPositionRelativeFrom.PAGE, offset: 922020 },
          verticalPosition: { relative: VerticalPositionRelativeFrom.PAGE, offset: 584835 },
          wrap: { type: TextWrappingType.NONE }, allowOverlap: true, behindDocument: false,
        },
      }),
      run('บันทึกข้อความ', { size: 56, bold: true }),
    ],
  });
  return [title, ...headerRows(m), ...bodyParas(m, idx === 0)];
}

function build(spec) {
  const footerImg = fs.readFileSync(path.join(ASSETS, 'footer_slogan.png'));
  const children = spec.memos.flatMap(memoBlock);
  return new Document({
    creator: spec.creator || '', title: spec.title || 'บันทึกข้อความ',
    styles: {
      default: {
        document: {
          run: { font: FONT_ALL, size: 32, language: { value: 'en-US', eastAsia: 'en-US', bidirectional: 'th-TH' } },
          paragraph: { spacing: { before: 0, after: 0, line: 240, lineRule: LineRuleType.AUTO } },   // single — ต้องระบุ lineRule เสมอ
        },
      },
    },
    sections: [{
      properties: {
        page: {
          size: { width: PAGE.width, height: PAGE.height },
          margin: { top: MARGIN, right: MARGIN, bottom: MARGIN, left: MARGIN, header: 720, footer: 720 },
        },
      },
      footers: {
        default: new Footer({
          children: [new Paragraph({
            alignment: AlignmentType.CENTER,
            children: [new ImageRun({
              type: 'png', data: footerImg, transformation: { width: 221, height: 41 },
              altText: { title: 'คำขวัญ', description: 'คำขวัญหน่วยงาน', name: 'slogan' },
            })],
          })],
        }),
      },
      children,
    }],
  });
}

if (require.main === module) {
  const [specPath, outPath] = process.argv.slice(2);
  if (!specPath || !outPath) { console.error('usage: node make_memo.js memo.json out.docx'); process.exit(1); }
  const spec = JSON.parse(fs.readFileSync(specPath, 'utf8'));
  Packer.toBuffer(build(spec)).then((buf) => { fs.writeFileSync(outPath, buf); console.log('wrote', outPath); });
}
module.exports = { build, clean, toThaiDigits, normalizeThai };
