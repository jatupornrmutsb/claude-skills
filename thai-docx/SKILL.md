---
name: thai-docx
description: "Use whenever creating or editing a .docx Word document that contains Thai-language text, alongside the base docx skill. Prevents common Thai-document corruption: font fallback breaking sara/vowel-tone rendering, table columns/margins shifting, and paragraph/line-spacing drift. Trigger whenever the user asks for a Word document, report, memo, or template with Thai content, or mentions เอกสาร, รายงาน, .docx ภาษาไทย, or complains about จัดหน้าเพี้ยน / ฟอนต์เพี้ยน / ตารางเลื่อน."
---

# Thai DOCX Formatting Guardrails

Companion to the base `docx` skill. Read that skill first for the general docx-js
workflow (creation, editing, verification). This skill adds Thai-specific rules
that prevent the most common failure modes: font fallback, table drift, and
paragraph/spacing corruption.

## Why Thai .docx files break

Thai script combines base consonants with vowels (sara) and tone marks that
render *above* or *below* the consonant, not inline. Word/LibreOffice rendering
depends entirely on:
1. A font that has correct Thai glyph positioning tables
2. Text staying within single, correctly-tagged runs (not split across `\n` or
   fragmented `Paragraph`/`TextRun` boundaries)
3. Exact DXA-based sizing for tables and page setup (percentage-based sizing
   renders inconsistently across Word/LibreOffice/Google Docs, and Thai text
   reflows less predictably than Latin text when column widths are ambiguous)

## Mandatory rules when generating Thai .docx with docx-js

### 1. Always set a Thai-safe font explicitly
Never rely on the docx-js default font. Set it at the document defaults level
AND per-run if mixing languages:

```javascript
const doc = new Document({
  styles: {
    default: {
      document: {
        run: { font: "TH Sarabun New", size: 32 } // 32 half-points = 16pt, standard Thai doc size
      }
    }
  },
  // ...
});
```

Preferred fonts, in order: `TH Sarabun New` (Thai government standard, most
compatible), `TH Sarabun PSK`, `Angsana New`. Avoid leaving font unset — the
fallback font Word substitutes often lacks correct mark positioning, which is
what causes สระ/วรรณยุกต์ to appear shifted or overlapping.

### 2. Never use `\n` — always separate Paragraph elements
This rule from the base docx skill is even more critical for Thai: a `\n`
inside a single `TextRun` can cause the line-break to interact badly with
Thai mark stacking and produce visibly shifted spacing on the wrapped line.

```javascript
// WRONG
new Paragraph({ children: [new TextRun("บรรทัดที่ 1\nบรรทัดที่ 2")] })

// RIGHT
new Paragraph({ children: [new TextRun("บรรทัดที่ 1")] }),
new Paragraph({ children: [new TextRun("บรรทัดที่ 2")] }),
```

### 3. Page size: set explicitly, never rely on default
docx-js defaults to A4, which is usually correct for Thai documents, but set
it explicitly so it's never ambiguous:

```javascript
page: {
  size: { width: 11906, height: 16838 }, // A4 in DXA (1440 = 1 inch)
  margin: { top: 1440, bottom: 1440, left: 1440, right: 1440 } // 1" margins, adjust to match required template
}
```

If following a Thai government/university template (e.g. RMUTSB memo format),
margins are often non-default (commonly 2.5cm/1cm variants) — confirm the
required margin values from the template before generating, rather than
using generic 1-inch margins.

### 4. Tables: dual DXA widths, always
This is the single most common cause of "ตารางเลื่อน":
- Set `columnWidths` (array, DXA) on the `Table` itself
- Set `width: { size: <matching value>, type: WidthType.DXA }` on **every
  cell** in every row
- The `columnWidths` array values must sum to the table's total width
- Never use `WidthType.PERCENTAGE` — it is the most common source of drift
  between Word, LibreOffice, and Google Docs

```javascript
new Table({
  columnWidths: [3000, 6000], // sums to 9000 DXA
  rows: [
    new TableRow({
      children: [
        new TableCell({ width: { size: 3000, type: WidthType.DXA }, children: [...] }),
        new TableCell({ width: { size: 6000, type: WidthType.DXA }, children: [...] }),
      ]
    })
  ]
})
```

### 5. Line/paragraph spacing: set explicitly for Thai
Thai text with tone marks needs slightly more line spacing than pure Latin
text to avoid marks from adjacent lines visually colliding. Set explicit
`spacing` rather than relying on Word defaults:

```javascript
new Paragraph({
  children: [...],
  spacing: { line: 360, lineRule: LineRuleType.AUTO, after: 120 } // 360 = 1.5 line spacing (240=single)
})
```

### 6. Verify before delivering — non-negotiable for Thai docs
Follow the base docx skill's verification step every time, with extra
attention to Thai-specific defects:

```bash
python scripts/office/soffice.py --headless --convert-to pdf output.docx
pdftoppm -jpeg -r 100 output.pdf page
```

Then visually inspect (`view` the resulting images) specifically for:
- Sara/vowel marks sitting at the wrong height or overlapping the consonant
- Table column widths matching what was specified (not compressed/expanded)
- Consistent margins across all pages, especially page 2+ after a table or
  image
- Line spacing consistent across paragraphs (no visually "tight" or
  "loose" lines mixed together)

If any of these look wrong, the fix is almost always one of: font not set
explicitly (#1), a stray `\n` (#2), or percentage-based table widths (#4) —
check those three first before re-deriving from scratch.

## Quick checklist before sending a Thai .docx to the user

- [ ] Font set explicitly in document defaults (TH Sarabun New or equivalent)
- [ ] No `\n` anywhere in TextRun content
- [ ] Page size and margins set explicitly (not left to default)
- [ ] All tables have matching `columnWidths` + per-cell DXA `width`
- [ ] Paragraph `spacing` set explicitly where line spacing matters
- [ ] Rendered to PDF → JPEG and visually inspected before delivering
