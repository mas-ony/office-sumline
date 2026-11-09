/** CSV import (a new single-sheet workbook) and export (displayed values of one sheet). */
import { parseDelimited, sniffDelimiter } from './clipboard';
import { parseInput } from './format/input';
import { normalizeFormulaInput } from './format/formulaInput';
import { Workbook } from './model/workbook';
import type { Sheet } from './model/sheet';
import { MAX_COLS, MAX_ROWS } from './refs';
import type { SpreadsheetDocument } from './document';

/** Decodes CSV bytes: UTF-8 (with or without BOM), UTF-16 with BOM, else Windows-1252. */
export function decodeCsv(bytes: Uint8Array): string {
  if (bytes[0] === 0xff && bytes[1] === 0xfe) return new TextDecoder('utf-16le').decode(bytes.subarray(2));
  if (bytes[0] === 0xfe && bytes[1] === 0xff) return new TextDecoder('utf-16be').decode(bytes.subarray(2));
  try {
    return new TextDecoder('utf-8', { fatal: true }).decode(bytes).replace(/^﻿/, '');
  } catch {
    return new TextDecoder('windows-1252').decode(bytes);
  }
}

export function csvToWorkbook(text: string, sheetName = 'Sheet1'): Workbook {
  const wb = Workbook.blank(sheetName.replace(/[\\/?*[\]:]/g, '_').slice(0, 31) || 'Sheet1');
  wb.fromCsv = true;
  const sheet = wb.sheets[0];
  const body = text.replace(/\r\n?/g, '\n').replace(/\n$/, '');
  const rows = parseDelimited(body, sniffDelimiter(body));
  const formats = new Map<string, number>();
  rows.forEach((row, r) => {
    if (r >= MAX_ROWS) return;
    row.forEach((raw, c) => {
      if (c >= MAX_COLS || raw === '') return;
      const parsed = parseInput(raw);
      if (parsed.formula !== undefined) {
        sheet.setCell(r, c, { v: null, f: normalizeFormulaInput(parsed.formula), dyn: true });
        return;
      }
      let s: number | undefined;
      if (parsed.numFmt) {
        s = formats.get(parsed.numFmt);
        if (s === undefined) {
          s = wb.styles.patch(0, { numFmt: parsed.numFmt });
          formats.set(parsed.numFmt, s);
        }
      }
      sheet.setCell(r, c, s ? { v: parsed.value, s } : { v: parsed.value });
    });
  });
  return wb;
}

function quoteField(text: string, delimiter: string): string {
  if (text.includes('"') || text.includes(delimiter) || /[\r\n]/.test(text) || /^\s|\s$/.test(text)) return `"${text.replace(/"/g, '""')}"`;
  return text;
}

/** The used range of a sheet as CSV, with values as displayed (like Excel's "CSV UTF-8"). */
export function sheetToCsv(doc: SpreadsheetDocument, sheet: Sheet = doc.sheet, delimiter = ','): string {
  const used = sheet.usedRange(false);
  if (!used) return '';
  const lines: string[] = [];
  for (let r = 0; r <= used.r1; r++) {
    const fields: string[] = [];
    for (let c = 0; c <= used.c1; c++) {
      const m = sheet.mergeAt(r, c);
      const text = m && (m.r0 !== r || m.c0 !== c) ? '' : doc.displayText(r, c, sheet);
      fields.push(quoteField(text, delimiter));
    }
    lines.push(fields.join(delimiter));
  }
  return `${lines.join('\r\n')}\r\n`;
}
