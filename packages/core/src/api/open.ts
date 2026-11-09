/** Opening files: XLSX/XLSM (and templates), XLSB, XLS (Excel 97-2003) and CSV/TSV text. */
import { SpreadsheetDocument } from '../core/document';
import { csvToWorkbook, decodeCsv } from '../core/csv';
import { readXlsb } from '../core/xlsb/read';
import { readXls } from '../core/xls/read';
import { unzipSync } from 'fflate';

export const ACCEPT = '.xlsx,.xlsm,.xltx,.xltm,.xlsb,.xls,.csv,.tsv,.txt,application/vnd.openxmlformats-officedocument.spreadsheetml.sheet,application/vnd.ms-excel,text/csv';

const isCsvName = (name: string) => /\.(csv|tsv|txt)$/i.test(name);

/** Blank workbook styled like the online editor's new spreadsheet: 100px columns, 26px rows. */
export function newDocument(name = 'Untitled spreadsheet.xlsx'): SpreadsheetDocument {
  const doc = SpreadsheetDocument.blank(name);
  for (const sheet of doc.wb.sheets) {
    sheet.defaultColWidthPx = 100;
    sheet.defaultRowHeightPt = 19.5;
  }
  doc.modified = false;
  return doc;
}

export type FileKind = 'xlsx' | 'xlsb' | 'xls' | 'csv';

/** What a file is, from its first bytes (and its name for text files). */
export function sniff(bytes: Uint8Array, name: string): FileKind {
  if (bytes[0] === 0x50 && bytes[1] === 0x4b) {
    try {
      const files = unzipSync(bytes, { filter: (f) => f.name === 'xl/workbook.bin' || f.name === '[Content_Types].xml' });
      if (files['xl/workbook.bin']) return 'xlsb';
    } catch {
      // Not a readable zip: let the XLSX reader report it.
    }
    return 'xlsx';
  }
  if (bytes[0] === 0xd0 && bytes[1] === 0xcf && bytes[2] === 0x11 && bytes[3] === 0xe0) return 'xls';
  if (/\.xlsb$/i.test(name)) return 'xlsb';
  if (/\.xls$/i.test(name) && !isCsvName(name)) {
    // Some ".xls" files are HTML or text exports; treat what is not binary as text.
    return bytes.slice(0, 512).some((b) => b === 0) ? 'xls' : 'csv';
  }
  return 'csv';
}

/** Builds a document from file bytes. The name decides the name used when saving. */
export function openDocument(bytes: Uint8Array, name: string): SpreadsheetDocument {
  const kind = sniff(bytes, name);
  const base = name.replace(/\.[^.]+$/, '') || 'Workbook';
  switch (kind) {
    case 'xlsx':
      return SpreadsheetDocument.fromXlsx(bytes, name.replace(/\.xlt([xm])$/i, '.xls$1'));
    case 'xlsb':
      // Saved as XLSX (writing the binary format is not supported).
      return new SpreadsheetDocument(readXlsb(bytes), `${base}.xlsx`);
    case 'xls':
      return new SpreadsheetDocument(readXls(bytes), `${base}.xlsx`);
    case 'csv': {
      const wb = csvToWorkbook(decodeCsv(bytes), base);
      return new SpreadsheetDocument(wb, isCsvName(name) ? name : `${base}.csv`);
    }
  }
}
