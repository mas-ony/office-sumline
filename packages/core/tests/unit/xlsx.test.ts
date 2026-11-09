import { describe, expect, it } from 'vitest';
import ExcelJS from 'exceljs';
import { unzipSync, strFromU8 } from 'fflate';
import { readXlsx } from '../../src/core/xlsx/import';
import { writeXlsx } from '../../src/core/xlsx/export';
import { CalcEngine } from '../../src/core/formula/engine';
import { Workbook } from '../../src/core/model/workbook';
import { resolveColor } from '../../src/core/model/color';
import { formatValue } from '../../src/core/format/display';
import { formulaFromFile, formulaToFile } from '../../src/core/xlsx/formulaText';
import { readFixture } from './helpers';
import { needsArrayFormula } from '../../src/core/xlsx/arrayFormulas';
import { csvToWorkbook, decodeCsv } from '../../src/core/csv';
import { CellError } from '../../src/core/values';

const fixture = readFixture;

/** ExcelJS takes an ArrayBuffer (its Buffer type no longer matches Node's). */
function asArrayBuffer(bytes: Uint8Array): ExcelJS.Buffer {
  return bytes.buffer.slice(bytes.byteOffset, bytes.byteOffset + bytes.byteLength) as ExcelJS.Buffer;
}

function load(name: string) {
  const wb = readXlsx(fixture(name), { fileName: name });
  const engine = new CalcEngine(wb);
  engine.rebuild();
  engine.recalc();
  return { wb, engine };
}

function display(wb: Workbook, sheet: number, r: number, c: number): string {
  const s = wb.sheets[sheet];
  const cell = s.getCell(r, c);
  return formatValue(cell?.v ?? null, wb.styles.get(s.styleAt(r, c)).numFmt, wb.date1904).text;
}

describe('XLSX import', () => {
  it('reads sheets, values, formulas and number formats', () => {
    const { wb } = load('kitchen-sink.xlsx');
    expect(wb.sheets.map((s) => [s.name, s.state])).toEqual([['Sales', 'visible'], ['Data', 'visible'], ['Hidden', 'hidden']]);
    const sales = wb.sheets[0];
    expect(sales.getCell(0, 0)?.v).toBe('Quarterly Sales Report');
    expect(sales.getCell(2, 3)?.f).toBe('B3*C3');
    expect(display(wb, 0, 2, 2)).toBe('$9.50');
    expect(display(wb, 0, 3, 4)).toBe('-3.4%');
    expect(display(wb, 0, 2, 5)).toBe('2023-03-15');
    expect(display(wb, 0, 15, 1)).toBe('(1,234.50)');
    expect(display(wb, 0, 7, 3)).toBe('$6,501.25');
    // Recalculated from the Data sheet (the cached value in the file is 0).
    expect(sales.getCell(11, 1)?.v).toBe(4335);
    expect(sales.getCell(9, 1)?.v).toBe('East');
    expect(sales.getCell(17, 1)?.v).toBeInstanceOf(CellError);
    expect(sales.getCell(18, 1)?.v).toBe(6501.25);
    expect(sales.getCell(14, 1)?.v).toBe(true);
  });

  it('reads styles, layout and view settings', () => {
    const { wb } = load('kitchen-sink.xlsx');
    const sales = wb.sheets[0];
    const title = wb.styles.get(sales.styleAt(0, 0));
    expect(title.font).toMatchObject({ bold: true, size: 16, name: 'Calibri' });
    expect(resolveColor(title.font.color, wb.theme, '000000')).toBe('FFFFFF');
    expect(title.fill && 'pattern' in title.fill && title.fill.pattern).toBe('solid');
    expect(title.align).toMatchObject({ horizontal: 'center', vertical: 'center' });
    const header = wb.styles.get(sales.styleAt(1, 0));
    expect(header.border?.bottom?.style).toBe('medium');
    expect(header.align?.wrap).toBe(true);
    expect(sales.merges).toEqual([{ r0: 0, c0: 0, r1: 0, c1: 5 }]);
    expect(sales.freeze).toEqual({ rows: 2, cols: 0 });
    expect(sales.colWidthPx(0)).toBe(126);
    expect(sales.rowHeightPt(0)).toBe(30);
    expect(sales.hyperlinks[0]).toMatchObject({ target: 'https://example.com/', tooltip: 'Open the site' });
    expect(sales.validations[0]).toMatchObject({ type: 'list', formula1: '"North,South,East,West,Central"' });
    expect(sales.condFormats).toHaveLength(1);
    expect(wb.names).toEqual([{ name: 'Revenue', formula: 'Sales!$D$3:$D$7' }].map((n) => expect.objectContaining(n)));
  });

  it('reads images, other drawing objects and legacy notes', () => {
    const { wb } = load('kitchen-sink.xlsx');
    const sales = wb.sheets[0];
    expect(sales.images).toHaveLength(1);
    expect(sales.images[0]).toMatchObject({ name: 'Picture 1', mime: 'image/png', anchorType: 'twoCell', from: { r: 2, c: 7 } });
    // The chart becomes an editable chart (its part is kept while unchanged).
    expect(sales.otherAnchors).toHaveLength(0);
    expect(sales.charts).toHaveLength(1);
    expect(sales.charts[0]).toMatchObject({ type: 'column', title: 'Revenue by region', categories: 'Sales!$A$3:$A$7' });
    expect(sales.charts[0].series[0]).toMatchObject({ nameRef: 'Sales!$D$2', values: 'Sales!$D$3:$D$7' });
    expect(sales.charts[0].source?.chartPath).toBe('xl/charts/chart1.xml');
    expect(sales.comments).toHaveLength(1);
    expect(sales.comments[0]).toMatchObject({ r: 7, c: 3, kind: 'note', comments: [{ author: 'Reviewer', text: 'Check the West region figure.' }] });
  });

  it('reads threaded comments, dynamic arrays, shared formulas and rich text', () => {
    const { wb } = load('modern.xlsx');
    const s = wb.sheets[0];
    expect(s.getCell(3, 2)?.f).toBe('B4*2');
    expect(s.getCell(4, 2)?.v).toBe(2);
    const anchor = s.getCell(1, 4)!;
    expect(anchor.f).toBe('SORT(UNIQUE(A2:A5))');
    expect(anchor.dyn).toBe(true);
    expect([s.getCell(1, 4)?.v, s.getCell(2, 4)?.v, s.getCell(3, 4)?.v]).toEqual(['apple', 'fig', 'pear']);
    expect(s.getCell(1, 6)?.f).toBe('XLOOKUP("fig",A2:A5,B2:B5)');
    expect(s.getCell(1, 6)?.v).toBe(2);
    expect(s.getCell(2, 6)?.f).toBe('SUM(E2#)');
    expect(s.getCell(3, 6)?.rt?.[0]).toMatchObject({ text: 'Rich', font: { bold: true } });
    expect(s.comments).toHaveLength(2);
    expect(s.comments[0]).toMatchObject({ kind: 'threaded', resolved: false, comments: [{ author: 'Ana Lee', text: 'Is pear in stock?' }, { author: 'Bo Chen', text: 'Yes, three left.' }] });
    expect(s.comments[1]).toMatchObject({ resolved: true });
  });
});

describe('formula text conversion', () => {
  it('strips and restores future function prefixes', () => {
    expect(formulaFromFile('_xlfn._xlws.SORT(_xlfn.UNIQUE(A1:A3))')).toBe('SORT(UNIQUE(A1:A3))');
    expect(formulaToFile('SORT(UNIQUE(A1:A3))')).toBe('_xlfn._xlws.SORT(_xlfn.UNIQUE(A1:A3))');
    expect(formulaFromFile('SUM(_xlfn.ANCHORARRAY(E2))')).toBe('SUM(E2#)');
    expect(formulaToFile('SUM(E2#)')).toBe('SUM(_xlfn.ANCHORARRAY(E2))');
    expect(formulaFromFile('_xlfn.SINGLE(A1:A3)')).toBe('@A1:A3');
    expect(formulaToFile('@A1:A3+1')).toBe('_xlfn.SINGLE(A1:A3)+1');
    expect(formulaToFile('SUM(A1:A3)')).toBe('SUM(A1:A3)');
  });
});

describe('export fast paths agree with the full analysis', () => {
  it('formulaToFile leaves formulas without newer functions untouched', () => {
    expect(formulaToFile('IF(C2=0,0,E2/C2)')).toBe('IF(C2=0,0,E2/C2)');
    expect(formulaToFile('SUM(A1:A3)*2')).toBe('SUM(A1:A3)*2');
    expect(formulaToFile('XLOOKUP(A1,B:B,C:C)')).toBe('_xlfn.XLOOKUP(A1,B:B,C:C)');
    expect(formulaToFile('CONCAT("x(",A1)')).toBe('_xlfn.CONCAT("x(",A1)');
    expect(formulaToFile('SUM(A1#)')).toBe('SUM(_xlfn.ANCHORARRAY(A1))');
    expect(formulaToFile('sort(A1:A3)')).toBe('_xlfn._xlws.SORT(A1:A3)');
    expect(formulaToFile('IFERROR(A1,#N/A)')).toBe('IFERROR(A1,#N/A)');
  });

  it('needsArrayFormula', () => {
    expect(needsArrayFormula('C2*D2')).toBe(false);
    expect(needsArrayFormula('IF(C2=0,0,E2/C2)')).toBe(false);
    expect(needsArrayFormula('1E5*$B$2')).toBe(false);
    expect(needsArrayFormula('SEQUENCE(3)')).toBe(true);
    expect(needsArrayFormula('sequence(3)+A1')).toBe(true);
    expect(needsArrayFormula('SUM(A1:A3)')).toBe(false);
    expect(needsArrayFormula('A1:A3*2')).toBe(true);
    expect(needsArrayFormula('ABS(A1:A3)')).toBe(true);
    expect(needsArrayFormula('Sheet2!A1*2')).toBe(false);
  });
});

describe('number precision', () => {
  it('keeps full double precision in saved files', () => {
    const wb = Workbook.blank();
    wb.sheets[0].setCell(0, 0, { v: 1 / 3 });
    wb.sheets[0].setCell(0, 1, { v: 0.1 + 0.2 });
    wb.sheets[0].setCell(0, 2, { v: 1.5e-12 });
    wb.sheets[0].setCell(0, 3, { v: 123456789012345680000 });
    const again = readXlsx(writeXlsx(wb), { fileName: 'p.xlsx' });
    expect(again.sheets[0].getCell(0, 0)?.v).toBe(1 / 3);
    expect(again.sheets[0].getCell(0, 1)?.v).toBe(0.1 + 0.2);
    expect(again.sheets[0].getCell(0, 2)?.v).toBe(1.5e-12);
    expect(again.sheets[0].getCell(0, 3)?.v).toBe(123456789012345680000);
  });
});

describe('CSV import', () => {
  it('parses quoted fields, numbers and dates', () => {
    const wb = csvToWorkbook(decodeCsv(new TextEncoder().encode('\uFEFFa;b;c\n"x;y";1,5;2024-01-15\n')), 'data');
    const s = wb.sheets[0];
    expect(s.name).toBe('data');
    expect(s.getCell(1, 0)?.v).toBe('x;y');
    expect(s.getCell(1, 2)?.v).toBe(45306);
  });
});

describe('XLSX round trip', () => {
  it('keeps content, styles and unknown parts through a save', async () => {
    const { wb } = load('kitchen-sink.xlsx');
    const bytes = writeXlsx(wb);
    const files = unzipSync(bytes);
    // Unknown and preserved parts.
    expect(files['xl/charts/chart1.xml']).toBeDefined();
    expect(Object.keys(files).some((f) => f.startsWith('xl/media/'))).toBe(true);
    const sheetXml = strFromU8(files['xl/worksheets/sheet1.xml']);
    expect(sheetXml).toContain('<conditionalFormatting sqref="E3:E7">');
    expect(sheetXml).toContain('<pageSetup orientation="landscape"/>');
    expect(sheetXml).toContain('state="frozen"');
    expect(strFromU8(files['xl/workbook.xml'])).toContain('fullCalcOnLoad="1"');
    expect(files['xl/calcChain.xml']).toBeUndefined();

    const again = readXlsx(bytes);
    const engine = new CalcEngine(again);
    engine.rebuild();
    engine.recalc();
    expect(display(again, 0, 2, 2)).toBe('$9.50');
    expect(again.sheets[0].getCell(11, 1)?.v).toBe(4335);
    expect(again.sheets[0].merges).toEqual(wb.sheets[0].merges);
    expect(again.sheets[0].images).toHaveLength(1);
    expect(again.sheets[0].charts).toHaveLength(1);
    expect(again.sheets[0].charts[0].title).toBe('Revenue by region');
    expect(again.sheets[0].comments[0].comments[0].text).toBe('Check the West region figure.');
    expect(again.sheets.map((s) => s.state)).toEqual(['visible', 'visible', 'hidden']);

    // An independent reader accepts the file.
    const x = new ExcelJS.Workbook();
    await x.xlsx.load(asArrayBuffer(bytes));
    const ws = x.getWorksheet('Sales')!;
    expect(ws.getCell('A1').value).toBe('Quarterly Sales Report');
    expect(ws.getCell('C3').numFmt).toBe('"$"#,##0.00');
    expect((ws.getCell('D3').value as { formula: string }).formula).toBe('B3*C3');
    expect(ws.getCell('A1').font?.bold).toBe(true);
  });

  it('keeps threaded comments and dynamic arrays', async () => {
    const { wb } = load('modern.xlsx');
    const bytes = writeXlsx(wb);
    const files = unzipSync(bytes);
    const sheetXml = strFromU8(files['xl/worksheets/sheet1.xml']);
    expect(sheetXml).toContain('_xlfn._xlws.SORT(_xlfn.UNIQUE(A2:A5))');
    expect(sheetXml).toMatch(/<c r="E2" cm="1"[^>]*><f t="array" ref="E2:E4" aca="false">/);
    expect(sheetXml).toContain('SUM(_xlfn.ANCHORARRAY(E2))');
    const again = readXlsx(bytes);
    const s = again.sheets[0];
    expect(s.comments.map((t) => [t.kind, t.resolved, t.comments.length])).toEqual([['threaded', false, 2], ['threaded', true, 1]]);
    expect(s.comments[0].comments[1].author).toBe('Bo Chen');
    expect(s.getCell(1, 4)?.dyn).toBe(true);
    const x = new ExcelJS.Workbook();
    await x.xlsx.load(asArrayBuffer(bytes));
    expect(x.getWorksheet('Stock')!.getCell('A2').value).toBe('pear');
  });

  it('writes a new workbook from scratch', async () => {
    const wb = Workbook.blank();
    const sheet = wb.sheets[0];
    sheet.setCell(0, 0, { v: 'Hello' });
    sheet.setCell(1, 0, { v: 41 });
    sheet.setCell(2, 0, { v: null, f: 'A2+1', dyn: true });
    sheet.setCell(0, 1, { v: null, f: 'SEQUENCE(2)', dyn: true });
    const engine = new CalcEngine(wb);
    engine.rebuild();
    engine.recalc();
    const bytes = writeXlsx(wb);
    const x = new ExcelJS.Workbook();
    await x.xlsx.load(asArrayBuffer(bytes));
    const ws = x.worksheets[0];
    expect(ws.getCell('A1').value).toBe('Hello');
    expect((ws.getCell('A3').value as { result: number }).result).toBe(42);
    const again = readXlsx(bytes);
    expect(again.sheets[0].getCell(1, 1)?.v).toBe(2);
    expect(again.sheets[0].getCell(0, 1)?.dyn).toBe(true);
  });
});
