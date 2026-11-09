import { describe, expect, it } from 'vitest';
import { existsSync, readFileSync } from 'node:fs';
import { readXls } from '../../src/core/xls/read';
import { readXlsb } from '../../src/core/xlsb/read';
import { openDocument, sniff } from '../../src/api/open';
import { resolveColor } from '../../src/core/model/color';
import { CellError } from '../../src/core/values';
import { buildXls, buildXlsb } from './binaryWorkbooks';

describe('Excel 97-2003 (.xls) files', () => {
  it('reads values, shared strings across CONTINUE records and formulas', () => {
    const wb = readXls(buildXls());
    expect(wb.sheets.map((s) => s.name)).toEqual(['Sheet1', 'Sheet2']);
    const s = wb.sheets[0];
    const v = (r: number, c: number) => s.getCell(r, c)?.v;
    const f = (r: number, c: number) => s.getCell(r, c)?.f;
    expect([v(0, 0), v(0, 1), v(0, 2), v(0, 3), v(0, 4)]).toEqual(['Region', 12.5, 7, 0.25, true]);
    expect(v(0, 5)).toBeInstanceOf(CellError);
    expect(wb.sheets[1].getCell(0, 0)?.v).toBe('Other');
    expect([f(1, 0), f(1, 1), f(1, 2), f(1, 3), f(1, 4)]).toEqual(['B1*2', 'SUM(B1:D1)', '"a"&"b"', 'Sheet2!$A$1', 'IF(B1>1,"big","small")']);
    expect([v(1, 2), v(1, 3), v(1, 4)]).toEqual(['ab', 'Other', 'big']);
    // Shared formula, relative to each cell.
    expect([f(2, 0), f(3, 0)]).toEqual(['A2+1', 'A3+1']);
    expect(wb.names).toEqual([{ name: 'MyRange', formula: 'Sheet1!$A$1:$B$2' }]);
  });

  it('reads styles, sizes, merges and frozen panes', () => {
    const wb = readXls(buildXls());
    const s = wb.sheets[0];
    const b1 = wb.styles.get(s.getCell(0, 1)?.s);
    expect(b1.font.bold).toBe(true);
    expect(resolveColor(b1.font.color, wb.theme, '000000')).toBe('FF0000');
    expect(b1.fill && 'pattern' in b1.fill ? b1.fill.pattern : null).toBe('solid');
    expect(b1.border?.left?.style).toBe('thin');
    const d1 = wb.styles.get(s.getCell(0, 3)?.s);
    expect(d1.numFmt).toBe('0.0%');
    expect(d1.align?.wrap).toBe(true);
    expect(s.cols.get(1)?.w).toBe(20);
    expect(s.rows.get(0)?.h).toBe(30);
    expect(s.merges).toEqual([{ r0: 5, r1: 5, c0: 0, c1: 1 }]);
    expect(s.freeze).toEqual({ rows: 1, cols: 0 });
  });

  it('opens real files and reports unsupported ones', () => {
    expect(sniff(buildXls(), 'book.xls')).toBe('xls');
    const doc = openDocument(buildXls(), 'Budget.xls');
    expect(doc.fileName).toBe('Budget.xlsx');
    expect(doc.sheet.getCell(1, 1)?.v).toBe(19.75);
    expect(() => readXls(new Uint8Array(600))).toThrow(/97-2003/);
    // A template made by an office suite, when one is installed here.
    const template = '/opt/kingsoft/wps-office/templates/XLS Worksheet.xls';
    if (existsSync(template)) expect(readXls(new Uint8Array(readFileSync(template))).sheets.length).toBeGreaterThan(0);
  });
});

describe('Excel binary (.xlsb) files', () => {
  it('reads values, formulas, names and styles', () => {
    const bytes = buildXlsb();
    expect(sniff(bytes, 'data.xlsb')).toBe('xlsb');
    const wb = readXlsb(bytes);
    const s = wb.sheets[0];
    expect(s.name).toBe('Data');
    const v = (r: number, c: number) => s.getCell(r, c)?.v;
    expect([v(0, 0), v(0, 1), v(0, 2), v(0, 3)]).toEqual(['Name', 3.5, 42, true]);
    expect([s.getCell(1, 0)?.f, s.getCell(1, 1)?.f, s.getCell(1, 2)?.f]).toEqual(['"x"', 'B1*2', 'SUM(B1:C1)']);
    expect([v(1, 0), v(1, 1), v(1, 2)]).toEqual(['x', 7, 45.5]);
    expect(wb.names).toEqual([{ name: 'Total', formula: 'Data!$B$2' }]);
    const b1 = wb.styles.get(s.getCell(0, 1)?.s);
    expect(b1.font.bold).toBe(true);
    expect(b1.numFmt).toBe('0.000');
    expect(b1.align?.horizontal).toBe('center');
    expect(resolveColor(b1.font.color, wb.theme, '000000')).toBe('123456');
    expect(b1.fill && 'fg' in b1.fill ? resolveColor(b1.fill.fg, wb.theme, '') : null).toBe('FFFF00');
    expect(s.cols.get(1)).toMatchObject({ w: 15, hidden: true });
    expect(s.merges).toEqual([{ r0: 2, r1: 2, c0: 0, c1: 1 }]);
    expect(s.freeze).toEqual({ rows: 2, cols: 0 });
    expect(s.defaultRowHeightPt).toBe(15);
  });
});
