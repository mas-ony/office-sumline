import { strFromU8, unzipSync } from 'fflate';
import { describe, expect, it } from 'vitest';
import { SpreadsheetDocument, adjustDecimalFormat } from '../../src/core/document';
import { makeRange, parseRange } from '../../src/core/refs';
import { readXlsx } from '../../src/core/xlsx/import';
import { CellError } from '../../src/core/values';
import { readFixture } from './helpers';

function doc() {
  return SpreadsheetDocument.blank();
}

function v(d: SpreadsheetDocument, ref: string) {
  const r = parseRange(ref)!;
  return d.sheet.getCell(r.r0, r.c0)?.v ?? null;
}

function put(d: SpreadsheetDocument, cells: Record<string, string>) {
  for (const [ref, text] of Object.entries(cells)) {
    const r = parseRange(ref)!;
    d.setInput(r.r0, r.c0, text);
  }
}

describe('typing into cells', () => {
  it('parses numbers, percentages, currency and dates with formats', () => {
    const d = doc();
    put(d, { A1: '1,234.5', A2: '15%', A3: '$1,000', A4: '2024-01-15', A5: "'00123", A6: 'TRUE', A7: '=sum(a1:a3)' });
    expect(v(d, 'A1')).toBe(1234.5);
    expect(v(d, 'A2')).toBe(0.15);
    expect(d.displayText(1, 0)).toBe('15%');
    expect(d.displayText(2, 0)).toBe('$1,000');
    expect(v(d, 'A4')).toBe(45306);
    expect(d.displayText(3, 0)).toBe('2024-01-15');
    expect(v(d, 'A5')).toBe('00123');
    expect(d.editText(4, 0)).toBe("'00123");
    expect(v(d, 'A6')).toBe(true);
    expect(d.sheet.getCell(6, 0)?.f).toBe('SUM(A1:A3)');
    expect(v(d, 'A7')).toBeCloseTo(2234.65);
  });

  it('closes parentheses and creates links for web addresses', () => {
    const d = doc();
    put(d, { A1: '=ROUND(2.5,0', B1: 'https://example.com' });
    expect(d.sheet.getCell(0, 0)?.f).toBe('ROUND(2.5,0)');
    expect(v(d, 'A1')).toBe(3);
    expect(d.sheet.hyperlinkAt(0, 1)?.target).toBe('https://example.com');
  });
});

describe('undo and redo', () => {
  it('restores values and recalculates dependents', () => {
    const d = doc();
    put(d, { A1: '1', A2: '=A1*10' });
    put(d, { A1: '5' });
    expect(v(d, 'A2')).toBe(50);
    d.undo();
    expect(v(d, 'A1')).toBe(1);
    expect(v(d, 'A2')).toBe(10);
    d.redo();
    expect(v(d, 'A2')).toBe(50);
  });

  it('undoes structural changes', () => {
    const d = doc();
    put(d, { A1: 'a', A2: 'b', B3: '=A2' });
    d.insertRows(1, 2);
    expect(v(d, 'A4')).toBe('b');
    expect(d.sheet.getCell(4, 1)?.f).toBe('A4');
    d.undo();
    expect(v(d, 'A2')).toBe('b');
    expect(d.sheet.getCell(2, 1)?.f).toBe('A2');
  });
});

describe('rows and columns', () => {
  it('shifts references in every sheet and makes deleted ones #REF!', () => {
    const d = doc();
    d.addSheet('Other');
    d.setActiveSheet(0);
    put(d, { A1: '10', A2: '20', A3: '=SUM(A1:A2)' });
    d.setActiveSheet(1);
    put(d, { A1: '=Sheet1!A2*2' });
    d.setActiveSheet(0);
    d.insertRows(0, 1);
    expect(d.sheet.getCell(3, 0)?.f).toBe('SUM(A2:A3)');
    expect(d.wb.sheets[1].getCell(0, 0)?.f).toBe('Sheet1!A3*2');
    d.deleteRows(2, 1);
    expect(d.sheet.getCell(2, 0)?.f).toBe('SUM(A2:A2)');
    expect(d.wb.sheets[1].getCell(0, 0)?.f).toBe('Sheet1!#REF!*2');
    expect(d.wb.sheets[1].getCell(0, 0)?.v).toBeInstanceOf(CellError);
  });

  it('inserts columns as wide as the selection and copies formatting', () => {
    const d = doc();
    put(d, { A1: 'x', B1: 'y' });
    d.select(makeRange(0, 0, 0, 0));
    d.toggleFont('bold');
    d.select(makeRange(0, 1, 0, 2));
    d.insertColsAtSelection('left');
    expect(v(d, 'D1')).toBe('y');
    expect(d.styleAt(0, 1).font.bold).toBe(true);
  });
});

describe('formatting', () => {
  it('applies borders around and inside a range', () => {
    const d = doc();
    d.select(parseRange('B2:C3')!);
    d.setBorder('outside', { style: 'medium' });
    expect(d.styleAt(1, 1).border).toMatchObject({ top: { style: 'medium' }, left: { style: 'medium' } });
    expect(d.styleAt(1, 1).border?.right).toBeUndefined();
    d.setBorder('inside', { style: 'thin' });
    expect(d.styleAt(1, 1).border?.right?.style).toBe('thin');
    d.setBorder('none');
    expect(d.styleAt(2, 2).border).toBeUndefined();
  });

  it('formats whole columns without creating cells', () => {
    const d = doc();
    put(d, { C5: 'v' });
    d.selectCols(2, 2);
    d.format({ font: { italic: true } });
    expect(d.sheet.cols.get(2)?.s).toBeGreaterThan(0);
    expect(d.styleAt(1000, 2).font.italic).toBe(true);
    expect(d.styleAt(4, 2).font.italic).toBe(true);
    expect(d.sheet.cells.size).toBe(1);
  });

  it('adjusts decimals', () => {
    expect(adjustDecimalFormat('0.00', 1)).toBe('0.000');
    expect(adjustDecimalFormat('$#,##0.00', -1)).toBe('$#,##0.0');
    expect(adjustDecimalFormat('0.0%', -1)).toBe('0%');
    expect(adjustDecimalFormat('#,##0.00;(#,##0.00)', 1)).toBe('#,##0.000;(#,##0.000)');
    expect(adjustDecimalFormat('yyyy-mm-dd', 1)).toBeNull();
    const d = doc();
    put(d, { A1: '3.14159' });
    d.select(makeRange(0, 0));
    d.adjustDecimals(-1);
    expect(d.displayText(0, 0)).toBe('3.1416');
  });

  it('merges keeping the top-left value', () => {
    const d = doc();
    put(d, { A1: 'title', B1: 'gone' });
    d.select(parseRange('A1:C2')!);
    d.toggleMerge();
    expect(d.sheet.merges).toEqual([parseRange('A1:C2')]);
    expect(v(d, 'B1')).toBe(null);
    d.toggleMerge();
    expect(d.sheet.merges).toEqual([]);
    d.undo();
    d.undo();
    expect(v(d, 'B1')).toBe('gone');
  });
});

describe('clipboard and fill', () => {
  it('copies formulas with relative references and tiles into larger targets', () => {
    const d = doc();
    put(d, { A1: '1', A2: '2', B1: '=A1*2' });
    d.select(makeRange(0, 1));
    const clip = d.copy()!;
    expect(clip.text).toBe('2');
    d.select(parseRange('B2:B3')!);
    d.paste(clip);
    expect(d.sheet.getCell(1, 1)?.f).toBe('A2*2');
    expect(d.sheet.getCell(2, 1)?.f).toBe('A3*2');
    expect(v(d, 'B2')).toBe(4);
  });

  it('moves cells on cut and paste and updates references to them', () => {
    const d = doc();
    put(d, { A1: '7', C1: '=A1+1' });
    d.select(makeRange(0, 0));
    d.cut();
    d.select(makeRange(5, 0));
    d.paste({});
    expect(v(d, 'A1')).toBe(null);
    expect(v(d, 'A6')).toBe(7);
    expect(d.sheet.getCell(0, 2)?.f).toBe('A6+1');
    expect(v(d, 'C1')).toBe(8);
  });

  it('pastes external tab separated text', () => {
    const d = doc();
    d.select(makeRange(0, 0));
    d.paste({ text: 'a\t1\nb\t2,000' });
    expect(v(d, 'A2')).toBe('b');
    expect(v(d, 'B2')).toBe(2000);
  });

  it('fills series, dates, names and formulas', () => {
    const d = doc();
    put(d, { A1: '1', A2: '3', B1: 'Item 1', C1: 'Mon', D1: '2024-01-31', E1: '=A1*2' });
    d.select(parseRange('A1:A2')!);
    d.fill(parseRange('A1:A5')!);
    expect([v(d, 'A3'), v(d, 'A4'), v(d, 'A5')]).toEqual([5, 7, 9]);
    d.select(parseRange('B1:E1')!);
    d.fill(parseRange('B1:E3')!);
    expect(v(d, 'B3')).toBe('Item 3');
    expect(v(d, 'C3')).toBe('Wed');
    expect(d.displayText(2, 3)).toBe('2024-02-02');
    expect(d.sheet.getCell(2, 4)?.f).toBe('A3*2');
  });
});

describe('sheets', () => {
  it('adds, renames, moves and removes sheets', () => {
    const d = doc();
    put(d, { A1: '=Data!A1' });
    const s = d.addSheet('Data')!;
    expect(d.sheet).toBe(s);
    put(d, { A1: '42' });
    d.setActiveSheet(0);
    expect(v(d, 'A1')).toBe(42);
    expect(d.renameSheet(1, 'My Data')).toBeNull();
    expect(d.sheet.getCell(0, 0)?.f).toBe("'My Data'!A1");
    expect(d.renameSheet(1, 'Sheet1')).toMatch(/already exists/);
    d.moveSheet(1, 0);
    expect(d.wb.sheets.map((x) => x.name)).toEqual(['My Data', 'Sheet1']);
    expect(d.removeSheet(0)).toBe(true);
    expect(d.sheet.getCell(0, 0)?.f).toBe('#REF!');
    expect(d.removeSheet(0)).toBe(false);
    d.undo();
    expect(d.wb.sheets).toHaveLength(2);
  });
});

describe('comments, search and sort', () => {
  it('manages comment threads', () => {
    const d = doc();
    d.author = 'Tester';
    const id = d.addComment(0, 0, 'First');
    const replyId = d.addComment(0, 0, 'Second');
    const thread = d.sheet.commentAt(0, 0)!;
    expect(thread.comments.map((c) => c.text)).toEqual(['First', 'Second']);
    expect(/^\{[0-9A-F-]{36}\}$/.test(id)).toBe(true);
    d.setResolved(id, true);
    expect(d.sheet.commentAt(0, 0)?.resolved).toBe(true);
    d.editComment(replyId, 'Edited');
    expect(d.sheet.commentAt(0, 0)?.resolved).toBe(false);
    d.deleteComment(replyId);
    expect(d.sheet.commentAt(0, 0)?.comments).toHaveLength(1);
    d.deleteComment(id);
    expect(d.sheet.comments).toHaveLength(0);
  });

  it('finds and replaces', () => {
    const d = doc();
    put(d, { A1: 'apple pie', A2: 'Apple', B1: '=A1' });
    expect(d.find('apple').map((h) => [h.r, h.c])).toEqual([[0, 0], [0, 1], [1, 0]]);
    expect(d.find('apple', { matchCase: true })).toHaveLength(2);
    expect(d.replaceAll('apple', 'pear')).toBe(2);
    expect(v(d, 'A1')).toBe('pear pie');
    expect(v(d, 'B1')).toBe('pear pie');
  });

  it('sorts rows by a column', () => {
    const d = doc();
    put(d, { A1: 'name', B1: 'n', A2: 'b', B2: '2', A3: 'a', B3: '1', A4: 'c', B4: '' });
    d.sortRange(parseRange('A1:B4')!, 1, true, true);
    expect([v(d, 'A2'), v(d, 'A3'), v(d, 'A4')]).toEqual(['a', 'b', 'c']);
  });
});

describe('saving', () => {
  it('round-trips edits through XLSX', () => {
    const d = doc();
    put(d, { A1: 'Total', B1: '=SUM(1,2)', A2: '12%' });
    d.select(makeRange(0, 0));
    d.toggleFont('bold');
    d.addComment(1, 0, 'note');
    d.setFreeze(1, 0);
    const wb = readXlsx(d.save());
    const s = wb.sheets[0];
    expect(s.getCell(0, 1)?.f).toBe('SUM(1,2)');
    expect(wb.styles.get(s.styleAt(0, 0)).font.bold).toBe(true);
    expect(wb.styles.get(s.styleAt(1, 0)).numFmt).toBe('0%');
    expect(s.comments[0].comments[0].text).toBe('note');
    expect(s.freeze.rows).toBe(1);
  });
});

describe('chart references', () => {
  const load = () => SpreadsheetDocument.fromXlsx(readFixture('kitchen-sink.xlsx'), 'k.xlsx');
  const refs = (d: SpreadsheetDocument) => [...d.wb.chartRefs.values()][0];

  it('follow inserted rows, renamed sheets and undo', () => {
    const d = load();
    expect(refs(d)).toEqual(['Sales!$D$2', 'Sales!$A$3:$A$7', 'Sales!$D$3:$D$7']);
    d.insertRows(1, 2);
    expect(refs(d)).toEqual(['Sales!$D$4', 'Sales!$A$5:$A$9', 'Sales!$D$5:$D$9']);
    d.renameSheet(0, 'Q1 Sales');
    expect(refs(d)).toEqual(["'Q1 Sales'!$D$4", "'Q1 Sales'!$A$5:$A$9", "'Q1 Sales'!$D$5:$D$9"]);
    d.undo();
    d.undo();
    expect(refs(d)).toEqual(['Sales!$D$2', 'Sales!$A$3:$A$7', 'Sales!$D$3:$D$7']);
    d.redo();
    expect(refs(d)).toEqual(['Sales!$D$4', 'Sales!$A$5:$A$9', 'Sales!$D$5:$D$9']);
  });

  it('are written into the saved chart part', () => {
    const d = load();
    d.deleteRows(3, 1);
    d.renameSheet(0, 'Totals');
    const files = unzipSync(d.save());
    const xml = strFromU8(files['xl/charts/chart1.xml']);
    expect(xml).toContain('<c:f>Totals!$A$3:$A$6</c:f>');
    expect(xml).toContain('<c:f>Totals!$D$3:$D$6</c:f>');
    expect(xml).not.toContain('Sales!');
  });

  it('point at #REF! when their sheet is deleted', () => {
    const d = load();
    d.addSheet('Other');
    d.removeSheet(0);
    expect(refs(d).every((f) => f.includes('#REF!'))).toBe(true);
  });
});
