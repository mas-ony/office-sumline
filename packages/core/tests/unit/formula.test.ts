import { describe, expect, it } from 'vitest';
import { errCode, evalFormula, makeBook } from './helpers';
import { parseFormula } from '../../src/core/formula/parser';
import { tokenize } from '../../src/core/formula/tokenizer';
import { moveRefsInFormula, renameSheetInFormula, shiftFormula, translateFormula } from '../../src/core/formula/adjust';
import { ArrayValue } from '../../src/core/formula/runtime';

const close = (a: unknown, b: number, digits = 9) => expect(a as number).toBeCloseTo(b, digits);

describe('tokenizer and parser', () => {
  it('tokenizes references with sheets, ranges and whole columns', () => {
    const t = tokenize("SUM('My Sheet'!$A$1:B2, Data!C:C, 3:5) + A1#");
    const refs = t.filter((x) => x.type === 'ref');
    expect(refs.map((r) => r.text)).toEqual(["'My Sheet'!$A$1:B2", 'Data!C:C', '3:5', 'A1']);
    expect(refs[0].ref).toMatchObject({ sheet: 'My Sheet', kind: 'range', r0: 0, c0: 0, r1: 1, c1: 1, absR0: true, absC0: true });
    expect(t.some((x) => x.type === 'spill')).toBe(true);
  });

  it('distinguishes functions, names and cell-like names', () => {
    const t = tokenize('LOG10(TaxRate)+_xlfn.XLOOKUP(1,A1:A2,B1:B2)');
    expect(t.filter((x) => x.type === 'func').map((x) => x.value)).toEqual(['LOG10', '_xlfn.XLOOKUP']);
    expect(t.find((x) => x.type === 'name')?.value).toBe('TaxRate');
  });

  it('follows Excel precedence', () => {
    expect(evalFormula('=-2^2')).toBe(4);
    expect(evalFormula('=2^3^2')).toBe(64);
    expect(evalFormula('=1+2*3')).toBe(7);
    expect(evalFormula('=50%*2')).toBe(1);
    expect(evalFormula('="a"&1+1')).toBe('a2');
    expect(evalFormula('=1<2=TRUE')).toBe(true);
    expect(parseFormula('A1:B2 B2:C3')).toMatchObject({ t: 'bin', op: ' ' });
  });
});

describe('operators and coercion', () => {
  it('coerces text and booleans in arithmetic', () => {
    expect(evalFormula('="3"+TRUE')).toBe(4);
    expect(evalFormula('="1,000"*2')).toBe(2000);
    expect(errCode(evalFormula('="abc"+1'))).toBe('#VALUE!');
    expect(errCode(evalFormula('=1/0'))).toBe('#DIV/0!');
  });
  it('compares like Excel', () => {
    expect(evalFormula('="a"="A"')).toBe(true);
    expect(evalFormula('=1<"a"')).toBe(true);
    expect(evalFormula('="z"<TRUE')).toBe(true);
    expect(evalFormula('=0.1+0.2=0.3')).toBe(true);
  });
  it('treats blank cells as zero or empty text', () => {
    expect(evalFormula('=A1+1')).toBe(1);
    expect(evalFormula('=A1&"x"')).toBe('x');
    expect(evalFormula('=A1')).toBe(0);
  });
});

describe('math and statistics', () => {
  const data = { A1: 1, A2: 2, A3: 3, A4: 'text', A5: true, A6: null, B1: 10, B2: 20, B3: 30 };
  it('aggregates following the reference rules', () => {
    expect(evalFormula('=SUM(A1:A6)', data)).toBe(6);
    expect(evalFormula('=SUM("3",TRUE,A1)', data)).toBe(5);
    expect(evalFormula('=AVERAGE(A1:A6)', data)).toBe(2);
    expect(evalFormula('=AVERAGEA(A1:A6)', data)).toBe(1.4);
    expect(evalFormula('=COUNT(A1:A6)', data)).toBe(3);
    expect(evalFormula('=COUNTA(A1:A6)', data)).toBe(5);
    expect(evalFormula('=COUNTBLANK(A1:A6)', data)).toBe(1);
    expect(evalFormula('=MAX(A1:A6)', data)).toBe(3);
    expect(evalFormula('=MIN(B1:B3)', data)).toBe(10);
    expect(evalFormula('=PRODUCT(A1:A3)', data)).toBe(6);
    expect(evalFormula('=SUMPRODUCT(A1:A3,B1:B3)', data)).toBe(140);
    close(evalFormula('=STDEV(B1:B3)', data), 10);
    close(evalFormula('=VAR.P(B1:B3)', data), 66.6666666667);
    expect(evalFormula('=MEDIAN(1,5,3,2)')).toBe(2.5);
    expect(evalFormula('=LARGE(B1:B3,2)', data)).toBe(20);
    expect(evalFormula('=RANK(20,B1:B3)', data)).toBe(2);
    close(evalFormula('=PERCENTILE(B1:B3,0.25)', data), 15);
  });
  it('rounds like Excel', () => {
    expect(evalFormula('=ROUND(2.675,2)')).toBe(2.68);
    expect(evalFormula('=ROUND(-2.5,0)')).toBe(-3);
    expect(evalFormula('=ROUNDUP(3.14159,3)')).toBe(3.142);
    expect(evalFormula('=ROUNDDOWN(-3.14159,1)')).toBe(-3.1);
    expect(evalFormula('=ROUND(1234.5678,-2)')).toBe(1200);
    expect(evalFormula('=MROUND(10,3)')).toBe(9);
    expect(evalFormula('=CEILING(2.5,1)')).toBe(3);
    expect(evalFormula('=CEILING(-2.5,-2)')).toBe(-4);
    expect(evalFormula('=FLOOR(-2.5,2)')).toBe(-4);
    expect(evalFormula('=MOD(-3,2)')).toBe(1);
    expect(evalFormula('=INT(-1.5)')).toBe(-2);
    expect(evalFormula('=TRUNC(-1.5)')).toBe(-1);
  });
  it('handles conditional aggregates and criteria', () => {
    const d = { A1: 'apple', A2: 'banana', A3: 'Apple pie', A4: 'cherry', B1: 1, B2: 2, B3: 3, B4: 4 };
    expect(evalFormula('=COUNTIF(A1:A4,"apple")', d)).toBe(1);
    expect(evalFormula('=COUNTIF(A1:A4,"a*")', d)).toBe(2);
    expect(evalFormula('=SUMIF(A1:A4,"<>banana",B1:B4)', d)).toBe(8);
    expect(evalFormula('=SUMIF(B1:B4,">2")', d)).toBe(7);
    expect(evalFormula('=SUMIFS(B1:B4,A1:A4,"*e*",B1:B4,">1")', d)).toBe(7);
    expect(evalFormula('=AVERAGEIF(B1:B4,">=2")', d)).toBe(3);
    expect(evalFormula('=COUNTIFS(A1:A4,"?????*",B1:B4,"<4")', d)).toBe(3);
    expect(evalFormula('=MAXIFS(B1:B4,A1:A4,"*a*")', d)).toBe(3);
  });
  it('supports SUBTOTAL ignoring nested subtotals', () => {
    const b = makeBook();
    b.set({ A1: 1, A2: 2, A3: '=SUBTOTAL(9,A1:A2)', A4: 4, A5: '=SUBTOTAL(9,A1:A4)' });
    expect(b.get('A3')).toBe(3);
    expect(b.get('A5')).toBe(7);
  });
});

describe('text functions', () => {
  it('slices and searches', () => {
    expect(evalFormula('=LEFT("Hello",2)')).toBe('He');
    expect(evalFormula('=RIGHT("Hello",3)')).toBe('llo');
    expect(evalFormula('=MID("Hello",2,3)')).toBe('ell');
    expect(evalFormula('=FIND("l","Hello")')).toBe(3);
    expect(evalFormula('=SEARCH("L?O","hello")')).toBe(3);
    expect(errCode(evalFormula('=FIND("z","Hello")'))).toBe('#VALUE!');
    expect(evalFormula('=SUBSTITUTE("a-b-c","-","+",2)')).toBe('a-b+c');
    expect(evalFormula('=TRIM("  a   b  ")')).toBe('a b');
    expect(evalFormula('=PROPER("hello wORLD")')).toBe('Hello World');
    expect(evalFormula('=REPT("ab",3)')).toBe('ababab');
    expect(evalFormula('=TEXTJOIN("-",TRUE,"a","","b")')).toBe('a-b');
    expect(evalFormula('=TEXTBEFORE("a.b.c",".",-1)')).toBe('a.b');
    expect(evalFormula('=TEXTAFTER("a.b.c",".")')).toBe('b.c');
  });
  it('formats with TEXT and parses VALUE', () => {
    expect(evalFormula('=TEXT(1234.567,"#,##0.00")')).toBe('1,234.57');
    expect(evalFormula('=TEXT(0.256,"0.0%")')).toBe('25.6%');
    expect(evalFormula('=TEXT(45306,"yyyy-mm-dd")')).toBe('2024-01-15');
    expect(evalFormula('=VALUE("$1,234.50")')).toBe(1234.5);
    expect(evalFormula('=DOLLAR(-1234.567)')).toBe('($1,234.57)');
    expect(evalFormula('=FIXED(1234.567,1)')).toBe('1,234.6');
  });
});

describe('logical and errors', () => {
  it('evaluates IF lazily and IFERROR', () => {
    expect(evalFormula('=IF(TRUE,1,1/0)')).toBe(1);
    expect(evalFormula('=IF(FALSE,1)')).toBe(false);
    expect(evalFormula('=IFERROR(1/0,"div")')).toBe('div');
    expect(evalFormula('=IFNA(NA(),"na")')).toBe('na');
    expect(evalFormula('=IFS(1>2,"a",2>1,"b")')).toBe('b');
    expect(evalFormula('=SWITCH(2,1,"one",2,"two","other")')).toBe('two');
    expect(evalFormula('=AND(TRUE,1,"TRUE")')).toBe(true);
    expect(evalFormula('=XOR(TRUE,TRUE,TRUE)')).toBe(true);
    expect(evalFormula('=CHOOSE(2,"a","b","c")')).toBe('b');
    expect(evalFormula('=LET(x,2,y,3,x*y+x)')).toBe(8);
  });
  it('reports information', () => {
    expect(evalFormula('=ISBLANK(A1)')).toBe(true);
    expect(evalFormula('=ISNUMBER(1)')).toBe(true);
    expect(evalFormula('=TYPE("a")')).toBe(2);
    expect(evalFormula('=ERROR.TYPE(1/0)')).toBe(2);
    expect(evalFormula('=N(TRUE)')).toBe(1);
    expect(errCode(evalFormula('=NOSUCHFUNCTION(1)'))).toBe('#NAME?');
  });
});

describe('lookup and reference', () => {
  const t = { A1: 'id', B1: 'name', A2: 1, B2: 'one', A3: 2, B3: 'two', A4: 3, B4: 'three' };
  it('looks up values', () => {
    expect(evalFormula('=VLOOKUP(2,A2:B4,2,FALSE)', t)).toBe('two');
    expect(evalFormula('=VLOOKUP(2.5,A2:B4,2)', t)).toBe('two');
    expect(errCode(evalFormula('=VLOOKUP(9,A2:B4,2,FALSE)', t))).toBe('#N/A');
    expect(evalFormula('=HLOOKUP("name",A1:B4,3,FALSE)', t)).toBe('two');
    expect(evalFormula('=MATCH(3,A2:A4,0)', t)).toBe(3);
    expect(evalFormula('=MATCH("t*",B2:B4,0)', t)).toBe(2);
    expect(evalFormula('=INDEX(A2:B4,3,2)', t)).toBe('three');
    expect(evalFormula('=INDEX(B2:B4,MATCH(2,A2:A4,0))', t)).toBe('two');
    expect(evalFormula('=XLOOKUP(3,A2:A4,B2:B4)', t)).toBe('three');
    expect(evalFormula('=XLOOKUP(9,A2:A4,B2:B4,"none")', t)).toBe('none');
    expect(evalFormula('=XLOOKUP(2.5,A2:A4,B2:B4,,1)', t)).toBe('three');
    expect(evalFormula('=LOOKUP(2,A2:A4,B2:B4)', t)).toBe('two');
  });
  it('builds references', () => {
    expect(evalFormula('=SUM(OFFSET(A2,0,0,3,1))', t)).toBe(6);
    expect(evalFormula('=SUM(INDIRECT("A2:A4"))', t)).toBe(6);
    expect(evalFormula('=ROWS(A1:B4)*COLUMNS(A1:B4)', t)).toBe(8);
    expect(evalFormula('=ADDRESS(2,3)')).toBe('$C$2');
    expect(evalFormula('=ADDRESS(2,3,4,TRUE,"My Sheet")')).toBe("'My Sheet'!C2");
    expect(evalFormula('=SUM(A2:INDEX(A2:A4,2))', t)).toBe(3);
  });
});

describe('dates', () => {
  it('converts serials with the 1900 leap year bug', () => {
    expect(evalFormula('=DATE(2024,1,15)')).toBe(45306);
    expect(evalFormula('=DATE(1900,3,1)')).toBe(61);
    expect(evalFormula('=DATE(2024,14,1)')).toBe(evalFormula('=DATE(2025,2,1)'));
    expect(evalFormula('=YEAR(45306)')).toBe(2024);
    expect(evalFormula('=MONTH(45306)')).toBe(1);
    expect(evalFormula('=DAY(60)')).toBe(29);
    expect(evalFormula('=WEEKDAY(45306)')).toBe(2);
    expect(evalFormula('=WEEKDAY(45306,2)')).toBe(1);
    expect(evalFormula('=EOMONTH(45306,1)')).toBe(evalFormula('=DATE(2024,2,29)'));
    expect(evalFormula('=EDATE(DATE(2024,1,31),1)')).toBe(evalFormula('=DATE(2024,2,29)'));
    expect(evalFormula('=DATEDIF(DATE(2020,5,15),DATE(2024,1,10),"Y")')).toBe(3);
    expect(evalFormula('=DATEDIF(DATE(2020,5,15),DATE(2024,1,10),"YM")')).toBe(7);
    expect(evalFormula('=NETWORKDAYS(DATE(2024,1,1),DATE(2024,1,31))')).toBe(23);
    expect(evalFormula('=WORKDAY(DATE(2024,1,5),1)')).toBe(evalFormula('=DATE(2024,1,8)'));
    expect(evalFormula('=ISOWEEKNUM(DATE(2021,1,1))')).toBe(53);
    expect(evalFormula('=TIME(12,30,0)')).toBe(0.5208333333333334);
    expect(evalFormula('=HOUR(0.75)')).toBe(18);
    expect(evalFormula('=DATEVALUE("2024-01-15")')).toBe(45306);
    close(evalFormula('=YEARFRAC(DATE(2024,1,1),DATE(2024,7,1),1)'), 0.4972677596, 8);
    expect(evalFormula('=DAYS360(DATE(2024,1,31),DATE(2024,3,31))')).toBe(60);
  });
});

describe('financial', () => {
  it('computes loans and investments', () => {
    close(evalFormula('=PMT(0.05/12,360,200000)'), -1073.643246, 5);
    close(evalFormula('=FV(0.06/12,10,-200,-500,1)'), 2581.4033740, 6);
    close(evalFormula('=NPV(0.1,-10000,3000,4200,6800)'), 1188.443412, 5);
    close(evalFormula('=IRR({-70000,12000,15000,18000,21000,26000})'), 0.08663094803, 8);
    close(evalFormula('=PRICE(DATE(2008,2,15),DATE(2017,11,15),0.0575,0.065,100,2,0)'), 94.63436162, 6);
    close(evalFormula('=YIELD(DATE(2008,2,15),DATE(2016,11,15),0.0575,95.04287,100,2,0)'), 0.065, 6);
    expect(evalFormula('=COUPNUM(DATE(2011,1,25),DATE(2011,11,15),2,1)')).toBe(2);
  });
});

describe('dynamic arrays', () => {
  it('spills array results and reports #SPILL!', () => {
    const b = makeBook();
    b.set({ A1: '=SEQUENCE(3)' });
    expect([b.get('A1'), b.get('A2'), b.get('A3')]).toEqual([1, 2, 3]);
    b.set({ A3: 'x' });
    expect(errCode(b.get('A1'))).toBe('#SPILL!');
    expect(b.get('A2')).toBe(null);
    b.set({ A3: null });
    expect(b.get('A3')).toBe(3);
  });
  it('sorts, filters and de-duplicates', () => {
    const b = makeBook();
    b.set({ A1: 3, A2: 1, A3: 2, A4: 1, C1: '=SORT(A1:A4)', D1: '=UNIQUE(A1:A4)', E1: '=FILTER(A1:A4,A1:A4>1)', F1: '=SUM(C1#)' });
    expect([b.get('C1'), b.get('C2'), b.get('C3'), b.get('C4')]).toEqual([1, 1, 2, 3]);
    expect([b.get('D1'), b.get('D2'), b.get('D3'), b.get('D4')]).toEqual([3, 1, 2, null]);
    expect([b.get('E1'), b.get('E2')]).toEqual([3, 2]);
    expect(b.get('F1')).toBe(7);
  });
  it('maps scalar functions over arrays', () => {
    const b = makeBook();
    b.set({ A1: -1, A2: -2, B1: '=ABS(A1:A2)', C1: '=SUM(A1:A2*2)' });
    expect([b.get('B1'), b.get('B2')]).toEqual([1, 2]);
    expect(b.get('C1')).toBe(-6);
  });
});

describe('recalculation', () => {
  it('updates dependents incrementally', () => {
    const b = makeBook(['Sheet1', 'Data']);
    b.set({ 'Data!A1': 5, A1: '=Data!A1*2', A2: '=A1+1', A3: '=SUM(A1:A2)' });
    expect(b.get('A3')).toBe(21);
    b.set({ 'Data!A1': 10 });
    expect(b.get('A1')).toBe(20);
    expect(b.get('A3')).toBe(41);
  });
  it('handles long dependency chains without stack overflow', () => {
    const b = makeBook();
    const cells: Record<string, string | number> = { A1: 1 };
    for (let i = 2; i <= 6000; i++) cells[`A${i}`] = `=A${i - 1}+1`;
    b.set(cells);
    expect(b.get('A6000')).toBe(6000);
    b.set({ A1: 2 });
    expect(b.get('A6000')).toBe(6001);
  });
  it('resolves circular references to zero', () => {
    const b = makeBook();
    b.set({ A1: '=B1+1', B1: '=A1+1' });
    expect(typeof b.get('A1')).toBe('number');
    expect(b.engine.circular.size).toBeGreaterThan(0);
  });
  it('evaluates defined names', () => {
    const b = makeBook();
    b.wb.names.push({ name: 'Rate', formula: 'Sheet1!$B$1' });
    b.set({ B1: 0.5, A1: '=100*Rate' });
    expect(b.get('A1')).toBe(50);
    b.set({ B1: 0.25 });
    expect(b.get('A1')).toBe(25);
  });
});

describe('formula rewriting', () => {
  it('shifts references when rows are inserted or deleted', () => {
    expect(shiftFormula('SUM(A1:A10)+B5', 'Sheet1', 'Sheet1', 'row', 4, 2)).toBe('SUM(A1:A12)+B7');
    expect(shiftFormula('SUM(A1:A10)+B5', 'Sheet1', 'Sheet1', 'row', 5, -1)).toBe('SUM(A1:A9)+#REF!');
    expect(shiftFormula('Data!C3*2', 'Sheet1', 'Data', 'col', 1, 1)).toBe('Data!D3*2');
    expect(shiftFormula('C3*2', 'Sheet1', 'Data', 'col', 1, 1)).toBe('C3*2');
    expect(shiftFormula('SUM(A:A)', 'Sheet1', 'Sheet1', 'row', 0, 5)).toBe('SUM(A:A)');
  });
  it('translates relative references for copy and fill', () => {
    expect(translateFormula('A1+$B$1+C$2+$D3', 2, 1)).toBe('B3+$B$1+D$2+$D5');
    expect(translateFormula('A1', -1, 0)).toBe('#REF!');
  });
  it('renames sheets and follows moved cells', () => {
    expect(renameSheetInFormula("SUM(Data!A1:A3)+'Data'!B1", 'Data', 'My Data')).toBe("SUM('My Data'!A1:A3)+'My Data'!B1");
    expect(moveRefsInFormula('A1+A5', 'Sheet1', 'Sheet1', { r0: 0, c0: 0, r1: 1, c1: 0 }, 'Sheet1', 0, 3)).toBe('D1+A5');
  });
  it('produces arrays from array constants', () => {
    const v = makeBook();
    v.set({ A1: '={1,2;3,4}' });
    expect(v.get('B2')).toBe(4);
    expect(new ArrayValue([[1]]).rows).toBe(1);
  });
});
