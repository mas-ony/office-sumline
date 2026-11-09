import { describe, expect, it } from 'vitest';
import { SpreadsheetDocument } from '../../src/core/document';
import type { DataValidation } from '../../src/core/model/sheet';
import { parseRange } from '../../src/core/refs';
import { chartSettingsOf, chartSettingsPatch, insertChart, parseNum, relayoutChart, withSeriesStyle } from '../../src/ui/chartDialogs';
import { applyCellEditor, maskCode, maskFromCode, type MaskState } from '../../src/ui/formatMask';
import { hexToRgb, normalizeHex, PALETTE_ROWS, rgbToHex } from '../../src/ui/palette';
import { borderStyleFor, fitToolbar, isDivider, TOOLBAR_ITEMS } from '../../src/ui/toolbar';
import { blankValidationForm, ruleFromForm, TEXT_CRITERIA, type ValidationForm, validationFormOf } from '../../src/ui/validationForm';

function salesDoc(): SpreadsheetDocument {
  const d = SpreadsheetDocument.blank();
  d.setValues(0, 0, [
    ['Region', 'Q1', 'Q2', 'Q3'],
    ['North', 120, 150, 90],
    ['South', 90, 110, 130],
    ['East', 60, 80, 100],
  ]);
  return d;
}

describe('toolbar', () => {
  const total = TOOLBAR_ITEMS.reduce((s, i) => s + i.width, 0);
  const tools = TOOLBAR_ITEMS.map((i) => i.key).filter((k) => !isDivider(k));

  it('shows every item when they fit', () => {
    expect(fitToolbar(total)).toEqual({ shown: TOOLBAR_ITEMS.map((i) => i.key), overflow: [] });
  });

  it('moves the items that do not fit, next to the "more" button, to the overflow', () => {
    const { shown, overflow } = fitToolbar(total - 1);
    expect(overflow).toEqual(['rule', 'shape']);
    expect(shown.at(-1)).toBe('chart');
    expect([...shown.filter((k) => !isDivider(k)), ...overflow]).toEqual(tools);
  });

  it('never ends the shown items with a divider, and leaves dividers out of the overflow', () => {
    // Room up to the second divider (40 px for the "more" button).
    const upToDivider = 40 + TOOLBAR_ITEMS.slice(0, TOOLBAR_ITEMS.findIndex((i) => i.key === 'd2') + 1).reduce((s, i) => s + i.width, 0);
    const { shown, overflow } = fitToolbar(upToDivider);
    expect(shown.at(-1)).toBe('border');
    expect(overflow).toEqual(['image', 'search', 'chart', 'rule', 'shape']);
    expect(fitToolbar(10)).toEqual({ shown: [], overflow: tools });
  });

  it('maps the thickness and line kind to an XLSX border style', () => {
    expect([1, 2, 3, 5].map((w) => borderStyleFor(w, 'solid'))).toEqual(['thin', 'medium', 'thick', 'thick']);
    expect([1, 2].map((w) => borderStyleFor(w, 'dotted'))).toEqual(['dotted', 'mediumDashDotDot']);
    expect([1, 2].map((w) => borderStyleFor(w, 'dashed'))).toEqual(['dashed', 'mediumDashed']);
    expect(borderStyleFor(4, 'double')).toBe('double');
  });
});

describe('colour palette', () => {
  it('has the shades as rows and the hues as columns', () => {
    expect(PALETTE_ROWS).toHaveLength(10);
    expect(PALETTE_ROWS.every((row) => row.length === 14)).toBe(true);
    expect(PALETTE_ROWS[0][0]).toBe('#FFEBEE');
    expect(PALETTE_ROWS[9][13]).toBe('#000000');
  });

  it('converts between hex and RGB', () => {
    expect(hexToRgb('#F44336')).toEqual([244, 67, 54]);
    expect(hexToRgb('abc')).toEqual([170, 187, 204]);
    expect(rgbToHex(244, 67, 54)).toBe('#F44336');
    expect(rgbToHex(300, -5, 12.6)).toBe('#FF000D');
  });

  it('normalises typed colours', () => {
    expect(normalizeHex(' abc ')).toBe('#AABBCC');
    expect(normalizeHex('#c62828')).toBe('#C62828');
    expect(normalizeHex('#12345')).toBeNull();
    expect(normalizeHex('red')).toBeNull();
  });
});

describe('format cells: mask', () => {
  const mask = (s: Partial<MaskState>): MaskState => ({ category: 'number', decimals: '2', thousands: true, symbol: '$', negative: 0, custom: '', ...s });

  it('makes number format codes', () => {
    expect([0, 1, 2, 3].map((negative) => maskCode(mask({ negative })))).toEqual(['#,##0.00', '#,##0.00;[Red]#,##0.00', '#,##0.00;\\(#,##0.00\\)', '#,##0.00;[Red]\\(#,##0.00\\)']);
    expect(maskCode(mask({ thousands: false, decimals: '0' }))).toBe('0');
    expect(maskCode(mask({ category: 'currency', symbol: 'R"p', decimals: 'x' }))).toBe('"Rp"#,##0');
    expect(maskCode(mask({ category: 'accounting' }))).toBe('_("$"* #,##0.00_);_("$"* \\(#,##0.00\\);_("$"* "-"??_);_(@_)');
    expect(maskCode(mask({ category: 'percentage', decimals: '1' }))).toBe('0.0%');
    expect(maskCode(mask({ category: 'general' }))).toBe('General');
    expect(maskCode(mask({ category: 'text' }))).toBe('@');
    expect(maskCode(mask({ category: 'custom', custom: '' }))).toBe('General');
    expect(maskCode(mask({ decimals: '99' }))).toBe(`#,##0.${'0'.repeat(30)}`);
  });

  it('reads the settings back from the codes it makes', () => {
    const states: Partial<MaskState>[] = [
      { category: 'number', negative: 0 },
      { category: 'number', negative: 1, thousands: false, decimals: '3' },
      { category: 'number', negative: 2 },
      { category: 'number', negative: 3, decimals: '0' },
      { category: 'currency', symbol: '€', negative: 1 },
      { category: 'accounting', symbol: '£', decimals: '0' },
      { category: 'percentage', decimals: '2' },
      { category: 'general' },
      { category: 'text' },
    ];
    for (const s of states) {
      const code = maskCode(mask(s));
      const back = maskFromCode(code);
      expect(back.category, code).toBe(s.category);
      expect(maskCode(back), code).toBe(code);
    }
    expect(maskFromCode('yyyy-mm-dd')).toMatchObject({ category: 'custom', custom: 'yyyy-mm-dd' });
  });
});

describe('format cells: editor', () => {
  it('sets the format and the validation of the cell type', () => {
    const d = SpreadsheetDocument.blank();
    d.setValues(0, 0, [[0.25, 45306]]);
    d.select(parseRange('A1')!);
    applyCellEditor(d, 'percent', '', true);
    expect(d.displayText(0, 0)).toBe('25%');
    expect(d.styleAt(0, 0).align?.wrap).toBe(true);
    d.select(parseRange('B1')!);
    applyCellEditor(d, 'calendar', '', false);
    expect(d.displayText(0, 1)).toBe('15/01/2024');
    expect(d.sheet.validations.at(-1)).toMatchObject({ type: 'date', operator: 'greaterThan', errorStyle: 'stop' });
    d.select(parseRange('C1:C3')!);
    applyCellEditor(d, 'dropdown', 'red, "green"\nblue,,', false);
    expect(d.sheet.validations.at(-1)).toMatchObject({ type: 'list', formula1: '"red,green,blue"', sqref: [parseRange('C1:C3')] });
    applyCellEditor(d, 'checkbox', '', false);
    expect(d.sheet.validations.at(-1)?.formula1).toBe('"TRUE,FALSE"');
  });
});

describe('data validation form', () => {
  const sheet = 'Sales 2024';
  const form = (f: Partial<ValidationForm>): ValidationForm => ({ ...blankValidationForm(sheet, parseRange('B2:B9')!), ...f });
  const rule = (f: Partial<ValidationForm>): DataValidation => {
    const res = ruleFromForm(form(f), sheet, false);
    if ('error' in res) throw new Error(res.error);
    return res.rule;
  };
  const roundTrip = (f: Partial<ValidationForm>) => validationFormOf(rule(f), sheet);

  it('starts with the selection of the sheet', () => {
    expect(blankValidationForm(sheet, parseRange('A1:B3')!)).toMatchObject({ range: "'Sales 2024'!A1:B3", action: 'warning', kind: 'number', criteria: 'between', allowBlank: true });
  });

  it('makes number rules with a message, and reads them back', () => {
    const f = { action: 'reject', message: 'From 1 to 10', criteria: 'between', value1: '1', value2: ' 10 ' } as const;
    expect(rule(f)).toMatchObject({ sqref: [parseRange('B2:B9')], type: 'decimal', operator: 'between', formula1: '1', formula2: '10', errorStyle: 'stop', error: 'From 1 to 10', errorTitle: 'Invalid entry', showErrorMessage: true });
    expect(roundTrip(f)).toMatchObject({ range: "'Sales 2024'!B2:B9", action: 'reject', message: 'From 1 to 10', kind: 'number', criteria: 'between', value1: '1', value2: '10' });
    const formula = rule({ criteria: 'greaterThan', value1: '=$A$1' });
    expect(formula).toMatchObject({ operator: 'greaterThan', formula1: '$A$1' });
    expect(formula.formula2).toBeUndefined();
  });

  it('makes date, time and text length rules', () => {
    expect(rule({ kind: 'date', criteria: 'greaterThan', value1: '2024-01-15' }).formula1).toBe('45306');
    expect(roundTrip({ kind: 'date', criteria: 'between', value1: '2024-01-15', value2: '2024-02-01' })).toMatchObject({ kind: 'date', value1: '2024-01-15', value2: '2024-02-01' });
    expect(Number(rule({ kind: 'time', criteria: 'lessThan', value1: '12:30' }).formula1)).toBeCloseTo(0.5208333, 6);
    expect(roundTrip({ kind: 'time', criteria: 'lessThan', value1: '12:30' })).toMatchObject({ kind: 'time', value1: '12:30:00' });
    // Workbooks with the 1904 date system.
    const res = ruleFromForm(form({ kind: 'date', criteria: 'equal', value1: '2024-01-15' }), sheet, true);
    if ('error' in res) throw new Error(res.error);
    expect(res.rule.formula1).toBe(String(45306 - 1462));
    expect(validationFormOf(res.rule, sheet, true).value1).toBe('2024-01-15');
    expect(rule({ kind: 'textLength', criteria: 'lessThanOrEqual', value1: '5' })).toMatchObject({ type: 'textLength', operator: 'lessThanOrEqual', formula1: '5' });
  });

  it('makes text rules as formulas of the first cell, and recognises them again', () => {
    expect(rule({ kind: 'text', textCriteria: 'contains', value1: 'a"b' }).formula1).toBe('ISNUMBER(SEARCH("a""b",B2))');
    for (const { key } of TEXT_CRITERIA) {
      const value1 = key === 'email' || key === 'url' ? '' : 'x"y';
      expect(roundTrip({ kind: 'text', textCriteria: key, value1 }), key).toMatchObject({ kind: 'text', textCriteria: key, value1 });
    }
  });

  it('makes list, empty, not empty and formula rules', () => {
    expect(rule({ kind: 'list', value1: ' a, b,,"c" ', dropdown: false })).toMatchObject({ type: 'list', formula1: '"a,b,c"', hideDropDown: true });
    expect(roundTrip({ kind: 'list', value1: 'a, b' })).toMatchObject({ kind: 'list', value1: 'a,b', dropdown: true });
    expect(roundTrip({ kind: 'list', value1: '=$D$1:$D$5' })).toMatchObject({ kind: 'list', value1: '=$D$1:$D$5' });
    expect(rule({ kind: 'empty', allowBlank: false })).toMatchObject({ type: 'custom', formula1: 'LEN(B2)=0', allowBlank: true });
    expect(rule({ kind: 'notEmpty' })).toMatchObject({ type: 'custom', formula1: 'LEN(B2)>0', allowBlank: false });
    expect(roundTrip({ kind: 'empty' }).kind).toBe('empty');
    expect(roundTrip({ kind: 'notEmpty' }).kind).toBe('notEmpty');
    expect(roundTrip({ kind: 'formula', value1: '=B2>0' })).toMatchObject({ kind: 'formula', value1: '=B2>0' });
  });

  it('makes format style rules', () => {
    const style = { color: '#1565C0', bold: true };
    expect(rule({ action: 'format', style, value1: '0', value2: '1' })).toMatchObject({ formatStyle: style, showErrorMessage: false, errorStyle: 'information' });
    expect(roundTrip({ action: 'format', style, value1: '0', value2: '1' })).toMatchObject({ action: 'format', style });
  });

  it('takes several ranges, with or without the sheet name', () => {
    expect(rule({ range: "'Sales 2024'!a1:a3 C1, 'sales 2024'!E5", value1: '1', value2: '2' }).sqref).toEqual([parseRange('A1:A3'), parseRange('C1'), parseRange('E5')]);
    const blank = blankValidationForm("Bob's", parseRange('A1')!);
    expect(blank.range).toBe("'Bob''s'!A1");
    const quoted = ruleFromForm({ ...blank, value1: '1', value2: '2' }, "Bob's", false);
    expect('rule' in quoted && quoted.rule.sqref).toEqual([parseRange('A1')]);
  });

  it('says what is wrong with the form', () => {
    const error = (f: Partial<ValidationForm>) => {
      const res = ruleFromForm(form(f), sheet, false);
      return 'error' in res ? res.error : null;
    };
    expect(error({ range: 'Other!A1', value1: '1', value2: '2' })).toMatch(/cell range of this worksheet/);
    expect(error({ range: 'A1:', value1: '1', value2: '2' })).toMatch(/cell range/);
    expect(error({ value1: 'abc', value2: '2' })).toBe('Enter numbers or formulas starting with =.');
    expect(error({ value1: '1' })).toBe('Enter numbers or formulas starting with =.');
    expect(error({ kind: 'date', criteria: 'equal', value1: 'soon' })).toBe('Enter valid dates.');
    expect(error({ kind: 'time', criteria: 'equal', value1: 'noon' })).toBe('Enter valid times (hh:mm).');
    expect(error({ kind: 'text', textCriteria: 'beginsWith' })).toBe('Enter the text to compare with.');
    expect(error({ kind: 'text', textCriteria: 'email' })).toBeNull();
    expect(error({ kind: 'list', value1: ' ' })).toMatch(/list items/);
    expect(error({ kind: 'formula', value1: '=' })).toMatch(/Enter a formula/);
  });
});

describe('chart dialogs', () => {
  it('inserts a chart next to its data, or below it on phones', () => {
    const d = salesDoc();
    const res = insertChart(d, 'A1:D4', 'column', false);
    if ('error' in res) throw new Error(res.error);
    expect(res.chart).toMatchObject({ type: 'column', name: 'Chart 1', from: { r: 0, c: 4 } });
    const phone = insertChart(d, 'A1:D4', 'line', true);
    if ('error' in phone) throw new Error(phone.error);
    expect(phone.chart).toMatchObject({ type: 'line', name: 'Chart 2', from: { r: 4, c: 0 } });
    expect(d.sheet.charts).toHaveLength(2);
  });

  it('refuses ranges that are not valid or have no numbers', () => {
    const d = salesDoc();
    expect(insertChart(d, 'A1:', 'column', false)).toEqual({ error: 'Invalid data range.' });
    expect(insertChart(d, 'A1:A4', 'column', false)).toEqual({ error: 'The selected range has no numbers to plot.' });
    expect(d.sheet.charts).toHaveLength(0);
  });

  it('reads the settings of a chart and writes changes back', () => {
    const d = salesDoc();
    const res = insertChart(d, 'A1:D4', 'column', false);
    if ('error' in res || !res.chart) throw new Error('no chart');
    const settings = chartSettingsOf(d, res.chart, d.sheet);
    expect(settings.custom).toBe(false);
    expect(settings.config).toMatchObject({ type: 'column', headers: true });
    expect(settings.seriesStyles).toHaveLength(res.chart.series.length);

    // Series in the other direction: one series per row instead of per column, or the reverse.
    const other = settings.config.seriesIn === 'rows' ? 'columns' : 'rows';
    const { config, laidOut } = relayoutChart(d, d.sheet, settings.config, { seriesIn: other });
    expect(laidOut).toBe(true);
    expect(config.seriesIn).toBe(other);
    expect(config.series).toHaveLength(3);
    expect(relayoutChart(d, d.sheet, settings.config, { range: 'nonsense' }).laidOut).toBe(false);

    const seriesStyles = withSeriesStyle(settings.seriesStyles, 0, { color: '#FF0000', lineWidth: 3 });
    const target = { sheet: d.sheet, range: parseRange('A1:D4')! };
    const patch = chartSettingsPatch(res.chart, { ...settings, config: { ...config, type: 'line' }, style: { ...settings.style, title: '  Sales  ', yMax: '200' }, seriesStyles }, target);
    expect(patch).toMatchObject({ type: 'line', title: 'Sales', yMax: 200, yMin: undefined, seriesIn: other });
    expect(patch.series).toHaveLength(3);
    expect(patch.series?.[0]).toMatchObject({ color: '#FF0000', lineWidth: 3 });
    expect(parseNum(' ')).toBeUndefined();
    expect(parseNum('1e3')).toBe(1000);
    expect(withSeriesStyle([], 2, { markers: true })).toEqual([{}, {}, { markers: true }]);
  });
});
