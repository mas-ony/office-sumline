/**
 * The Validations dialog's rule form: read from a data validation rule, and the rule it makes
 * (number, date, time and text length criteria, text rules as formulas, lists, empty / not empty,
 * custom formulas; a message, or a format style).
 */
import { formatValue } from '../core/format/display';
import { parseInput } from '../core/format/input';
import type { DataValidation, ValidationFormatStyle } from '../core/model/sheet';
import { parseRange, quoteSheetName, type Range, rangeName } from '../core/refs';

export type ValidationAction = 'warning' | 'reject' | 'format';
export type ValidationKind = 'number' | 'text' | 'date' | 'time' | 'list' | 'textLength' | 'empty' | 'notEmpty' | 'formula';
export type ValidationCriteria = 'between' | 'notBetween' | 'lessThan' | 'lessThanOrEqual' | 'greaterThan' | 'greaterThanOrEqual' | 'equal' | 'notEqual';
export type TextCriteria = 'contains' | 'notContains' | 'beginsWith' | 'endsWith' | 'equal' | 'notEqual' | 'email' | 'url';

export const VALIDATION_KINDS: { key: ValidationKind; label: string }[] = [
  { key: 'number', label: 'Number' },
  { key: 'text', label: 'Text' },
  { key: 'date', label: 'Date' },
  { key: 'time', label: 'Time' },
  { key: 'list', label: 'List' },
  { key: 'textLength', label: 'Text length' },
  { key: 'empty', label: 'Cell is empty' },
  { key: 'notEmpty', label: 'Cell is not empty' },
  { key: 'formula', label: 'Formula' },
];

export const VALIDATION_CRITERIA: { key: ValidationCriteria; label: string }[] = [
  { key: 'between', label: 'between' },
  { key: 'notBetween', label: 'is not between' },
  { key: 'lessThan', label: 'less than' },
  { key: 'lessThanOrEqual', label: 'less than or equal to' },
  { key: 'greaterThan', label: 'bigger than' },
  { key: 'greaterThanOrEqual', label: 'bigger than or equal to' },
  { key: 'equal', label: 'equal to' },
  { key: 'notEqual', label: 'different to' },
];

export const TEXT_CRITERIA: { key: TextCriteria; label: string }[] = [
  { key: 'contains', label: 'contains' },
  { key: 'notContains', label: 'not contains' },
  { key: 'beginsWith', label: 'begins with' },
  { key: 'endsWith', label: 'ends with' },
  { key: 'equal', label: 'equal to' },
  { key: 'notEqual', label: 'different to' },
  { key: 'email', label: 'valid email' },
  { key: 'url', label: 'valid URL' },
];

/** The rule form of the Validations dialog. */
export interface ValidationForm {
  range: string;
  action: ValidationAction;
  message: string;
  kind: ValidationKind;
  criteria: ValidationCriteria;
  textCriteria: TextCriteria;
  value1: string;
  value2: string;
  allowBlank: boolean;
  dropdown: boolean;
  style: ValidationFormatStyle;
}

const quote = (s: string) => `"${s.replace(/"/g, '""')}"`;

/** Text-rule formulas: built from the first cell of the range, recognised again when editing. */
function textFormula(c: TextCriteria, cell: string, value: string): string {
  const v = quote(value);
  switch (c) {
    case 'contains': return `ISNUMBER(SEARCH(${v},${cell}))`;
    case 'notContains': return `ISERROR(SEARCH(${v},${cell}))`;
    case 'beginsWith': return `LEFT(${cell},LEN(${v}))=${v}`;
    case 'endsWith': return `RIGHT(${cell},LEN(${v}))=${v}`;
    case 'equal': return `EXACT(${cell},${v})`;
    case 'notEqual': return `NOT(EXACT(${cell},${v}))`;
    case 'email': return `AND(ISNUMBER(SEARCH("@",${cell})),ISNUMBER(SEARCH(".",${cell},SEARCH("@",${cell}))))`;
    case 'url': return `OR(LEFT(${cell},7)="http://",LEFT(${cell},8)="https://")`;
  }
}

function parseTextFormula(f: string): { criteria: TextCriteria; value: string } | null {
  const val = '"((?:[^"]|"")*)"';
  const cell = '\\$?[A-Z]{1,3}\\$?\\d+';
  const tests: [TextCriteria, RegExp][] = [
    ['contains', new RegExp(`^ISNUMBER\\(SEARCH\\(${val},${cell}\\)\\)$`)],
    ['notContains', new RegExp(`^ISERROR\\(SEARCH\\(${val},${cell}\\)\\)$`)],
    ['beginsWith', new RegExp(`^LEFT\\(${cell},LEN\\(${val}\\)\\)=`)],
    ['endsWith', new RegExp(`^RIGHT\\(${cell},LEN\\(${val}\\)\\)=`)],
    ['notEqual', new RegExp(`^NOT\\(EXACT\\(${cell},${val}\\)\\)$`)],
    ['equal', new RegExp(`^EXACT\\(${cell},${val}\\)$`)],
  ];
  for (const [criteria, re] of tests) {
    const m = re.exec(f);
    if (m) return { criteria, value: m[1].replace(/""/g, '"') };
  }
  if (/^AND\(ISNUMBER\(SEARCH\("@"/.test(f)) return { criteria: 'email', value: '' };
  if (/^OR\(LEFT\(.*"http:\/\/"/.test(f)) return { criteria: 'url', value: '' };
  return null;
}

/** A value of a rule as the date and time inputs take it. */
function display(n: number, kind: ValidationKind, date1904: boolean): string {
  if (kind === 'date') return formatValue(n, 'yyyy-mm-dd', date1904).text;
  if (kind === 'time') return formatValue(n, 'hh:mm:ss').text;
  return String(n);
}

/** The form for a rule of the sheet `sheetName`. */
export function validationFormOf(v: DataValidation, sheetName: string, date1904 = false): ValidationForm {
  const range = v.sqref.map((r) => `${quoteSheetName(sheetName)}!${rangeName(r)}`).join(' ');
  const base: ValidationForm = {
    range,
    action: v.formatStyle ? 'format' : v.errorStyle === 'stop' ? 'reject' : 'warning',
    message: v.error ?? '',
    kind: 'number',
    criteria: (v.operator as ValidationCriteria) ?? 'between',
    textCriteria: 'contains',
    value1: v.formula1 ?? '',
    value2: v.formula2 ?? '',
    allowBlank: !!v.allowBlank,
    dropdown: !v.hideDropDown,
    style: v.formatStyle ?? { color: '#C62828', background: '#FFEBEE' },
  };
  const num = (s: string) => Number(s);
  switch (v.type) {
    case 'whole':
    case 'decimal':
      return { ...base, kind: 'number' };
    case 'date':
      return { ...base, kind: 'date', value1: Number.isFinite(num(base.value1)) ? display(num(base.value1), 'date', date1904) : base.value1, value2: Number.isFinite(num(base.value2)) && base.value2 ? display(num(base.value2), 'date', date1904) : base.value2 };
    case 'time':
      return { ...base, kind: 'time', value1: Number.isFinite(num(base.value1)) ? display(num(base.value1), 'time', date1904) : base.value1, value2: Number.isFinite(num(base.value2)) && base.value2 ? display(num(base.value2), 'time', date1904) : base.value2 };
    case 'textLength':
      return { ...base, kind: 'textLength' };
    case 'list':
      return { ...base, kind: 'list', value1: (v.formula1 ?? '').startsWith('"') ? (v.formula1 ?? '').replace(/^"|"$/g, '') : `=${v.formula1 ?? ''}` };
    case 'custom': {
      const f = v.formula1 ?? '';
      if (/^LEN\([^)]*\)=0$/.test(f)) return { ...base, kind: 'empty' };
      if (/^LEN\([^)]*\)>0$/.test(f)) return { ...base, kind: 'notEmpty' };
      const t = parseTextFormula(f);
      if (t) return { ...base, kind: 'text', textCriteria: t.criteria, value1: t.value };
      return { ...base, kind: 'formula', value1: `=${f}` };
    }
    default:
      return base;
  }
}

/** The form of a new rule for the selection. */
export function blankValidationForm(sheetName: string, selection: Range): ValidationForm {
  return {
    range: `${quoteSheetName(sheetName)}!${rangeName(selection)}`,
    action: 'warning',
    message: '',
    kind: 'number',
    criteria: 'between',
    textCriteria: 'contains',
    value1: '',
    value2: '',
    allowBlank: true,
    dropdown: true,
    style: { color: '#C62828', background: '#FFEBEE' },
  };
}

/** Ranges of the sheet `sheetName` ("A1:A10 C1", optionally with the sheet's name). */
function parseRanges(text: string, sheetName: string): Range[] | null {
  // Separated by spaces, commas or semicolons; a quoted sheet name may contain them.
  const parts = text.match(/(?:'(?:[^']|'')*'!)?[^\s,;]+/g) ?? [];
  const out: Range[] = [];
  for (const p of parts) {
    const m = /^(?:'((?:[^']|'')+)'|([^!]+))!(.+)$/.exec(p);
    const name = m ? (m[1] ?? m[2]).replace(/''/g, "'") : undefined;
    if (name && name.toLowerCase() !== sheetName.toLowerCase()) return null;
    const r = parseRange((m ? m[3] : p).toUpperCase());
    if (!r) return null;
    out.push(r);
  }
  return out.length ? out : null;
}

function toNumberText(text: string, kind: ValidationKind, date1904: boolean): string | null {
  const t = text.trim();
  if (t.startsWith('=')) return t.slice(1);
  if (kind === 'number' || kind === 'textLength') return t !== '' && Number.isFinite(Number(t)) ? String(Number(t)) : null;
  const parsed = parseInput(t, { date1904 });
  if (typeof parsed.value === 'number') return String(kind === 'time' ? parsed.value - Math.floor(parsed.value) : parsed.value);
  return null;
}

/** The rule the form makes for the sheet `sheetName`, or what is wrong with the form. */
export function ruleFromForm(form: ValidationForm, sheetName: string, date1904: boolean): { rule: DataValidation } | { error: string } {
  const sqref = parseRanges(form.range, sheetName);
  if (!sqref) return { error: 'Enter a cell range of this worksheet, for example A1:A10.' };
  const first = rangeName({ r0: sqref[0].r0, c0: sqref[0].c0, r1: sqref[0].r0, c1: sqref[0].c0 });
  const rule: DataValidation = {
    sqref,
    allowBlank: form.allowBlank,
    showErrorMessage: form.action !== 'format',
    errorStyle: form.action === 'reject' ? 'stop' : form.action === 'warning' ? 'warning' : 'information',
    error: form.message || undefined,
    errorTitle: form.message ? (form.action === 'reject' ? 'Invalid entry' : 'Warning') : undefined,
    formatStyle: form.action === 'format' ? form.style : undefined,
  };
  const two = form.criteria === 'between' || form.criteria === 'notBetween';
  switch (form.kind) {
    case 'number':
    case 'date':
    case 'time':
    case 'textLength': {
      const a = toNumberText(form.value1, form.kind, date1904);
      const b = two ? toNumberText(form.value2, form.kind, date1904) : undefined;
      if (a === null || (two && (b === null || b === undefined))) {
        return { error: form.kind === 'date' ? 'Enter valid dates.' : form.kind === 'time' ? 'Enter valid times (hh:mm).' : 'Enter numbers or formulas starting with =.' };
      }
      rule.type = form.kind === 'number' ? 'decimal' : form.kind;
      rule.operator = form.criteria;
      rule.formula1 = a;
      if (two) rule.formula2 = b ?? undefined;
      break;
    }
    case 'text': {
      if (!form.value1 && form.textCriteria !== 'email' && form.textCriteria !== 'url') return { error: 'Enter the text to compare with.' };
      rule.type = 'custom';
      rule.formula1 = textFormula(form.textCriteria, first, form.value1);
      break;
    }
    case 'list': {
      const src = form.value1.trim();
      if (!src) return { error: 'Enter the list items separated by commas, or a range such as =$D$1:$D$5.' };
      rule.type = 'list';
      rule.formula1 = src.startsWith('=') ? src.slice(1) : `"${src.split(',').map((s) => s.trim()).filter(Boolean).join(',').replace(/"/g, '')}"`;
      rule.hideDropDown = !form.dropdown;
      break;
    }
    case 'empty':
      rule.type = 'custom';
      rule.formula1 = `LEN(${first})=0`;
      rule.allowBlank = true;
      break;
    case 'notEmpty':
      rule.type = 'custom';
      rule.formula1 = `LEN(${first})>0`;
      rule.allowBlank = false;
      break;
    case 'formula': {
      const f = form.value1.trim().replace(/^=/, '');
      if (!f) return { error: 'Enter a formula, for example =A1>0.' };
      rule.type = 'custom';
      rule.formula1 = f;
      break;
    }
  }
  return { rule };
}
