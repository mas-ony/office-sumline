/**
 * Data validation: which rule applies to a cell, whether a value satisfies it (Excel semantics for
 * whole numbers, decimals, lists, dates, times, text length and custom formulas) and the items of a
 * list rule.
 */
import type { DataValidation, Sheet } from './model/sheet';
import type { CalcEngine } from './formula/engine';
import { translateFormula } from './formula/adjust';
import { ArrayValue, RefValue } from './formula/runtime';
import { CellError, Errors, type Scalar } from './values';

export function validationAt(sheet: Sheet, r: number, c: number): DataValidation | undefined {
  for (let i = sheet.validations.length - 1; i >= 0; i--) {
    const v = sheet.validations[i];
    if (v.sqref.some((rg) => r >= rg.r0 && r <= rg.r1 && c >= rg.c0 && c <= rg.c1)) return v;
  }
  return undefined;
}

/** Top-left cell of a rule's first range: relative references in its formulas are written for it. */
function origin(v: DataValidation): { r: number; c: number } {
  const first = v.sqref[0] ?? { r0: 0, c0: 0 };
  let r = first.r0;
  let c = first.c0;
  for (const rg of v.sqref) {
    if (rg.r0 < r || (rg.r0 === r && rg.c0 < c)) {
      r = rg.r0;
      c = rg.c0;
    }
  }
  return { r, c };
}

function evaluate(engine: CalcEngine, sheet: Sheet, formula: string | undefined, v: DataValidation, r: number, c: number): Scalar | Scalar[] {
  if (formula === undefined || formula === '') return null;
  const o = origin(v);
  const f = translateFormula(formula.replace(/^=/, ''), r - o.r, c - o.c);
  try {
    const res = engine.evaluateText(f, sheet, r, c);
    if (res instanceof RefValue) {
      const out: Scalar[] = [];
      for (const area of res.areas) {
        for (let rr = area.range.r0; rr <= Math.min(area.range.r1, area.range.r0 + 5000); rr++) {
          for (let cc = area.range.c0; cc <= area.range.c1; cc++) out.push(engine.cellValue(area.sheet, rr, cc));
        }
      }
      return out.length === 1 ? out[0] : out;
    }
    if (res instanceof ArrayValue) {
      const flat = res.data.flat();
      return flat.length === 1 ? flat[0] : flat;
    }
    return res as Scalar;
  } catch {
    return Errors.VALUE;
  }
}

function toNumber(v: Scalar | Scalar[]): number | null {
  if (Array.isArray(v)) return toNumber(v[0] ?? null);
  if (typeof v === 'number') return v;
  if (typeof v === 'boolean') return v ? 1 : 0;
  if (typeof v === 'string' && v.trim() !== '' && Number.isFinite(Number(v))) return Number(v);
  return null;
}

function compare(op: string | undefined, x: number, a: number | null, b: number | null): boolean {
  if (a === null) return true;
  switch (op ?? 'between') {
    case 'between':
      return b === null ? x >= a : x >= Math.min(a, b) && x <= Math.max(a, b);
    case 'notBetween':
      return b === null ? x < a : x < Math.min(a, b) || x > Math.max(a, b);
    case 'equal':
      return Math.abs(x - a) < 1e-9;
    case 'notEqual':
      return Math.abs(x - a) >= 1e-9;
    case 'greaterThan':
      return x > a;
    case 'lessThan':
      return x < a;
    case 'greaterThanOrEqual':
      return x >= a;
    case 'lessThanOrEqual':
      return x <= a;
    default:
      return true;
  }
}

/** The items of a list rule: a quoted, comma separated text or the values of a range. */
export function listItems(engine: CalcEngine, sheet: Sheet, v: DataValidation, r: number, c: number, display?: (value: Scalar, sheet: Sheet, r: number, c: number) => string): string[] {
  const f = (v.formula1 ?? '').trim();
  if (!f) return [];
  if (f.startsWith('"')) return f.slice(1, f.endsWith('"') ? -1 : undefined).split(',').map((s) => s.trim()).filter((s) => s !== '');
  const o = origin(v);
  const formula = translateFormula(f.replace(/^=/, ''), r - o.r, c - o.c);
  const out: string[] = [];
  try {
    const res = engine.evaluateText(formula, sheet, r, c);
    if (res instanceof RefValue) {
      for (const area of res.areas) {
        for (let rr = area.range.r0; rr <= Math.min(area.range.r1, area.range.r0 + 1000); rr++) {
          for (let cc = area.range.c0; cc <= area.range.c1; cc++) {
            const val = engine.cellValue(area.sheet, rr, cc);
            if (val === null || val === '') continue;
            out.push(display ? display(val, area.sheet, rr, cc) : String(val));
          }
        }
      }
    } else if (res instanceof ArrayValue) {
      for (const val of res.data.flat()) if (val !== null && val !== '') out.push(String(val));
    } else if (res !== null && !(res instanceof CellError)) out.push(String(res));
  } catch {
    // An invalid list formula offers nothing.
  }
  return out;
}

/** Whether `value` (already parsed, as it would be stored) satisfies the rule at (r, c). */
export function satisfies(engine: CalcEngine, sheet: Sheet, v: DataValidation, r: number, c: number, value: Scalar, display?: (value: Scalar, sheet: Sheet, r: number, c: number) => string): boolean {
  const type = v.type ?? 'none';
  if (type === 'none') return true;
  if (value === null || value === '') {
    if (v.allowBlank) return true;
    return type === 'custom' ? customOk(engine, sheet, v, r, c) : false;
  }
  switch (type) {
    case 'whole':
    case 'decimal':
    case 'date':
    case 'time': {
      const x = typeof value === 'number' ? value : null;
      if (x === null) return false;
      if (type === 'whole' && !Number.isInteger(x)) return false;
      // Times compare their fraction of a day (a date-time keeps its time part).
      const a = toNumber(evaluate(engine, sheet, v.formula1, v, r, c));
      const b = toNumber(evaluate(engine, sheet, v.formula2, v, r, c));
      const subject = type === 'time' ? x - Math.floor(x) : x;
      const a2 = type === 'time' && a !== null ? a - Math.floor(a) : a;
      const b2 = type === 'time' && b !== null ? b - Math.floor(b) : b;
      return compare(v.operator, subject, a2, b2);
    }
    case 'textLength': {
      const text = typeof value === 'string' ? value : display ? display(value, sheet, r, c) : String(value);
      const a = toNumber(evaluate(engine, sheet, v.formula1, v, r, c));
      const b = toNumber(evaluate(engine, sheet, v.formula2, v, r, c));
      return compare(v.operator, text.length, a, b);
    }
    case 'list': {
      const items = listItems(engine, sheet, v, r, c, display).map((s) => s.toLowerCase());
      const text = (typeof value === 'string' ? value : display ? display(value, sheet, r, c) : String(value)).toLowerCase();
      return items.includes(text) || (typeof value === 'number' && items.includes(String(value)));
    }
    case 'custom':
      return customOk(engine, sheet, v, r, c);
    default:
      return true;
  }
}

function customOk(engine: CalcEngine, sheet: Sheet, v: DataValidation, r: number, c: number): boolean {
  const res = evaluate(engine, sheet, v.formula1, v, r, c);
  const x = Array.isArray(res) ? res[0] : res;
  if (x instanceof CellError) return false;
  if (typeof x === 'boolean') return x;
  if (typeof x === 'number') return x !== 0;
  return false;
}

/** A short description of a rule for lists (e.g. "Number between 1 and 10"). */
export function describeValidation(v: DataValidation): string {
  const ops: Record<string, string> = {
    between: 'between', notBetween: 'not between', equal: 'equal to', notEqual: 'different to', greaterThan: 'bigger than',
    lessThan: 'less than', greaterThanOrEqual: 'bigger than or equal to', lessThanOrEqual: 'less than or equal to',
  };
  const kind: Record<string, string> = {
    whole: 'Whole number', decimal: 'Number', date: 'Date', time: 'Time', textLength: 'Text length', list: 'List', custom: 'Formula',
  };
  const t = kind[v.type ?? ''] ?? 'Any value';
  if (v.type === 'list') return `List: ${(v.formula1 ?? '').replace(/^"|"$/g, '')}`;
  if (v.type === 'custom') return `Formula: =${v.formula1 ?? ''}`;
  const op = ops[v.operator ?? 'between'] ?? v.operator ?? '';
  const two = v.operator === undefined || v.operator === 'between' || v.operator === 'notBetween';
  return `${t} ${op} ${v.formula1 ?? ''}${two && v.formula2 !== undefined ? ` and ${v.formula2}` : ''}`.trim();
}

const OP_SYMBOL: Record<string, string> = {
  equal: '=', notEqual: '<>', greaterThan: '>', lessThan: '<', greaterThanOrEqual: '>=', lessThanOrEqual: '<=',
};

/**
 * A formula that is TRUE when the top-left cell of the rule (`cell`, e.g. "B2") holds a valid value.
 * Used to highlight invalid cells with a conditional format ("Format style" action).
 */
export function validityFormula(v: DataValidation, cell: string): string | null {
  const f1 = v.formula1 ?? '';
  const f2 = v.formula2 ?? '';
  const cmp = (subject: string) => {
    const op = v.operator ?? 'between';
    if (op === 'between') return `AND(${subject}>=${f1},${subject}<=${f2 || f1})`;
    if (op === 'notBetween') return `OR(${subject}<${f1},${subject}>${f2 || f1})`;
    return `${subject}${OP_SYMBOL[op] ?? '='}${f1}`;
  };
  switch (v.type) {
    case 'whole':
      return `AND(ISNUMBER(${cell}),INT(${cell})=${cell},${cmp(cell)})`;
    case 'decimal':
    case 'date':
      return `AND(ISNUMBER(${cell}),${cmp(cell)})`;
    case 'time':
      return `AND(ISNUMBER(${cell}),${cmp(`MOD(${cell},1)`)})`;
    case 'textLength':
      return cmp(`LEN(${cell})`);
    case 'list': {
      const f = f1.trim();
      if (f.startsWith('"')) {
        const items = f.replace(/^"|"$/g, '').split(',').map((s) => `"${s.trim().replace(/"/g, '""')}"`);
        return `ISNUMBER(MATCH(${cell},{${items.join(',')}},0))`;
      }
      return `ISNUMBER(MATCH(${cell},${f},0))`;
    }
    case 'custom':
      return f1 || null;
    default:
      return null;
  }
}
