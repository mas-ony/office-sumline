/** Runtime values of the formula evaluator and Excel's coercion rules. */
import type { Range } from '../refs';
import type { Sheet } from '../model/sheet';
import { CellError, Errors, numberToString, type Scalar } from '../values';
import { parseDateText, parseTimeText } from './dates';

/** A 2D array of values (rows of columns). */
export class ArrayValue {
  constructor(readonly data: Scalar[][]) {}

  get rows(): number {
    return this.data.length;
  }

  get cols(): number {
    return this.data[0]?.length ?? 0;
  }

  get(r: number, c: number): Scalar {
    return this.data[r]?.[c] ?? null;
  }

  static fill(rows: number, cols: number, v: Scalar): ArrayValue {
    return new ArrayValue(Array.from({ length: rows }, () => new Array<Scalar>(cols).fill(v)));
  }

  static column(values: Scalar[]): ArrayValue {
    return new ArrayValue(values.map((v) => [v]));
  }

  static row(values: Scalar[]): ArrayValue {
    return new ArrayValue([values]);
  }
}

export interface Area {
  sheet: Sheet;
  range: Range;
}

/** A reference to one or more rectangular areas (unions give several). */
export class RefValue {
  constructor(readonly areas: Area[]) {}

  get single(): Area {
    return this.areas[0];
  }

  get rows(): number {
    const r = this.areas[0].range;
    return r.r1 - r.r0 + 1;
  }

  get cols(): number {
    const r = this.areas[0].range;
    return r.c1 - r.c0 + 1;
  }

  isCell(): boolean {
    if (this.areas.length !== 1) return false;
    const r = this.areas[0].range;
    return r.r0 === r.r1 && r.c0 === r.c1;
  }
}

export type Value = Scalar | ArrayValue | RefValue;

/** A deferred argument for functions with lazy evaluation (IF, IFERROR, CHOOSE, ...). */
export type Thunk = () => Value;

export function isArrayLike(v: Value): v is ArrayValue | RefValue {
  return v instanceof ArrayValue || v instanceof RefValue;
}

export function isBlank(v: Scalar): boolean {
  return v === null;
}

/** Excel's number coercion for operators: text is parsed, booleans become 0/1, blanks 0. */
export function toNumber(v: Scalar): number | CellError {
  if (v === null) return 0;
  if (typeof v === 'number') return v;
  if (typeof v === 'boolean') return v ? 1 : 0;
  if (v instanceof CellError) return v;
  const n = textToNumber(v);
  return n === null ? Errors.VALUE : n;
}

/** Parses numeric text the way Excel coerces strings ("1,000", "10%", "$5", dates, times). */
export function textToNumber(text: string): number | null {
  const t = text.trim();
  if (!t) return null;
  const plain = /^[+-]?(\d+\.?\d*|\.\d+)([eE][+-]?\d+)?$/.exec(t);
  if (plain) return Number(t);
  let s = t;
  let neg = false;
  if (/^\(.*\)$/.test(s)) {
    neg = true;
    s = s.slice(1, -1).trim();
  }
  if (s.startsWith('-')) {
    neg = !neg;
    s = s.slice(1).trim();
  } else if (s.startsWith('+')) s = s.slice(1).trim();
  if (s.startsWith('$')) s = s.slice(1).trim();
  let pct = false;
  if (s.endsWith('%')) {
    pct = true;
    s = s.slice(0, -1).trim();
  }
  if (/^(\d{1,3}(,\d{3})+|\d+)(\.\d*)?([eE][+-]?\d+)?$/.test(s) || /^\.\d+$/.test(s)) {
    let n = Number(s.replace(/,/g, ''));
    if (pct) n /= 100;
    return neg ? -n : n;
  }
  if (!pct && !neg) {
    const d = parseDateText(t);
    if (d !== null) return d;
    const tm = parseTimeText(t);
    if (tm !== null) return tm;
  }
  return null;
}

export function toBoolean(v: Scalar): boolean | CellError {
  if (v === null) return false;
  if (typeof v === 'boolean') return v;
  if (typeof v === 'number') return v !== 0;
  if (v instanceof CellError) return v;
  const u = v.trim().toUpperCase();
  if (u === 'TRUE') return true;
  if (u === 'FALSE') return false;
  return Errors.VALUE;
}

export function toText(v: Scalar): string | CellError {
  if (v instanceof CellError) return v;
  if (v === null) return '';
  if (typeof v === 'boolean') return v ? 'TRUE' : 'FALSE';
  if (typeof v === 'number') return numberToString(v);
  return v;
}

/** Compares two scalars with Excel's ordering: numbers < text < booleans; text case-insensitive. */
export function compareScalars(a: Scalar, b: Scalar): number {
  const rank = (v: Scalar) => (v === null ? 0 : typeof v === 'number' ? 1 : typeof v === 'string' ? 2 : typeof v === 'boolean' ? 3 : 4);
  // Blank compares as 0 / "" / FALSE depending on the other side.
  if (a === null && b === null) return 0;
  if (a === null) a = typeof b === 'string' ? '' : typeof b === 'boolean' ? false : 0;
  if (b === null) b = typeof a === 'string' ? '' : typeof a === 'boolean' ? false : 0;
  const ra = rank(a);
  const rb = rank(b);
  if (ra !== rb) return ra - rb;
  if (typeof a === 'number' && typeof b === 'number') return a === b ? 0 : a < b ? -1 : 1;
  if (typeof a === 'string' && typeof b === 'string') {
    const x = a.toLowerCase();
    const y = b.toLowerCase();
    return x === y ? 0 : x < y ? -1 : 1;
  }
  if (typeof a === 'boolean' && typeof b === 'boolean') return a === b ? 0 : a ? 1 : -1;
  return 0;
}

export function firstError(...vals: unknown[]): CellError | null {
  for (const v of vals) if (v instanceof CellError) return v;
  return null;
}

/** Rounds tiny floating point noise the way Excel displays and compares results (15 digits). */
export function round15(n: number): number {
  if (!Number.isFinite(n) || n === 0) return n;
  return Number(n.toPrecision(15));
}
