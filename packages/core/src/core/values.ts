/** Spreadsheet values shared by the model, the formula engine and the file formats. */

export const ERROR_CODES = [
  '#NULL!',
  '#DIV/0!',
  '#VALUE!',
  '#REF!',
  '#NAME?',
  '#NUM!',
  '#N/A',
  '#GETTING_DATA',
  '#SPILL!',
  '#CALC!',
] as const;

export type ErrorCode = (typeof ERROR_CODES)[number];

/** An error value such as #DIV/0!. Instances are shared, so compare with `===` or `.code`. */
export class CellError {
  private constructor(readonly code: ErrorCode) {}

  private static readonly cache = new Map<string, CellError>();

  static of(code: ErrorCode): CellError {
    let e = CellError.cache.get(code);
    if (!e) {
      e = new CellError(code);
      CellError.cache.set(code, e);
    }
    return e;
  }

  toString(): string {
    return this.code;
  }
}

export const Errors = {
  NULL: CellError.of('#NULL!'),
  DIV0: CellError.of('#DIV/0!'),
  VALUE: CellError.of('#VALUE!'),
  REF: CellError.of('#REF!'),
  NAME: CellError.of('#NAME?'),
  NUM: CellError.of('#NUM!'),
  NA: CellError.of('#N/A'),
  DATA: CellError.of('#GETTING_DATA'),
  SPILL: CellError.of('#SPILL!'),
  CALC: CellError.of('#CALC!'),
} as const;

export function parseError(text: string): CellError | null {
  const upper = text.trim().toUpperCase();
  return (ERROR_CODES as readonly string[]).includes(upper) ? CellError.of(upper as ErrorCode) : null;
}

export function isError(v: unknown): v is CellError {
  return v instanceof CellError;
}

/** A cell value. `null` is an empty cell. */
export type Scalar = number | string | boolean | CellError | null;

export function scalarToString(v: Scalar): string {
  if (v === null) return '';
  if (typeof v === 'boolean') return v ? 'TRUE' : 'FALSE';
  if (typeof v === 'number') return numberToString(v);
  if (v instanceof CellError) return v.code;
  return v;
}

/** Number to text the way Excel's General format writes it into formulas and text functions. */
export function numberToString(n: number): string {
  if (!Number.isFinite(n)) return '#NUM!';
  if (Object.is(n, -0) || n === 0) return '0';
  const abs = Math.abs(n);
  if (abs >= 1e21 || abs < 1e-9) {
    // Excel uses 15 significant digits in scientific notation.
    const s = n.toExponential(14).replace(/\.?0+e/, 'e');
    const [mant, exp] = s.split('e');
    const e = Number(exp);
    return `${mant}E${e < 0 ? '-' : '+'}${String(Math.abs(e)).padStart(2, '0')}`;
  }
  const p = Number(n.toPrecision(15));
  return String(p);
}
