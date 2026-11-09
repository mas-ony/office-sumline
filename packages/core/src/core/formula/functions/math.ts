import { CellError, Errors, type Scalar } from '../../values';
import { ArrayValue, RefValue, toNumber, type Value } from '../runtime';
import { makeCriteria } from '../criteria';
import {
  collectNumbers, define, defineScalar, alias, err, forEachArg, num, toArray, toScalar, numOrError,
  type FnContext,
} from './helpers';
import {
  combin, factorial, gcd2, kth, mean, median, modeSingle, percentileExc, percentileInc, quartile, roundDown,
  roundHalfAway, roundUp, stdev, sum, variance,
} from './numeric';

type NumFn = (...xs: number[]) => number | CellError;

/** Defines a function of numeric scalars; missing optional arguments arrive as undefined. */
function numeric(name: string, min: number, max: number, f: (xs: (number | undefined)[]) => number | CellError): void {
  defineScalar(name, min, max, (args) => {
    const xs: (number | undefined)[] = [];
    for (let i = 0; i < args.length; i++) {
      const a = args[i];
      if (a === undefined) {
        xs.push(undefined);
        continue;
      }
      const n = toNumber(a);
      if (n instanceof CellError) return n;
      xs.push(n);
    }
    const r = f(xs);
    return typeof r === 'number' ? numOrError(r) : r;
  });
}

const unary = (name: string, f: NumFn) => numeric(name, 1, 1, ([x]) => f(x!));

unary('ABS', Math.abs);
unary('ACOS', (x) => (x < -1 || x > 1 ? Errors.NUM : Math.acos(x)));
unary('ACOSH', (x) => (x < 1 ? Errors.NUM : Math.acosh(x)));
unary('ACOT', (x) => Math.PI / 2 - Math.atan(x));
unary('ACOTH', (x) => (Math.abs(x) <= 1 ? Errors.NUM : 0.5 * Math.log((x + 1) / (x - 1))));
unary('ASIN', (x) => (x < -1 || x > 1 ? Errors.NUM : Math.asin(x)));
unary('ASINH', Math.asinh);
unary('ATAN', Math.atan);
numeric('ATAN2', 2, 2, ([x, y]) => (x === 0 && y === 0 ? Errors.DIV0 : Math.atan2(y!, x!)));
unary('ATANH', (x) => (Math.abs(x) >= 1 ? Errors.NUM : Math.atanh(x)));
unary('COS', Math.cos);
unary('COSH', Math.cosh);
unary('COT', (x) => (x === 0 ? Errors.DIV0 : 1 / Math.tan(x)));
unary('COTH', (x) => (x === 0 ? Errors.DIV0 : 1 / Math.tanh(x)));
unary('CSC', (x) => (x === 0 ? Errors.DIV0 : 1 / Math.sin(x)));
unary('CSCH', (x) => (x === 0 ? Errors.DIV0 : 1 / Math.sinh(x)));
unary('SEC', (x) => 1 / Math.cos(x));
unary('SECH', (x) => 1 / Math.cosh(x));
unary('SIN', Math.sin);
unary('SINH', Math.sinh);
unary('TAN', Math.tan);
unary('TANH', Math.tanh);
unary('DEGREES', (x) => (x * 180) / Math.PI);
unary('RADIANS', (x) => (x * Math.PI) / 180);
unary('EXP', Math.exp);
unary('LN', (x) => (x <= 0 ? Errors.NUM : Math.log(x)));
unary('LOG10', (x) => (x <= 0 ? Errors.NUM : Math.log10(x)));
numeric('LOG', 1, 2, ([x, b = 10]) => {
  if (x! <= 0 || b <= 0) return Errors.NUM;
  if (b === 1) return Errors.DIV0;
  return Math.log(x!) / Math.log(b);
});
define('PI', { min: 0, max: 0, impl: () => Math.PI });
numeric('POWER', 2, 2, ([x, y]) => power(x!, y!));
unary('SQRT', (x) => (x < 0 ? Errors.NUM : Math.sqrt(x)));
unary('SQRTPI', (x) => (x < 0 ? Errors.NUM : Math.sqrt(x * Math.PI)));
unary('SIGN', (x) => (x > 0 ? 1 : x < 0 ? -1 : 0));
unary('INT', (x) => Math.floor(x));
numeric('TRUNC', 1, 2, ([x, d = 0]) => roundDown(x!, d));
numeric('ROUND', 2, 2, ([x, d]) => roundHalfAway(x!, d ?? 0));
numeric('ROUNDUP', 2, 2, ([x, d]) => roundUp(x!, d ?? 0));
numeric('ROUNDDOWN', 2, 2, ([x, d]) => roundDown(x!, d ?? 0));
numeric('MROUND', 2, 2, ([x, m]) => {
  if (m === 0) return 0;
  if ((x! > 0 && m! < 0) || (x! < 0 && m! > 0)) return Errors.NUM;
  return roundHalfAway(x! / m!, 0) * m!;
});
numeric('CEILING', 1, 2, ([x, s = 1]) => {
  if (s === 0) return 0;
  if (x! > 0 && s < 0) return Errors.NUM;
  return ceilTo(x!, s, false);
});
numeric('CEILING.MATH', 1, 3, ([x, s = 1, mode = 0]) => {
  if (s === 0) return 0;
  const sig = Math.abs(s);
  if (x! < 0 && mode !== 0) return -ceilTo(-x!, sig, false);
  return ceilTo(x!, sig, false);
});
numeric('CEILING.PRECISE', 1, 2, ([x, s = 1]) => (s === 0 ? 0 : ceilTo(x!, Math.abs(s), false)));
alias('ISO.CEILING', 'CEILING.PRECISE');
numeric('FLOOR', 1, 2, ([x, s]) => {
  if (s === undefined) return Errors.NA;
  if (s === 0) return x === 0 ? 0 : Errors.DIV0;
  if (x! > 0 && s < 0) return Errors.NUM;
  return floorTo(x!, s, false);
});
numeric('FLOOR.MATH', 1, 3, ([x, s = 1, mode = 0]) => {
  if (s === 0) return 0;
  const sig = Math.abs(s);
  if (x! < 0 && mode !== 0) return -floorTo(-x!, sig, false);
  return floorTo(x!, sig, false);
});
numeric('FLOOR.PRECISE', 1, 2, ([x, s = 1]) => (s === 0 ? 0 : floorTo(x!, Math.abs(s), false)));
unary('EVEN', (x) => {
  const v = Math.ceil(Math.abs(x) / 2) * 2;
  return x < 0 ? -v : v;
});
unary('ODD', (x) => {
  let v = Math.ceil(Math.abs(x));
  if (v % 2 === 0) v += 1;
  return x < 0 ? -v : v;
});
unary('FACT', (x) => (x < 0 ? Errors.NUM : factorial(x)));
unary('FACTDOUBLE', (x) => {
  const n = Math.trunc(x);
  if (n < -1) return Errors.NUM;
  let r = 1;
  for (let i = n; i > 1; i -= 2) r *= i;
  return r;
});
numeric('COMBIN', 2, 2, ([n, k]) => {
  const r = combin(n!, k!);
  return Number.isNaN(r) ? Errors.NUM : r;
});
numeric('COMBINA', 2, 2, ([n, k]) => {
  if (n! < 0 || k! < 0) return Errors.NUM;
  if (n === 0 && k === 0) return 1;
  return combin(Math.trunc(n!) + Math.trunc(k!) - 1, Math.trunc(k!));
});
numeric('PERMUT', 2, 2, ([n, k]) => {
  const a = Math.trunc(n!);
  const b = Math.trunc(k!);
  if (a < 0 || b < 0 || b > a) return Errors.NUM;
  let r = 1;
  for (let i = 0; i < b; i++) r *= a - i;
  return r;
});
numeric('PERMUTATIONA', 2, 2, ([n, k]) => (n! < 0 || k! < 0 ? Errors.NUM : Math.pow(Math.trunc(n!), Math.trunc(k!))));
numeric('MOD', 2, 2, ([n, d]) => {
  if (d === 0) return Errors.DIV0;
  const r = n! - d! * Math.floor(n! / d!);
  return Math.abs(r) < 1e-12 ? 0 : Number(r.toPrecision(15));
});
numeric('QUOTIENT', 2, 2, ([n, d]) => (d === 0 ? Errors.DIV0 : Math.trunc(n! / d!)));
define('RAND', { min: 0, max: 0, volatile: true, impl: () => Math.random() });
define('RANDBETWEEN', {
  min: 2,
  max: 2,
  volatile: true,
  scalar: 'all',
  impl: ([a, b]) => {
    const lo = toNumber(a as Scalar);
    const hi = toNumber(b as Scalar);
    const e = err(lo, hi);
    if (e) return e;
    const l = Math.ceil(lo as number);
    const h = Math.floor(hi as number);
    if (l > h) return Errors.NUM;
    return l + Math.floor(Math.random() * (h - l + 1));
  },
});
define('RANDARRAY', {
  min: 0,
  max: 5,
  volatile: true,
  impl: (args, ctx) => {
    const rows = args[0] == null ? 1 : num(args[0], ctx);
    const cols = args[1] == null ? 1 : num(args[1], ctx);
    const lo = args[2] == null ? 0 : num(args[2], ctx);
    const hi = args[3] == null ? 1 : num(args[3], ctx);
    const whole = args[4] == null ? false : toScalar(args[4], ctx) === true || toNumber(toScalar(args[4], ctx)) === 1;
    const e = err(rows, cols, lo, hi);
    if (e) return e;
    const r = Math.trunc(rows as number);
    const c = Math.trunc(cols as number);
    if (r < 1 || c < 1 || (hi as number) < (lo as number)) return Errors.VALUE;
    const data: Scalar[][] = [];
    for (let i = 0; i < r; i++) {
      const row: Scalar[] = [];
      for (let j = 0; j < c; j++) {
        const x = (lo as number) + Math.random() * ((hi as number) - (lo as number));
        row.push(whole ? Math.floor((lo as number) + Math.random() * (Math.floor(hi as number) - Math.ceil(lo as number) + 1)) : x);
      }
      data.push(row);
    }
    return new ArrayValue(data);
  },
});
define('SEQUENCE', {
  min: 1,
  max: 4,
  impl: (args, ctx) => {
    const rows = num(args[0], ctx);
    const cols = args[1] == null ? 1 : num(args[1], ctx);
    const start = args[2] == null ? 1 : num(args[2], ctx);
    const step = args[3] == null ? 1 : num(args[3], ctx);
    const e = err(rows, cols, start, step);
    if (e) return e;
    const r = Math.trunc(rows as number);
    const c = Math.trunc(cols as number);
    if (r < 1 || c < 1) return Errors.CALC;
    if (r * c > 1048576 * 16) return Errors.NUM;
    const data: Scalar[][] = [];
    let v = start as number;
    for (let i = 0; i < r; i++) {
      const row: Scalar[] = [];
      for (let j = 0; j < c; j++) {
        row.push(v);
        v += step as number;
      }
      data.push(row);
    }
    return new ArrayValue(data);
  },
});

const ROMAN_TABLE: [number, string][] = [
  [1000, 'M'], [900, 'CM'], [500, 'D'], [400, 'CD'], [100, 'C'], [90, 'XC'], [50, 'L'], [40, 'XL'], [10, 'X'], [9, 'IX'], [5, 'V'], [4, 'IV'], [1, 'I'],
];
defineScalar('ROMAN', 1, 2, ([x]) => {
  const n = toNumber(x);
  if (n instanceof CellError) return n;
  let v = Math.trunc(n);
  if (v < 0 || v > 3999) return Errors.VALUE;
  let out = '';
  for (const [k, s] of ROMAN_TABLE) {
    while (v >= k) {
      out += s;
      v -= k;
    }
  }
  return out;
});
defineScalar('ARABIC', 1, 1, ([x]) => {
  if (typeof x !== 'string') return Errors.VALUE;
  let s = x.trim().toUpperCase();
  let neg = false;
  if (s.startsWith('-')) {
    neg = true;
    s = s.slice(1);
  }
  if (!/^[MDCLXVI]*$/.test(s)) return Errors.VALUE;
  const val: Record<string, number> = { M: 1000, D: 500, C: 100, L: 50, X: 10, V: 5, I: 1 };
  let total = 0;
  for (let i = 0; i < s.length; i++) {
    const a = val[s[i]];
    const b = val[s[i + 1]] ?? 0;
    total += a < b ? -a : a;
  }
  return neg ? -total : total;
});
defineScalar('BASE', 2, 3, ([x, radix, minLen]) => {
  const n = toNumber(x);
  const r = toNumber(radix);
  const m = minLen === undefined ? 0 : toNumber(minLen);
  const e = err(n, r, m);
  if (e) return e;
  if ((n as number) < 0 || (r as number) < 2 || (r as number) > 36) return Errors.NUM;
  return Math.trunc(n as number).toString(Math.trunc(r as number)).toUpperCase().padStart(Math.trunc(m as number), '0');
});
defineScalar('DECIMAL', 2, 2, ([x, radix]) => {
  const r = toNumber(radix);
  if (r instanceof CellError) return r;
  const s = String(x ?? '').trim();
  const base = Math.trunc(r);
  if (base < 2 || base > 36) return Errors.NUM;
  const n = parseInt(s, base);
  if (Number.isNaN(n) || !new RegExp(`^[${'0123456789abcdefghijklmnopqrstuvwxyz'.slice(0, base)}]+$`, 'i').test(s)) return Errors.NUM;
  return n;
});

define('GCD', {
  min: 1,
  max: 255,
  impl: (args, ctx) => {
    const xs = collectNumbers(args, ctx);
    if (xs instanceof CellError) return xs;
    if (xs.some((x) => x < 0)) return Errors.NUM;
    return xs.reduce((a, b) => gcd2(a, b), 0);
  },
});
define('LCM', {
  min: 1,
  max: 255,
  impl: (args, ctx) => {
    const xs = collectNumbers(args, ctx);
    if (xs instanceof CellError) return xs;
    if (xs.some((x) => x < 0)) return Errors.NUM;
    let l = 1;
    for (const x0 of xs) {
      const x = Math.trunc(x0);
      if (x === 0) return 0;
      l = (l * x) / gcd2(l, x);
    }
    return l;
  },
});
define('MULTINOMIAL', {
  min: 1,
  max: 255,
  impl: (args, ctx) => {
    const xs = collectNumbers(args, ctx);
    if (xs instanceof CellError) return xs;
    if (xs.some((x) => x < 0)) return Errors.NUM;
    let total = 0;
    let denom = 1;
    for (const x of xs) {
      total += Math.trunc(x);
      denom *= factorial(x);
    }
    return factorial(total) / denom;
  },
});

// -----------------------------------------------------------------------------------------------
// Sums and products

define('SUM', {
  min: 1,
  max: 255,
  impl: (args, ctx) => {
    const xs = collectNumbers(args, ctx);
    return xs instanceof CellError ? xs : sum(xs);
  },
});
define('SUMSQ', {
  min: 1,
  max: 255,
  impl: (args, ctx) => {
    const xs = collectNumbers(args, ctx);
    return xs instanceof CellError ? xs : sum(xs.map((x) => x * x));
  },
});
define('PRODUCT', {
  min: 1,
  max: 255,
  impl: (args, ctx) => {
    const xs = collectNumbers(args, ctx);
    if (xs instanceof CellError) return xs;
    if (!xs.length) return 0;
    return numOrError(xs.reduce((a, b) => a * b, 1));
  },
});

function sameShapeArrays(args: Value[], ctx: FnContext): ArrayValue[] | CellError {
  const arrays = args.map((a) => toArray(a, ctx));
  const r = arrays[0].rows;
  const c = arrays[0].cols;
  if (arrays.some((a) => a.rows !== r || a.cols !== c)) return Errors.VALUE;
  return arrays;
}

define('SUMPRODUCT', {
  min: 1,
  max: 255,
  impl: (args, ctx) => {
    const arrays = sameShapeArrays(args, ctx);
    if (arrays instanceof CellError) return arrays;
    let total = 0;
    const rows = arrays[0].rows;
    const cols = arrays[0].cols;
    for (let r = 0; r < rows; r++) {
      for (let c = 0; c < cols; c++) {
        let p = 1;
        for (const a of arrays) {
          const v = a.get(r, c);
          if (v instanceof CellError) return v;
          p *= typeof v === 'number' ? v : 0;
        }
        total += p;
      }
    }
    return total;
  },
});

function pairwise(name: string, f: (x: number, y: number) => number): void {
  define(name, {
    min: 2,
    max: 2,
    impl: (args, ctx) => {
      const arrays = sameShapeArrays(args, ctx);
      if (arrays instanceof CellError) return Errors.NA;
      const [a, b] = arrays;
      let total = 0;
      for (let r = 0; r < a.rows; r++) {
        for (let c = 0; c < a.cols; c++) {
          const x = a.get(r, c);
          const y = b.get(r, c);
          if (x instanceof CellError) return x;
          if (y instanceof CellError) return y;
          if (typeof x === 'number' && typeof y === 'number') total += f(x, y);
        }
      }
      return total;
    },
  });
}
pairwise('SUMX2MY2', (x, y) => x * x - y * y);
pairwise('SUMX2PY2', (x, y) => x * x + y * y);
pairwise('SUMXMY2', (x, y) => (x - y) * (x - y));

define('SERIESSUM', {
  min: 4,
  max: 4,
  impl: (args, ctx) => {
    const x = num(args[0], ctx);
    const n = num(args[1], ctx);
    const m = num(args[2], ctx);
    const e = err(x, n, m);
    if (e) return e;
    const coeffs = collectNumbers([args[3]], ctx);
    if (coeffs instanceof CellError) return coeffs;
    let total = 0;
    coeffs.forEach((a, i) => {
      total += a * Math.pow(x as number, (n as number) + i * (m as number));
    });
    return total;
  },
});

/** Visits the cells of conditional ranges and the matching cells of a value range. */
export function conditionalCells(
  ctx: FnContext,
  pairs: [Value, Value][],
  visit: (index: [number, number]) => void,
): CellError | undefined {
  const ranges = pairs.map(([range]) => toArray(range, ctx));
  const rows = ranges[0].rows;
  const cols = ranges[0].cols;
  if (ranges.some((r) => r.rows !== rows || r.cols !== cols)) return Errors.VALUE;
  const matchers = pairs.map(([, crit]) => makeCriteria(toScalar(crit, ctx)));
  for (let r = 0; r < rows; r++) {
    for (let c = 0; c < cols; c++) {
      let ok = true;
      for (let k = 0; k < ranges.length && ok; k++) ok = matchers[k](ranges[k].get(r, c));
      if (ok) visit([r, c]);
    }
  }
  return undefined;
}

/** Cell of a "sum range" at the offset of a matching criteria cell (Excel resizes it to the criteria range). */
export function offsetValue(sumRange: Value, r: number, c: number, ctx: FnContext): Scalar {
  if (sumRange instanceof RefValue) {
    const { sheet, range } = sumRange.single;
    return ctx.engine.cellValue(sheet, range.r0 + r, range.c0 + c);
  }
  if (sumRange instanceof ArrayValue) return sumRange.get(r, c);
  return r === 0 && c === 0 ? sumRange : null;
}

define('SUMIF', {
  min: 2,
  max: 3,
  impl: (args, ctx) => {
    const sumRange = args[2] ?? args[0];
    let total = 0;
    let error: CellError | undefined;
    const e = conditionalCells(ctx, [[args[0], args[1]]], ([r, c]) => {
      const v = offsetValue(sumRange, r, c, ctx);
      if (v instanceof CellError) error ??= v;
      else if (typeof v === 'number') total += v;
    });
    return e ?? error ?? total;
  },
});
define('SUMIFS', {
  min: 3,
  max: 255,
  impl: (args, ctx) => {
    if (args.length % 2 === 0) return Errors.VALUE;
    const pairs: [Value, Value][] = [];
    for (let i = 1; i < args.length; i += 2) pairs.push([args[i], args[i + 1]]);
    const sumArr = toArray(args[0], ctx);
    const first = toArray(args[1], ctx);
    if (sumArr.rows !== first.rows || sumArr.cols !== first.cols) return Errors.VALUE;
    let total = 0;
    let error: CellError | undefined;
    const e = conditionalCells(ctx, pairs, ([r, c]) => {
      const v = sumArr.get(r, c);
      if (v instanceof CellError) error ??= v;
      else if (typeof v === 'number') total += v;
    });
    return e ?? error ?? total;
  },
});

// -----------------------------------------------------------------------------------------------
// Matrices

function numericMatrix(v: Value, ctx: FnContext): number[][] | CellError {
  const a = toArray(v, ctx);
  const out: number[][] = [];
  for (const row of a.data) {
    const r: number[] = [];
    for (const x of row) {
      if (x instanceof CellError) return x;
      if (typeof x !== 'number') return Errors.VALUE;
      r.push(x);
    }
    out.push(r);
  }
  return out;
}

function determinant(m: number[][]): number {
  const n = m.length;
  const a = m.map((r) => [...r]);
  let det = 1;
  for (let i = 0; i < n; i++) {
    let p = i;
    for (let k = i + 1; k < n; k++) if (Math.abs(a[k][i]) > Math.abs(a[p][i])) p = k;
    if (a[p][i] === 0) return 0;
    if (p !== i) {
      [a[p], a[i]] = [a[i], a[p]];
      det = -det;
    }
    det *= a[i][i];
    for (let k = i + 1; k < n; k++) {
      const f = a[k][i] / a[i][i];
      for (let j = i; j < n; j++) a[k][j] -= f * a[i][j];
    }
  }
  return det;
}

define('MDETERM', {
  min: 1,
  max: 1,
  impl: ([v], ctx) => {
    const m = numericMatrix(v, ctx);
    if (m instanceof CellError) return m;
    if (m.length !== m[0].length) return Errors.VALUE;
    return Number(determinant(m).toPrecision(15));
  },
});
define('MINVERSE', {
  min: 1,
  max: 1,
  impl: ([v], ctx) => {
    const m = numericMatrix(v, ctx);
    if (m instanceof CellError) return m;
    const n = m.length;
    if (n !== m[0].length) return Errors.VALUE;
    const a = m.map((r, i) => [...r, ...Array.from({ length: n }, (_, j) => (i === j ? 1 : 0))]);
    for (let i = 0; i < n; i++) {
      let p = i;
      for (let k = i + 1; k < n; k++) if (Math.abs(a[k][i]) > Math.abs(a[p][i])) p = k;
      if (Math.abs(a[p][i]) < 1e-15) return Errors.NUM;
      [a[p], a[i]] = [a[i], a[p]];
      const piv = a[i][i];
      for (let j = 0; j < 2 * n; j++) a[i][j] /= piv;
      for (let k = 0; k < n; k++) {
        if (k === i) continue;
        const f = a[k][i];
        for (let j = 0; j < 2 * n; j++) a[k][j] -= f * a[i][j];
      }
    }
    return new ArrayValue(a.map((r) => r.slice(n).map((x) => Number(x.toPrecision(15)))));
  },
});
define('MMULT', {
  min: 2,
  max: 2,
  impl: ([x, y], ctx) => {
    const a = numericMatrix(x, ctx);
    const b = numericMatrix(y, ctx);
    if (a instanceof CellError) return a;
    if (b instanceof CellError) return b;
    if (a[0].length !== b.length) return Errors.VALUE;
    const out: Scalar[][] = a.map((row) => b[0].map((_, j) => row.reduce((s, v, k) => s + v * b[k][j], 0)));
    return new ArrayValue(out);
  },
});
define('MUNIT', {
  min: 1,
  max: 1,
  impl: ([v], ctx) => {
    const n = num(v, ctx);
    if (n instanceof CellError) return n;
    const k = Math.trunc(n);
    if (k < 1) return Errors.VALUE;
    return new ArrayValue(Array.from({ length: k }, (_, i) => Array.from({ length: k }, (_, j) => (i === j ? 1 : 0))));
  },
});

// -----------------------------------------------------------------------------------------------
// SUBTOTAL and AGGREGATE

function isSubtotalCell(ctx: FnContext, area: { sheet: import('../../model/sheet').Sheet }, r: number, c: number): boolean {
  const f = ctx.engine.formulaText(area.sheet, r, c);
  return !!f && /\b(SUBTOTAL|AGGREGATE)\s*\(/i.test(f);
}

interface AggOptions {
  skipHidden: boolean;
  skipErrors: boolean;
  skipNested: boolean;
}

function gatherForAggregate(refs: Value[], ctx: FnContext, opts: AggOptions): { nums: number[]; nonEmpty: number; error?: CellError } {
  const nums: number[] = [];
  let nonEmpty = 0;
  let error: CellError | undefined;
  for (const a of refs) {
    if (a instanceof RefValue) {
      for (const area of a.areas) {
        ctx.engine.forEachValue(area, (v, r, c) => {
          if (opts.skipHidden && area.sheet.isRowHidden(r)) return;
          if (opts.skipNested && isSubtotalCell(ctx, area, r, c)) return;
          if (v instanceof CellError) {
            if (!opts.skipErrors) error ??= v;
            return;
          }
          if (v !== null) nonEmpty++;
          if (typeof v === 'number') nums.push(v);
        });
      }
    } else {
      forEachArg([a], ctx, (v) => {
        if (v instanceof CellError) {
          if (!opts.skipErrors) error ??= v;
          return;
        }
        if (v !== null) nonEmpty++;
        if (typeof v === 'number') nums.push(v);
        else if (typeof v === 'string' || typeof v === 'boolean') {
          const n = toNumber(v);
          if (typeof n === 'number' && typeof v !== 'string') nums.push(n);
        }
      });
    }
  }
  return { nums, nonEmpty, error };
}

function applyAggregate(fn: number, nums: number[], nonEmpty: number, k?: number): number | CellError {
  switch (fn) {
    case 1:
      return mean(nums);
    case 2:
      return nums.length;
    case 3:
      return nonEmpty;
    case 4:
      return nums.length ? Math.max(...nums) : 0;
    case 5:
      return nums.length ? Math.min(...nums) : 0;
    case 6:
      return nums.length ? nums.reduce((a, b) => a * b, 1) : 0;
    case 7:
      return stdev(nums, true);
    case 8:
      return stdev(nums, false);
    case 9:
      return sum(nums);
    case 10:
      return variance(nums, true);
    case 11:
      return variance(nums, false);
    case 12:
      return median(nums);
    case 13:
      return modeSingle(nums);
    case 14:
      return k === undefined ? Errors.VALUE : kth(nums, k, true);
    case 15:
      return k === undefined ? Errors.VALUE : kth(nums, k, false);
    case 16:
      return k === undefined ? Errors.VALUE : percentileInc(nums, k);
    case 17:
      return k === undefined ? Errors.VALUE : quartile(nums, k, false);
    case 18:
      return k === undefined ? Errors.VALUE : percentileExc(nums, k);
    case 19:
      return k === undefined ? Errors.VALUE : quartile(nums, k, true);
    default:
      return Errors.VALUE;
  }
}

define('SUBTOTAL', {
  min: 2,
  max: 255,
  impl: (args, ctx) => {
    const f = num(args[0], ctx);
    if (f instanceof CellError) return f;
    const code = Math.trunc(f);
    const fn = code > 100 ? code - 100 : code;
    if (fn < 1 || fn > 11) return Errors.VALUE;
    const { nums, nonEmpty, error } = gatherForAggregate(args.slice(1), ctx, { skipHidden: code > 100, skipErrors: false, skipNested: true });
    if (error) return error;
    return applyAggregate(fn, nums, nonEmpty);
  },
});
define('AGGREGATE', {
  min: 3,
  max: 255,
  impl: (args, ctx) => {
    const f = num(args[0], ctx);
    const o = args[1] == null ? 0 : num(args[1], ctx);
    const e = err(f, o);
    if (e) return e;
    const fn = Math.trunc(f as number);
    const opt = Math.trunc(o as number);
    const opts: AggOptions = {
      skipNested: [0, 1, 2, 3].includes(opt),
      skipHidden: [1, 3, 5, 7].includes(opt),
      skipErrors: [2, 3, 6, 7].includes(opt),
    };
    if (fn >= 14) {
      const k = num(args[3] ?? null, ctx);
      if (k instanceof CellError) return k;
      const { nums, nonEmpty, error } = gatherForAggregate([args[2]], ctx, opts);
      if (error) return error;
      return applyAggregate(fn, nums, nonEmpty, k);
    }
    const { nums, nonEmpty, error } = gatherForAggregate(args.slice(2), ctx, opts);
    if (error) return error;
    return applyAggregate(fn, nums, nonEmpty);
  },
});

function power(x: number, y: number): number | CellError {
  if (x === 0 && y === 0) return Errors.NUM;
  if (x === 0 && y < 0) return Errors.DIV0;
  const r = Math.pow(x, y);
  if (Number.isNaN(r)) return Errors.NUM;
  return r;
}

export { power };

function ceilTo(x: number, s: number, awayFromZero: boolean): number {
  const q = Number((x / s).toPrecision(15));
  const v = awayFromZero ? -Math.ceil(-q) * s : Math.ceil(q) * s;
  return Number(v.toPrecision(15));
}

function floorTo(x: number, s: number, awayFromZero: boolean): number {
  const q = Number((x / s).toPrecision(15));
  const v = awayFromZero ? -Math.floor(-q) * s : Math.floor(q) * s;
  return Number(v.toPrecision(15));
}
