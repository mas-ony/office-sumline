import { cellName, colIndex, colName, MAX_COLS, MAX_ROWS, quoteSheetName, type Range } from '../../refs';
import { CellError, Errors, type Scalar } from '../../values';
import { ArrayValue, RefValue, compareScalars, toBoolean, toNumber, toText, type Value } from '../runtime';
import { hasWildcards, wildcardToRegExp } from '../criteria';
import { define, err, num, toArray, toScalar, type FnContext } from './helpers';

// -----------------------------------------------------------------------------------------------
// Matching primitives

function sameKind(a: Scalar, b: Scalar): boolean {
  if (typeof a === 'number') return typeof b === 'number';
  if (typeof a === 'string') return typeof b === 'string';
  if (typeof a === 'boolean') return typeof b === 'boolean';
  return false;
}

function equalsLookup(value: Scalar, target: Scalar, re: RegExp | null): boolean {
  if (re) return typeof value === 'string' && re.test(value);
  if (typeof target === 'string' && typeof value === 'string') return target.toLowerCase() === value.toLowerCase();
  return value === target;
}

/** Exact match position (0-based) in a list, with optional wildcards; -1 if none. */
function exactIndex(list: Scalar[], target: Scalar, wildcards: boolean, reverse = false): number {
  const re = wildcards && typeof target === 'string' && hasWildcards(target) ? wildcardToRegExp(target) : null;
  if (reverse) {
    for (let i = list.length - 1; i >= 0; i--) if (equalsLookup(list[i], target, re)) return i;
    return -1;
  }
  for (let i = 0; i < list.length; i++) if (equalsLookup(list[i], target, re)) return i;
  return -1;
}

/** Binary search for the last element <= target among elements of the same kind (ascending data). */
function approxAscending(list: Scalar[], target: Scalar): number {
  let lo = 0;
  let hi = list.length - 1;
  let best = -1;
  while (lo <= hi) {
    const mid = (lo + hi) >> 1;
    let m = mid;
    // Skip values of another type (Excel's binary search effectively ignores them).
    while (m <= hi && !sameKind(list[m], target)) m++;
    if (m > hi) {
      hi = mid - 1;
      continue;
    }
    const cmp = compareScalars(list[m], target);
    if (cmp <= 0) {
      best = m;
      lo = m + 1;
    } else hi = mid - 1;
  }
  return best;
}

/** Binary search for the last element >= target in descending data. */
function approxDescending(list: Scalar[], target: Scalar): number {
  let best = -1;
  for (let i = 0; i < list.length; i++) {
    if (!sameKind(list[i], target)) continue;
    if (compareScalars(list[i], target) >= 0) best = i;
    else break;
  }
  return best;
}

function vector(v: Value, ctx: FnContext): Scalar[] | CellError {
  const a = toArray(v, ctx);
  if (a.rows !== 1 && a.cols !== 1) return Errors.NA;
  return a.rows === 1 ? [...a.data[0]] : a.data.map((r) => r[0]);
}

// -----------------------------------------------------------------------------------------------
// VLOOKUP / HLOOKUP / LOOKUP / MATCH

function tableLookup(vertical: boolean) {
  return (args: Value[], ctx: FnContext): Value => {
    const target = toScalar(args[0], ctx);
    if (target instanceof CellError) return target;
    const idx = num(args[2], ctx);
    if (idx instanceof CellError) return idx;
    const approx = args[3] == null ? true : toBoolean(toScalar(args[3], ctx));
    if (approx instanceof CellError) return approx;
    const table = toArray(args[1], ctx);
    const k = Math.trunc(idx);
    const width = vertical ? table.cols : table.rows;
    if (k < 1) return Errors.VALUE;
    if (k > width) return Errors.REF;
    const keys = vertical ? table.data.map((r) => r[0]) : table.data[0] ?? [];
    let pos: number;
    if (approx) pos = approxAscending(keys, target === null ? 0 : target);
    else pos = exactIndex(keys, target, true);
    if (pos < 0) return Errors.NA;
    const v = vertical ? table.get(pos, k - 1) : table.get(k - 1, pos);
    return v ?? 0;
  };
}
define('VLOOKUP', { min: 3, max: 4, impl: tableLookup(true) });
define('HLOOKUP', { min: 3, max: 4, impl: tableLookup(false) });

define('LOOKUP', {
  min: 2,
  max: 3,
  impl: (args, ctx) => {
    const target = toScalar(args[0], ctx);
    if (target instanceof CellError) return target;
    const arr = toArray(args[1], ctx);
    let keys: Scalar[];
    let results: Scalar[];
    if (args[2] != null) {
      const k = vector(args[1], ctx);
      const r = vector(args[2], ctx);
      if (k instanceof CellError) return k;
      if (r instanceof CellError) return r;
      keys = k;
      results = r;
    } else if (arr.cols > arr.rows) {
      keys = arr.data[0];
      results = arr.data[arr.rows - 1];
    } else {
      keys = arr.data.map((r) => r[0]);
      results = arr.data.map((r) => r[arr.cols - 1]);
    }
    const pos = approxAscending(keys, target);
    if (pos < 0) return Errors.NA;
    return results[pos] ?? Errors.NA;
  },
});

define('MATCH', {
  min: 2,
  max: 3,
  impl: (args, ctx) => {
    const target = toScalar(args[0], ctx);
    if (target instanceof CellError) return target;
    const list = vector(args[1], ctx);
    if (list instanceof CellError) return list;
    const mt = args[2] == null ? 1 : num(args[2], ctx);
    if (mt instanceof CellError) return mt;
    let pos: number;
    if (mt === 0) pos = exactIndex(list, target, true);
    else if (mt > 0) pos = approxAscending(list, target);
    else pos = approxDescending(list, target);
    return pos < 0 ? Errors.NA : pos + 1;
  },
});

function xmatchIndex(list: Scalar[], target: Scalar, matchMode: number, searchMode: number): number {
  const reverse = searchMode === -1;
  if (matchMode === 0 || matchMode === 2) {
    if (searchMode === 2 || searchMode === -2) {
      const sorted = searchMode === 2 ? approxAscending(list, target) : approxDescending(list, target);
      return sorted >= 0 && compareScalars(list[sorted], target) === 0 ? sorted : -1;
    }
    return exactIndex(list, target, matchMode === 2, reverse);
  }
  // Exact or next smaller (-1) / next larger (1).
  const exact = exactIndex(list, target, false, reverse);
  if (exact >= 0) return exact;
  let best = -1;
  const order = reverse ? [...list.keys()].reverse() : [...list.keys()];
  for (const i of order) {
    const v = list[i];
    if (!sameKind(v, target)) continue;
    const cmp = compareScalars(v, target);
    if (matchMode === -1 && cmp < 0 && (best < 0 || compareScalars(v, list[best]) > 0)) best = i;
    if (matchMode === 1 && cmp > 0 && (best < 0 || compareScalars(v, list[best]) < 0)) best = i;
  }
  return best;
}

define('XMATCH', {
  min: 2,
  max: 4,
  impl: (args, ctx) => {
    const list = vector(args[1], ctx);
    if (list instanceof CellError) return list;
    const mm = args[2] == null ? 0 : num(args[2], ctx);
    const sm = args[3] == null ? 1 : num(args[3], ctx);
    const e = err(mm, sm);
    if (e) return e;
    const one = (t: Scalar): Scalar => {
      if (t instanceof CellError) return t;
      const i = xmatchIndex(list, t, mm as number, sm as number);
      return i < 0 ? Errors.NA : i + 1;
    };
    const tv = args[0];
    if (tv instanceof ArrayValue || (tv instanceof RefValue && !tv.isCell())) {
      const ta = toArray(tv, ctx);
      return new ArrayValue(ta.data.map((row) => row.map(one)));
    }
    return one(toScalar(tv, ctx));
  },
});

define('XLOOKUP', {
  min: 3,
  max: 6,
  impl: (args, ctx) => {
    const lookupArr = toArray(args[1], ctx);
    const returnArr = toArray(args[2], ctx);
    const vertical = lookupArr.cols === 1;
    if (lookupArr.rows !== 1 && lookupArr.cols !== 1) return Errors.VALUE;
    const list = vertical ? lookupArr.data.map((r) => r[0]) : lookupArr.data[0];
    if (vertical ? returnArr.rows !== lookupArr.rows : returnArr.cols !== lookupArr.cols) return Errors.VALUE;
    const mm = args[4] == null ? 0 : num(args[4], ctx);
    const sm = args[5] == null ? 1 : num(args[5], ctx);
    const e = err(mm, sm);
    if (e) return e;
    const one = (t: Scalar): Value => {
      if (t instanceof CellError) return t;
      const i = xmatchIndex(list, t, mm as number, sm as number);
      if (i < 0) return args[3] != null ? args[3] : Errors.NA;
      if (vertical) return returnArr.cols === 1 ? (returnArr.get(i, 0) ?? 0) : new ArrayValue([returnArr.data[i].map((v) => v ?? 0)]);
      return returnArr.rows === 1 ? (returnArr.get(0, i) ?? 0) : ArrayValue.column(returnArr.data.map((r) => r[i] ?? 0));
    };
    const tv = args[0];
    if (tv instanceof ArrayValue || (tv instanceof RefValue && !tv.isCell())) {
      const ta = toArray(tv, ctx);
      return new ArrayValue(ta.data.map((row) => row.map((t) => {
        const r = one(t);
        return r instanceof ArrayValue ? r.get(0, 0) : r instanceof RefValue ? toScalar(r, ctx) : r;
      })));
    }
    // Return a reference when the result array is a reference, so XLOOKUP can feed range operators.
    const res = one(toScalar(tv, ctx));
    if (args[2] instanceof RefValue && !(res instanceof CellError)) {
      const t = toScalar(tv, ctx);
      const i = t instanceof CellError ? -1 : xmatchIndex(list, t, mm as number, sm as number);
      if (i >= 0) {
        const { sheet, range } = args[2].single;
        const r: Range = vertical
          ? { r0: range.r0 + i, r1: range.r0 + i, c0: range.c0, c1: range.c1 }
          : { r0: range.r0, r1: range.r1, c0: range.c0 + i, c1: range.c0 + i };
        return new RefValue([{ sheet, range: r }]);
      }
    }
    return res;
  },
});

// -----------------------------------------------------------------------------------------------
// INDEX / OFFSET / INDIRECT / ROW / COLUMN

define('INDEX', {
  min: 2,
  max: 4,
  impl: (args, ctx) => {
    const src = args[0];
    const rowArg = args[1] == null ? 0 : num(args[1], ctx);
    const colArg = args[2] == null ? undefined : num(args[2], ctx);
    const areaArg = args[3] == null ? 1 : num(args[3], ctx);
    const e = err(rowArg, colArg, areaArg);
    if (e) return e;
    let r = Math.trunc(rowArg as number);
    let c = colArg === undefined ? undefined : Math.trunc(colArg as number);
    if (src instanceof RefValue) {
      const area = src.areas[Math.trunc(areaArg as number) - 1];
      if (!area) return Errors.REF;
      const { sheet, range } = area;
      const rows = range.r1 - range.r0 + 1;
      const cols = range.c1 - range.c0 + 1;
      if (c === undefined) {
        if (rows === 1) {
          c = r;
          r = 1;
        } else if (cols === 1) c = 1;
        else c = 0;
      }
      if (r < 0 || c < 0 || r > rows || c > cols) return Errors.REF;
      const out: Range = {
        r0: r === 0 ? range.r0 : range.r0 + r - 1,
        r1: r === 0 ? range.r1 : range.r0 + r - 1,
        c0: c === 0 ? range.c0 : range.c0 + c - 1,
        c1: c === 0 ? range.c1 : range.c0 + c - 1,
      };
      return new RefValue([{ sheet, range: out }]);
    }
    const arr = toArray(src, ctx);
    if (c === undefined) {
      if (arr.rows === 1) {
        c = r;
        r = 1;
      } else c = arr.cols === 1 ? 1 : 0;
    }
    if (r < 0 || c < 0 || r > arr.rows || c > arr.cols) return Errors.REF;
    if (r === 0 && c === 0) return arr;
    if (r === 0) return ArrayValue.column(arr.data.map((row) => row[c! - 1]));
    if (c === 0) return new ArrayValue([arr.data[r - 1]]);
    return arr.get(r - 1, c - 1);
  },
});

define('OFFSET', {
  min: 3,
  max: 5,
  volatile: true,
  impl: (args, ctx) => {
    const base = args[0];
    if (!(base instanceof RefValue)) return Errors.VALUE;
    const rows = num(args[1], ctx);
    const cols = num(args[2], ctx);
    const { sheet, range } = base.single;
    const h = args[3] == null ? range.r1 - range.r0 + 1 : num(args[3], ctx);
    const w = args[4] == null ? range.c1 - range.c0 + 1 : num(args[4], ctx);
    const e = err(rows, cols, h, w);
    if (e) return e;
    const r0 = range.r0 + Math.trunc(rows as number);
    const c0 = range.c0 + Math.trunc(cols as number);
    const hh = Math.trunc(h as number);
    const ww = Math.trunc(w as number);
    if (hh === 0 || ww === 0) return Errors.REF;
    const r1 = r0 + (hh > 0 ? hh - 1 : hh + 1);
    const c1 = c0 + (ww > 0 ? ww - 1 : ww + 1);
    const out = { r0: Math.min(r0, r1), r1: Math.max(r0, r1), c0: Math.min(c0, c1), c1: Math.max(c0, c1) };
    if (out.r0 < 0 || out.c0 < 0 || out.r1 >= MAX_ROWS || out.c1 >= MAX_COLS) return Errors.REF;
    return new RefValue([{ sheet, range: out }]);
  },
});

/** Parses R1C1 references ("R2C3", "R[1]C[-1]", "R2C3:R4C5") relative to the formula cell. */
export function parseR1C1(text: string, ctx: FnContext): { sheet?: string; range: Range } | null {
  let t = text.trim();
  let sheet: string | undefined;
  const bang = t.lastIndexOf('!');
  if (bang > 0) {
    sheet = t.slice(0, bang).replace(/^'(.*)'$/, '$1').replace(/''/g, "'");
    t = t.slice(bang + 1);
  }
  const part = (s: string): { r?: number; c?: number } | null => {
    const m = /^(?:R(\[-?\d+\]|\d+)?)?(?:C(\[-?\d+\]|\d+)?)?$/i.exec(s);
    if (!m || !s) return null;
    const conv = (x: string | undefined, base: number, present: boolean) => {
      if (!present) return undefined;
      if (x === undefined) return base;
      if (x.startsWith('[')) return base + Number(x.slice(1, -1));
      return Number(x) - 1;
    };
    return { r: conv(m[1], ctx.row, /^R/i.test(s)), c: conv(m[2], ctx.col, /C/i.test(s)) };
  };
  const [a, b] = t.split(':');
  const pa = part(a);
  if (!pa) return null;
  const pb = b === undefined ? pa : part(b);
  if (!pb) return null;
  const r0 = pa.r ?? 0;
  const r1 = pb.r ?? (pa.r === undefined ? MAX_ROWS - 1 : r0);
  const c0 = pa.c ?? 0;
  const c1 = pb.c ?? (pa.c === undefined ? MAX_COLS - 1 : c0);
  if (pa.r === undefined && pb.r === undefined) return { sheet, range: { r0: 0, r1: MAX_ROWS - 1, c0: Math.min(c0, c1), c1: Math.max(c0, c1) } };
  return { sheet, range: { r0: Math.min(r0, r1), r1: Math.max(r0, r1), c0: Math.min(c0, c1), c1: Math.max(c0, c1) } };
}

define('INDIRECT', {
  min: 1,
  max: 2,
  volatile: true,
  impl: (args, ctx) => {
    const t = toText(toScalar(args[0], ctx));
    if (t instanceof CellError) return t;
    const a1 = args[1] == null ? true : toBoolean(toScalar(args[1], ctx));
    if (a1 instanceof CellError) return a1;
    if (!a1) {
      const p = parseR1C1(t, ctx);
      if (!p) return Errors.REF;
      const sheet = p.sheet ? ctx.engine.sheetByName(p.sheet) : ctx.sheet;
      if (!sheet) return Errors.REF;
      return new RefValue([{ sheet, range: p.range }]);
    }
    return ctx.engine.resolveReferenceText(t, ctx);
  },
});

function rowsOrCols(isRow: boolean) {
  return (args: Value[], ctx: FnContext): Value => {
    const v = args[0];
    if (v === undefined || v === null) return (isRow ? ctx.row : ctx.col) + 1;
    if (!(v instanceof RefValue)) return Errors.VALUE;
    const { range } = v.single;
    if (isRow) {
      if (range.r0 === range.r1) return range.r0 + 1;
      return ArrayValue.column(Array.from({ length: Math.min(range.r1 - range.r0 + 1, 100000) }, (_, i) => range.r0 + i + 1));
    }
    if (range.c0 === range.c1) return range.c0 + 1;
    return ArrayValue.row(Array.from({ length: range.c1 - range.c0 + 1 }, (_, i) => range.c0 + i + 1));
  };
}
define('ROW', { min: 0, max: 1, impl: rowsOrCols(true) });
define('COLUMN', { min: 0, max: 1, impl: rowsOrCols(false) });
define('ROWS', {
  min: 1,
  max: 1,
  impl: ([v], ctx) => (v instanceof RefValue ? v.rows : toArray(v, ctx).rows),
});
define('COLUMNS', {
  min: 1,
  max: 1,
  impl: ([v], ctx) => (v instanceof RefValue ? v.cols : toArray(v, ctx).cols),
});
define('AREAS', { min: 1, max: 1, impl: ([v]) => (v instanceof RefValue ? v.areas.length : Errors.VALUE) });
define('ADDRESS', {
  min: 2,
  max: 5,
  scalar: [0, 1, 2, 3],
  impl: (args, ctx) => {
    const r = num(args[0], ctx);
    const c = num(args[1], ctx);
    const absNum = args[2] == null ? 1 : num(args[2], ctx);
    const a1 = args[3] == null ? true : toBoolean(toScalar(args[3], ctx));
    const e = err(r, c, absNum, a1);
    if (e) return e;
    const row = Math.trunc(r as number);
    const col = Math.trunc(c as number);
    if (row < 1 || col < 1 || row > MAX_ROWS || col > MAX_COLS) return Errors.VALUE;
    const mode = Math.trunc(absNum as number);
    if (mode < 1 || mode > 4) return Errors.VALUE;
    let ref: string;
    if (a1) ref = cellName(row - 1, col - 1, mode === 1 || mode === 2, mode === 1 || mode === 3);
    else {
      const rr = mode === 1 || mode === 2 ? `R${row}` : `R[${row}]`;
      const cc = mode === 1 || mode === 3 ? `C${col}` : `C[${col}]`;
      ref = rr + cc;
    }
    if (args[4] != null) {
      const s = toText(toScalar(args[4], ctx));
      if (s instanceof CellError) return s;
      return `${quoteSheetName(s)}!${ref}`;
    }
    return ref;
  },
});

// -----------------------------------------------------------------------------------------------
// Dynamic array functions

define('TRANSPOSE', {
  min: 1,
  max: 1,
  impl: ([v], ctx) => {
    const a = toArray(v, ctx);
    return new ArrayValue(Array.from({ length: a.cols }, (_, c) => Array.from({ length: a.rows }, (_, r) => a.get(r, c) ?? 0)));
  },
});

define('FILTER', {
  min: 2,
  max: 3,
  impl: (args, ctx) => {
    const a = toArray(args[0], ctx);
    const inc = toArray(args[1], ctx);
    const byRow = inc.cols === 1 && inc.rows === a.rows;
    const byCol = inc.rows === 1 && inc.cols === a.cols;
    if (!byRow && !byCol) return Errors.VALUE;
    const keep: boolean[] = [];
    const flags = byRow ? inc.data.map((r) => r[0]) : inc.data[0];
    for (const f of flags) {
      if (f instanceof CellError) return f;
      const b = typeof f === 'string' ? Errors.VALUE : toBoolean(f);
      if (b instanceof CellError) return b;
      keep.push(b);
    }
    const out = byRow ? a.data.filter((_, i) => keep[i]) : a.data.map((row) => row.filter((_, i) => keep[i]));
    if (!out.length || !out[0]?.length) return args[2] != null ? args[2] : Errors.CALC;
    return new ArrayValue(out.map((row) => row.map((v) => v ?? 0)));
  },
});

function sortRows(rows: Scalar[][], keys: { index: number; order: number }[]): Scalar[][] {
  return rows
    .map((row, i) => ({ row, i }))
    .sort((x, y) => {
      for (const k of keys) {
        const a = x.row[k.index];
        const b = y.row[k.index];
        // Blanks sort last regardless of order.
        if (a === null && b !== null) return 1;
        if (b === null && a !== null) return -1;
        const cmp = compareScalars(a, b);
        if (cmp) return cmp * k.order;
      }
      return x.i - y.i;
    })
    .map((x) => x.row);
}

function transpose(data: Scalar[][]): Scalar[][] {
  return (data[0] ?? []).map((_, c) => data.map((row) => row[c]));
}

define('SORT', {
  min: 1,
  max: 4,
  impl: (args, ctx) => {
    const a = toArray(args[0], ctx);
    const idx = args[1] == null ? 1 : num(args[1], ctx);
    const ord = args[2] == null ? 1 : num(args[2], ctx);
    const byCol = args[3] == null ? false : toBoolean(toScalar(args[3], ctx));
    const e = err(idx, ord, byCol);
    if (e) return e;
    const data = byCol ? transpose(a.data) : a.data;
    const k = Math.trunc(idx as number);
    if (k < 1 || k > (data[0]?.length ?? 0)) return Errors.VALUE;
    if (ord !== 1 && ord !== -1) return Errors.VALUE;
    const sorted = sortRows(data, [{ index: k - 1, order: ord as number }]);
    return new ArrayValue((byCol ? transpose(sorted) : sorted).map((row) => row.map((v) => v ?? 0)));
  },
});

define('SORTBY', {
  min: 2,
  max: 253,
  impl: (args, ctx) => {
    const a = toArray(args[0], ctx);
    const keys: Scalar[][] = [];
    const orders: number[] = [];
    let byCol: boolean | null = null;
    for (let i = 1; i < args.length; i += 2) {
      const by = toArray(args[i], ctx);
      const col = by.rows === 1 && by.cols === a.cols && a.rows !== 1 ? true : by.cols === 1 && by.rows === a.rows ? false : by.rows === 1 && by.cols === a.cols;
      if (byCol === null) byCol = col;
      if (byCol !== col) return Errors.VALUE;
      keys.push(col ? by.data[0] : by.data.map((r) => r[0]));
      const o = args[i + 1] == null ? 1 : num(args[i + 1], ctx);
      if (o instanceof CellError) return o;
      orders.push(o);
    }
    const data = byCol ? transpose(a.data) : a.data;
    const rows = data.map((row, i) => [...keys.map((k) => k[i]), ...row]);
    const sorted = sortRows(rows, keys.map((_, i) => ({ index: i, order: orders[i] }))).map((r) => r.slice(keys.length));
    return new ArrayValue((byCol ? transpose(sorted) : sorted).map((row) => row.map((v) => v ?? 0)));
  },
});

define('UNIQUE', {
  min: 1,
  max: 3,
  impl: (args, ctx) => {
    const a = toArray(args[0], ctx);
    const byCol = args[1] == null ? false : toBoolean(toScalar(args[1], ctx));
    const once = args[2] == null ? false : toBoolean(toScalar(args[2], ctx));
    const e = err(byCol, once);
    if (e) return e;
    const data = byCol ? transpose(a.data) : a.data;
    const key = (row: Scalar[]) => row.map((v) => (typeof v === 'string' ? `s${v.toLowerCase()}` : `${typeof v}${String(v)}`)).join('\u0000');
    const counts = new Map<string, number>();
    for (const row of data) counts.set(key(row), (counts.get(key(row)) ?? 0) + 1);
    const seen = new Set<string>();
    const out: Scalar[][] = [];
    for (const row of data) {
      const k = key(row);
      if (seen.has(k)) continue;
      seen.add(k);
      if (once && counts.get(k)! > 1) continue;
      out.push(row.map((v) => v ?? 0));
    }
    if (!out.length) return Errors.CALC;
    return new ArrayValue(byCol ? transpose(out) : out);
  },
});

function takeDrop(take: boolean) {
  return (args: Value[], ctx: FnContext): Value => {
    const a = toArray(args[0], ctx);
    const r = args[1] == null ? undefined : num(args[1], ctx);
    const c = args[2] == null ? undefined : num(args[2], ctx);
    const e = err(r, c);
    if (e) return e;
    const slice = <T>(list: T[], n: number | undefined): T[] => {
      if (n === undefined) return list;
      const k = Math.trunc(n);
      if (take) return k >= 0 ? list.slice(0, k) : list.slice(Math.max(0, list.length + k));
      return k >= 0 ? list.slice(k) : list.slice(0, Math.max(0, list.length + k));
    };
    const rows = slice(a.data, r as number | undefined).map((row) => slice(row, c as number | undefined));
    if (!rows.length || !rows[0].length) return Errors.CALC;
    return new ArrayValue(rows);
  };
}
define('TAKE', { min: 2, max: 3, impl: takeDrop(true) });
define('DROP', { min: 2, max: 3, impl: takeDrop(false) });

function chooseLines(rowsMode: boolean) {
  return (args: Value[], ctx: FnContext): Value => {
    const a = toArray(args[0], ctx);
    const data = rowsMode ? a.data : transpose(a.data);
    const picked: Scalar[][] = [];
    for (const arg of args.slice(1)) {
      for (const v of toArray(arg, ctx).data.flat()) {
        const n = toNumber(v);
        if (n instanceof CellError) return n;
        const k = Math.trunc(n);
        const idx = k > 0 ? k - 1 : data.length + k;
        if (k === 0 || idx < 0 || idx >= data.length) return Errors.VALUE;
        picked.push(data[idx]);
      }
    }
    return new ArrayValue(rowsMode ? picked : transpose(picked));
  };
}
define('CHOOSEROWS', { min: 2, max: 255, impl: chooseLines(true) });
define('CHOOSECOLS', { min: 2, max: 255, impl: chooseLines(false) });

define('EXPAND', {
  min: 2,
  max: 4,
  impl: (args, ctx) => {
    const a = toArray(args[0], ctx);
    const r = args[1] == null ? a.rows : num(args[1], ctx);
    const c = args[2] == null ? a.cols : num(args[2], ctx);
    const e = err(r, c);
    if (e) return e;
    const pad = args[3] == null ? Errors.NA : toScalar(args[3], ctx);
    const rows = Math.trunc(r as number);
    const cols = Math.trunc(c as number);
    if (rows < a.rows || cols < a.cols) return Errors.VALUE;
    return new ArrayValue(Array.from({ length: rows }, (_, i) => Array.from({ length: cols }, (_, j) => (i < a.rows && j < a.cols ? a.get(i, j) : pad))));
  },
});

function stack(horizontal: boolean) {
  return (args: Value[], ctx: FnContext): Value => {
    const arrays = args.map((v) => toArray(v, ctx));
    if (horizontal) {
      const rows = Math.max(...arrays.map((a) => a.rows));
      return new ArrayValue(Array.from({ length: rows }, (_, r) => arrays.flatMap((a) => Array.from({ length: a.cols }, (_, c) => (r < a.rows ? (a.get(r, c) ?? 0) : Errors.NA)))));
    }
    const cols = Math.max(...arrays.map((a) => a.cols));
    return new ArrayValue(arrays.flatMap((a) => a.data.map((row) => Array.from({ length: cols }, (_, c) => (c < row.length ? (row[c] ?? 0) : Errors.NA)))));
  };
}
define('HSTACK', { min: 1, max: 254, impl: stack(true) });
define('VSTACK', { min: 1, max: 254, impl: stack(false) });

function toLine(asColumn: boolean) {
  return (args: Value[], ctx: FnContext): Value => {
    const a = toArray(args[0], ctx);
    const ignore = args[1] == null ? 0 : num(args[1], ctx);
    const scanByCol = args[2] == null ? false : toBoolean(toScalar(args[2], ctx));
    const e = err(ignore, scanByCol);
    if (e) return e;
    const data = scanByCol ? transpose(a.data) : a.data;
    let values = data.flat();
    const ig = Math.trunc(ignore as number);
    if (ig === 1 || ig === 3) values = values.filter((v) => v !== null);
    if (ig === 2 || ig === 3) values = values.filter((v) => !(v instanceof CellError));
    if (!values.length) return Errors.CALC;
    const clean = values.map((v) => v ?? 0);
    return asColumn ? ArrayValue.column(clean) : ArrayValue.row(clean);
  };
}
define('TOCOL', { min: 1, max: 3, impl: toLine(true) });
define('TOROW', { min: 1, max: 3, impl: toLine(false) });

function wrapLines(byRows: boolean) {
  return (args: Value[], ctx: FnContext): Value => {
    const a = toArray(args[0], ctx);
    if (a.rows !== 1 && a.cols !== 1) return Errors.VALUE;
    const values = a.data.flat();
    const count = num(args[1], ctx);
    if (count instanceof CellError) return count;
    const k = Math.trunc(count);
    if (k < 1) return Errors.NUM;
    const pad = args[2] == null ? Errors.NA : toScalar(args[2], ctx);
    const lines: Scalar[][] = [];
    for (let i = 0; i < values.length; i += k) {
      const line: Scalar[] = values.slice(i, i + k).map((v) => v ?? 0);
      while (line.length < k) line.push(pad);
      lines.push(line);
    }
    return new ArrayValue(byRows ? lines : transpose(lines));
  };
}
define('WRAPROWS', { min: 2, max: 3, impl: wrapLines(true) });
define('WRAPCOLS', { min: 2, max: 3, impl: wrapLines(false) });

void colIndex;
void colName;
