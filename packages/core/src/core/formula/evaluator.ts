import { MAX_COLS, MAX_ROWS, type Range } from '../refs';
import { CellError, Errors, type Scalar } from '../values';
import type { Node } from './parser';
import { ArrayValue, RefValue, compareScalars, toNumber, toText, type Area, type Thunk, type Value } from './runtime';
import { implicitIntersection, registry, toArray, toScalar, type FnContext } from './functions/helpers';
import { power } from './functions/math';
import type { RefParts } from './tokenizer';
import type { Sheet } from '../model/sheet';

export interface EvalContext extends FnContext {
  /** Resolves a defined name to a value (reference, array or scalar). */
  resolveName(name: string, sheetQualifier: string | undefined, ctx: EvalContext): Value;
  /** Spill range of a dynamic array anchored at a cell (A1# operator). */
  spillRange(sheet: Sheet, r: number, c: number): Range | null;
  /** Called for every reference evaluated (dynamic dependency tracking). */
  onReference?(area: Area): void;
}

function refToAreas(ref: RefParts, ctx: EvalContext): Area[] | CellError {
  if (ref.kind === 'error') return Errors.REF;
  if (ref.book !== undefined) return Errors.REF;
  const range: Range = {
    r0: Math.min(ref.r0, ref.r1),
    c0: Math.min(ref.c0, ref.c1),
    r1: Math.max(ref.r0, ref.r1),
    c1: Math.max(ref.c0, ref.c1),
  };
  if (ref.sheet === undefined) return [{ sheet: ctx.sheet, range }];
  const first = ctx.engine.sheetByName(ref.sheet);
  if (!first) return Errors.REF;
  if (ref.sheet2 === undefined) return [{ sheet: first, range }];
  const last = ctx.engine.sheetByName(ref.sheet2);
  if (!last) return Errors.REF;
  const sheets = ctx.engine.workbook.sheets;
  const a = sheets.indexOf(first);
  const b = sheets.indexOf(last);
  return sheets.slice(Math.min(a, b), Math.max(a, b) + 1).map((sheet) => ({ sheet, range }));
}

function isMulti(v: Value): v is ArrayValue | RefValue {
  return v instanceof ArrayValue || (v instanceof RefValue && !v.isCell());
}

/** Applies `fn` element-wise with Excel's broadcasting rules. */
function elementWise(a: Value, b: Value, ctx: EvalContext, fn: (x: Scalar, y: Scalar) => Scalar): Value {
  if (ctx.legacy) {
    if (a instanceof RefValue && !a.isCell()) a = implicitIntersection(a, ctx);
    if (b instanceof RefValue && !b.isCell()) b = implicitIntersection(b, ctx);
  }
  if (!isMulti(a) && !isMulti(b)) return fn(toScalar(a, ctx), toScalar(b, ctx));
  const A = isMulti(a) ? toArray(a, ctx) : null;
  const B = isMulti(b) ? toArray(b, ctx) : null;
  const sa = A ? null : toScalar(a, ctx);
  const sb = B ? null : toScalar(b, ctx);
  const rows = Math.max(A?.rows ?? 1, B?.rows ?? 1);
  const cols = Math.max(A?.cols ?? 1, B?.cols ?? 1);
  const pick = (M: ArrayValue | null, s: Scalar, r: number, c: number): Scalar => {
    if (!M) return s;
    const rr = M.rows === 1 ? 0 : r;
    const cc = M.cols === 1 ? 0 : c;
    if (rr >= M.rows || cc >= M.cols) return Errors.NA;
    return M.get(rr, cc);
  };
  const out: Scalar[][] = [];
  for (let r = 0; r < rows; r++) {
    const row: Scalar[] = [];
    for (let c = 0; c < cols; c++) row.push(fn(pick(A, sa, r, c), pick(B, sb, r, c)));
    out.push(row);
  }
  return new ArrayValue(out);
}

function mapValue(v: Value, ctx: EvalContext, fn: (x: Scalar) => Scalar): Value {
  if (ctx.legacy && v instanceof RefValue && !v.isCell()) v = implicitIntersection(v, ctx);
  if (!isMulti(v)) return fn(toScalar(v, ctx));
  const a = toArray(v, ctx);
  return new ArrayValue(a.data.map((row) => row.map(fn)));
}

function arith(op: string): (x: Scalar, y: Scalar) => Scalar {
  return (x, y) => {
    if (x instanceof CellError) return x;
    if (y instanceof CellError) return y;
    const a = toNumber(x);
    if (a instanceof CellError) return a;
    const b = toNumber(y);
    if (b instanceof CellError) return b;
    let r: number;
    switch (op) {
      case '+':
        r = a + b;
        break;
      case '-':
        r = a - b;
        break;
      case '*':
        r = a * b;
        break;
      case '/':
        if (b === 0) return Errors.DIV0;
        r = a / b;
        break;
      default: {
        const p = power(a, b);
        if (p instanceof CellError) return p;
        r = p;
      }
    }
    if (!Number.isFinite(r)) return Errors.NUM;
    return r;
  };
}

function compareOp(op: string): (x: Scalar, y: Scalar) => Scalar {
  return (x, y) => {
    if (x instanceof CellError) return x;
    if (y instanceof CellError) return y;
    let cmp = compareScalars(x, y);
    if (typeof x === 'number' && typeof y === 'number' && cmp !== 0) {
      // Values equal to 15 significant digits compare equal, like Excel.
      if (Number(x.toPrecision(15)) === Number(y.toPrecision(15))) cmp = 0;
    }
    switch (op) {
      case '=':
        return cmp === 0;
      case '<>':
        return cmp !== 0;
      case '<':
        return cmp < 0;
      case '>':
        return cmp > 0;
      case '<=':
        return cmp <= 0;
      default:
        return cmp >= 0;
    }
  };
}

function concat(x: Scalar, y: Scalar): Scalar {
  const a = toText(x);
  if (a instanceof CellError) return a;
  const b = toText(y);
  if (b instanceof CellError) return b;
  return a + b;
}

function rangeOp(l: Value, r: Value): Value {
  if (l instanceof CellError) return l;
  if (r instanceof CellError) return r;
  if (!(l instanceof RefValue) || !(r instanceof RefValue)) return Errors.VALUE;
  const all = [...l.areas, ...r.areas];
  const sheet = all[0].sheet;
  if (all.some((a) => a.sheet !== sheet)) return Errors.VALUE;
  const range: Range = {
    r0: Math.min(...all.map((a) => a.range.r0)),
    c0: Math.min(...all.map((a) => a.range.c0)),
    r1: Math.max(...all.map((a) => a.range.r1)),
    c1: Math.max(...all.map((a) => a.range.c1)),
  };
  return new RefValue([{ sheet, range }]);
}

function intersectOp(l: Value, r: Value): Value {
  if (l instanceof CellError) return l;
  if (r instanceof CellError) return r;
  if (!(l instanceof RefValue) || !(r instanceof RefValue)) return Errors.VALUE;
  const out: Area[] = [];
  for (const a of l.areas) {
    for (const b of r.areas) {
      if (a.sheet !== b.sheet) continue;
      const range = {
        r0: Math.max(a.range.r0, b.range.r0),
        c0: Math.max(a.range.c0, b.range.c0),
        r1: Math.min(a.range.r1, b.range.r1),
        c1: Math.min(a.range.c1, b.range.c1),
      };
      if (range.r0 <= range.r1 && range.c0 <= range.c1) out.push({ sheet: a.sheet, range });
    }
  }
  return out.length ? new RefValue(out) : Errors.NULL;
}

export function evaluate(node: Node, ctx: EvalContext): Value {
  switch (node.t) {
    case 'num':
      return node.v;
    case 'str':
      return node.v;
    case 'bool':
      return node.v;
    case 'err':
      return node.v;
    case 'missing':
      return null;
    case 'ref': {
      const areas = refToAreas(node.ref, ctx);
      if (areas instanceof CellError) return areas;
      if (ctx.onReference) for (const a of areas) ctx.onReference(a);
      return new RefValue(areas);
    }
    case 'name': {
      if (!node.sheet && ctx.scope) {
        const bound = ctx.scope.get(node.name.toUpperCase());
        if (bound !== undefined) return bound;
      }
      return ctx.resolveName(node.name, node.sheet, ctx);
    }
    case 'struct':
      return ctx.resolveName(node.text, undefined, ctx);
    case 'arr':
      return new ArrayValue(node.rows.map((row) => row.map((n) => toScalar(evaluate(n, ctx), ctx))));
    case 'union': {
      const areas: Area[] = [];
      for (const item of node.items) {
        const v = evaluate(item, ctx);
        if (v instanceof CellError) return v;
        if (!(v instanceof RefValue)) return Errors.VALUE;
        areas.push(...v.areas);
      }
      return new RefValue(areas);
    }
    case 'spill': {
      const v = evaluate(node.arg, ctx);
      if (!(v instanceof RefValue) || !v.isCell()) return Errors.REF;
      const { sheet, range } = v.single;
      const spill = ctx.spillRange(sheet, range.r0, range.c0);
      if (!spill) return Errors.REF;
      const area = { sheet, range: spill };
      ctx.onReference?.(area);
      return new RefValue([area]);
    }
    case 'un': {
      const v = evaluate(node.arg, ctx);
      if (node.op === '@') {
        if (v instanceof RefValue) return v.isCell() ? v : implicitIntersection(v, ctx);
        if (v instanceof ArrayValue) return v.get(0, 0);
        return v;
      }
      if (node.op === '+') return v;
      return mapValue(v, ctx, (x) => {
        if (x instanceof CellError) return x;
        const n = toNumber(x);
        return n instanceof CellError ? n : -n;
      });
    }
    case 'pct': {
      const v = evaluate(node.arg, ctx);
      return mapValue(v, ctx, (x) => {
        if (x instanceof CellError) return x;
        const n = toNumber(x);
        return n instanceof CellError ? n : n / 100;
      });
    }
    case 'bin': {
      const l = evaluate(node.l, ctx);
      if (node.op === ':') return rangeOp(l, evaluate(node.r, ctx));
      if (node.op === ' ') return intersectOp(l, evaluate(node.r, ctx));
      const r = evaluate(node.r, ctx);
      switch (node.op) {
        case '+':
        case '-':
        case '*':
        case '/':
        case '^':
          return elementWise(l, r, ctx, arith(node.op));
        case '&':
          return elementWise(l, r, ctx, concat);
        default:
          return elementWise(l, r, ctx, compareOp(node.op));
      }
    }
    case 'func':
      return callFunction(node.name, node.args, ctx);
  }
}

function callFunction(name: string, argNodes: Node[], ctx: EvalContext): Value {
  if (name === 'LET') return evalLet(argNodes, ctx);
  const def = registry.get(name);
  if (!def) return Errors.NAME;
  if (argNodes.length < def.min || argNodes.length > def.max) return Errors.VALUE;
  if (def.lazy) {
    const thunks: Thunk[] = argNodes.map((n) => () => evaluate(n, ctx));
    return def.lazy(thunks, ctx);
  }
  const args: Value[] = argNodes.map((n) => evaluate(n, ctx));
  const impl = def.impl!;
  if (!def.scalar) return impl(args, ctx);
  const lifted = def.scalar === 'all' ? args.map((_, i) => i) : def.scalar.filter((i) => i < args.length);
  const arrays = new Map<number, ArrayValue>();
  for (const i of lifted) {
    const v = args[i];
    if (v instanceof RefValue) {
      if (v.isCell()) args[i] = toScalar(v, ctx);
      else if (ctx.legacy) args[i] = implicitIntersection(v, ctx);
      else arrays.set(i, toArray(v, ctx));
    } else if (v instanceof ArrayValue) {
      if (v.rows === 1 && v.cols === 1) args[i] = v.get(0, 0);
      else arrays.set(i, v);
    }
  }
  if (!arrays.size) return impl(args, ctx);
  let rows = 1;
  let cols = 1;
  for (const a of arrays.values()) {
    rows = Math.max(rows, a.rows);
    cols = Math.max(cols, a.cols);
  }
  const out: Scalar[][] = [];
  for (let r = 0; r < rows; r++) {
    const row: Scalar[] = [];
    for (let c = 0; c < cols; c++) {
      const callArgs = args.slice();
      let bad = false;
      for (const [i, a] of arrays) {
        const rr = a.rows === 1 ? 0 : r;
        const cc = a.cols === 1 ? 0 : c;
        if (rr >= a.rows || cc >= a.cols) bad = true;
        else callArgs[i] = a.get(rr, cc);
      }
      if (bad) {
        row.push(Errors.NA);
        continue;
      }
      const v = impl(callArgs, ctx);
      row.push(v instanceof ArrayValue || v instanceof RefValue ? toScalar(v, ctx) : v);
    }
    out.push(row);
  }
  return new ArrayValue(out);
}

function evalLet(args: Node[], ctx: EvalContext): Value {
  if (args.length < 3 || args.length % 2 === 0) return Errors.VALUE;
  const scope = new Map(ctx.scope ?? []);
  const inner: EvalContext = { ...ctx, scope };
  for (let i = 0; i < args.length - 1; i += 2) {
    const n = args[i];
    if (n.t !== 'name' || n.sheet) return Errors.VALUE;
    scope.set(n.name.toUpperCase(), evaluate(args[i + 1], inner));
  }
  return evaluate(args[args.length - 1], inner);
}

/** Converts the result of a whole formula to a scalar for a single cell. */
export function resultScalar(v: Value, ctx: FnContext): Scalar {
  if (v instanceof RefValue) {
    if (v.areas.length !== 1) return Errors.VALUE;
    if (v.isCell()) {
      const val = toScalar(v, ctx);
      return val === null ? 0 : val;
    }
    const val = implicitIntersection(v, ctx);
    return val === null ? 0 : val;
  }
  if (v instanceof ArrayValue) {
    const val = v.get(0, 0);
    return val === null ? 0 : val;
  }
  return v;
}

export const LIMITS = { MAX_ROWS, MAX_COLS };
