import type { Sheet } from '../../model/sheet';
import type { Workbook } from '../../model/workbook';
import { CellError, Errors, type Scalar } from '../../values';
import { ArrayValue, RefValue, toBoolean, toNumber, toText, type Area, type Thunk, type Value } from '../runtime';

/** What function implementations can ask the calculation engine. */
export interface EngineApi {
  readonly workbook: Workbook;
  cellValue(sheet: Sheet, r: number, c: number): Scalar;
  /** Visits existing (non-empty or formula) cells of an area, in row-major order. */
  forEachValue(area: Area, fn: (v: Scalar, r: number, c: number) => void): void;
  /** Materialises an area as an array. Whole rows/columns are clipped to the data extent. */
  areaToArray(area: Area): ArrayValue;
  sheetByName(name: string): Sheet | undefined;
  formulaText(sheet: Sheet, r: number, c: number): string | undefined;
  /** Evaluates a formula string in the given context (INDIRECT with names, EVALUATE). */
  resolveReferenceText(text: string, ctx: FnContext): RefValue | CellError;
}

export interface FnContext {
  engine: EngineApi;
  sheet: Sheet;
  row: number;
  col: number;
  /** Legacy (non-array) formula: references in scalar positions use implicit intersection. */
  legacy: boolean;
  /** Variables bound by LET and LAMBDA parameters. */
  scope?: Map<string, Value>;
}

export type FnImpl = (args: Value[], ctx: FnContext) => Value;
export type LazyImpl = (args: Thunk[], ctx: FnContext) => Value;

export interface FnDef {
  min: number;
  max: number;
  impl?: FnImpl;
  lazy?: LazyImpl;
  volatile?: boolean;
  /** Parameters that take a single value; arrays passed there are mapped element by element. */
  scalar?: 'all' | number[];
}

export const registry = new Map<string, FnDef>();

export function define(name: string, def: FnDef): void {
  registry.set(name, def);
}

/** Defines a function whose every parameter is a scalar, so it maps over arrays automatically. */
export function defineScalar(name: string, min: number, max: number, fn: (args: Scalar[], ctx: FnContext) => Value): void {
  registry.set(name, { min, max, scalar: 'all', impl: (args, ctx) => fn(args as Scalar[], ctx) });
}

export function alias(name: string, target: string): void {
  const def = registry.get(target);
  if (def) registry.set(name, def);
}

// -----------------------------------------------------------------------------------------------
// Conversions

/** Value of a single cell reference, implicit intersection of a range, or top-left of an array. */
export function toScalar(v: Value, ctx: FnContext): Scalar {
  if (v instanceof RefValue) {
    if (v.areas.length !== 1) return Errors.VALUE;
    const { sheet, range } = v.areas[0];
    if (range.r0 === range.r1 && range.c0 === range.c1) return ctx.engine.cellValue(sheet, range.r0, range.c0);
    return implicitIntersection(v, ctx);
  }
  if (v instanceof ArrayValue) return v.get(0, 0);
  return v;
}

export function implicitIntersection(v: RefValue, ctx: FnContext): Scalar {
  const { sheet, range } = v.areas[0];
  if (range.r0 === range.r1 && ctx.col >= range.c0 && ctx.col <= range.c1) return ctx.engine.cellValue(sheet, range.r0, ctx.col);
  if (range.c0 === range.c1 && ctx.row >= range.r0 && ctx.row <= range.r1) return ctx.engine.cellValue(sheet, ctx.row, range.c0);
  if (ctx.row >= range.r0 && ctx.row <= range.r1 && ctx.col >= range.c0 && ctx.col <= range.c1 && range.r0 === range.r1) {
    return ctx.engine.cellValue(sheet, ctx.row, ctx.col);
  }
  return Errors.VALUE;
}

export function toArray(v: Value, ctx: FnContext): ArrayValue {
  if (v instanceof ArrayValue) return v;
  if (v instanceof RefValue) {
    if (v.areas.length !== 1) return new ArrayValue([[Errors.VALUE]]);
    return ctx.engine.areaToArray(v.areas[0]);
  }
  return new ArrayValue([[v]]);
}

export function num(v: Value, ctx: FnContext): number | CellError {
  return toNumber(toScalar(v, ctx));
}

export function int(v: Value, ctx: FnContext): number | CellError {
  const n = num(v, ctx);
  return n instanceof CellError ? n : Math.trunc(n);
}

export function bool(v: Value, ctx: FnContext): boolean | CellError {
  return toBoolean(toScalar(v, ctx));
}

export function text(v: Value, ctx: FnContext): string | CellError {
  return toText(toScalar(v, ctx));
}

/** Optional argument helpers: a missing argument evaluates to null. */
export function optNum(v: Value | undefined, dflt: number, ctx: FnContext): number | CellError {
  if (v === undefined || v === null) return dflt;
  return num(v, ctx);
}

export function optBool(v: Value | undefined, dflt: boolean, ctx: FnContext): boolean | CellError {
  if (v === undefined || v === null) return dflt;
  return bool(v, ctx);
}

export function optText(v: Value | undefined, dflt: string, ctx: FnContext): string | CellError {
  if (v === undefined || v === null) return dflt;
  return text(v, ctx);
}

/** Returns the first CellError among already converted values. */
export function err(...vals: unknown[]): CellError | undefined {
  for (const v of vals) if (v instanceof CellError) return v;
  return undefined;
}

export interface FlatItem {
  v: Scalar;
  /** True when the value came from a reference or array (not typed as an argument). */
  fromRef: boolean;
}

/** Visits every scalar of the arguments: references and arrays are expanded. */
export function forEachArg(args: Value[], ctx: FnContext, fn: (v: Scalar, fromRef: boolean) => void): void {
  for (const a of args) {
    if (a instanceof RefValue) {
      for (const area of a.areas) ctx.engine.forEachValue(area, (v) => fn(v, true));
    } else if (a instanceof ArrayValue) {
      for (const row of a.data) for (const v of row) fn(v, true);
    } else fn(a, false);
  }
}

/**
 * Collects numbers following the SUM/AVERAGE rules: typed arguments are coerced (text that is
 * not a number gives #VALUE!), while text and booleans inside references are skipped.
 */
export function collectNumbers(args: Value[], ctx: FnContext, opts: { countBoolInRefs?: boolean; textAsZero?: boolean } = {}): number[] | CellError {
  const out: number[] = [];
  let error: CellError | undefined;
  forEachArg(args, ctx, (v, fromRef) => {
    if (error) return;
    if (v instanceof CellError) {
      error = v;
      return;
    }
    if (typeof v === 'number') {
      out.push(v);
      return;
    }
    if (fromRef) {
      if (typeof v === 'boolean' && opts.countBoolInRefs) out.push(v ? 1 : 0);
      else if (typeof v === 'string' && opts.textAsZero) out.push(0);
      return;
    }
    if (v === null) return;
    const n = toNumber(v);
    if (n instanceof CellError) error = n;
    else out.push(n);
  });
  return error ?? out;
}

/** All values of the arguments as a flat list (references expanded, blanks dropped). */
export function flatValues(args: Value[], ctx: FnContext, keepBlanks = false): Scalar[] {
  const out: Scalar[] = [];
  forEachArg(args, ctx, (v) => {
    if (v !== null || keepBlanks) out.push(v);
  });
  return out;
}

/** All cells of a value as a flat list, including blanks (for paired ranges such as SUMPRODUCT). */
export function cellsOf(v: Value, ctx: FnContext): Scalar[] {
  const arr = toArray(v, ctx);
  const out: Scalar[] = [];
  for (const row of arr.data) for (const x of row) out.push(x);
  return out;
}

export function matrixMap(a: ArrayValue, fn: (v: Scalar, r: number, c: number) => Scalar): ArrayValue {
  return new ArrayValue(a.data.map((row, r) => row.map((v, c) => fn(v, r, c))));
}

export function isMissing(v: Value | undefined): boolean {
  return v === undefined;
}

export function evalThunk(t: Thunk | undefined): Value {
  return t ? t() : null;
}

export function argCountError(): CellError {
  return Errors.VALUE;
}

export function numOrError(n: number): number | CellError {
  return Number.isFinite(n) ? n : Errors.NUM;
}
