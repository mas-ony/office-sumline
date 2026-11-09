/** Adapter that exposes @formulajs/formulajs implementations with Excel argument conversion. */
import * as formulajs from '@formulajs/formulajs';
import { CellError, Errors, parseError, type Scalar } from '../../values';
import { ArrayValue, toBoolean, toNumber, toText, type Value } from '../runtime';
import { jsDateFromSerial, serialFromJsDate } from '../dates';
import { collectNumbers, define, toArray, toScalar, type FnContext } from './helpers';

/**
 * Argument kinds: n = number, b = boolean, s = string, d = date (serial -> Date),
 * a = flat list of numbers (text and booleans in references skipped), m = 2D numeric array,
 * v = raw scalar, ? suffix = optional.
 */
type Kind = 'n' | 'b' | 's' | 'd' | 'a' | 'm' | 'v';

function lookup(path: string): ((...args: unknown[]) => unknown) | undefined {
  let obj: unknown = formulajs;
  for (const part of path.split('.')) {
    if (obj === null || obj === undefined) return undefined;
    obj = (obj as Record<string, unknown>)[part];
  }
  return typeof obj === 'function' ? (obj as (...args: unknown[]) => unknown) : undefined;
}

export function fromJs(result: unknown): Value {
  if (result instanceof Error) return parseError(result.message) ?? Errors.VALUE;
  if (result instanceof Date) return serialFromJsDate(result);
  if (typeof result === 'number') return Number.isFinite(result) ? result : Errors.NUM;
  if (typeof result === 'string' || typeof result === 'boolean') return result;
  if (result === null || result === undefined) return null;
  if (Array.isArray(result)) {
    const rows = Array.isArray(result[0]) ? (result as unknown[][]) : [result as unknown[]];
    return new ArrayValue(rows.map((row) => row.map((x) => {
      const v = fromJs(x);
      return v instanceof ArrayValue ? v.get(0, 0) : (v as Scalar);
    })));
  }
  return Errors.VALUE;
}

function convert(v: Value | undefined, kind: Kind, ctx: FnContext): unknown {
  switch (kind) {
    case 'n': {
      const n = toNumber(toScalar(v ?? null, ctx));
      return n;
    }
    case 'b':
      return toBoolean(toScalar(v ?? null, ctx));
    case 's':
      return toText(toScalar(v ?? null, ctx));
    case 'd': {
      const n = toNumber(toScalar(v ?? null, ctx));
      return n instanceof CellError ? n : jsDateFromSerial(n);
    }
    case 'a': {
      const xs = collectNumbers([v ?? null], ctx);
      return xs;
    }
    case 'm': {
      const a = toArray(v ?? null, ctx);
      for (const row of a.data) for (const x of row) if (x instanceof CellError) return x;
      return a.data.map((row) => row.map((x) => (typeof x === 'number' ? x : x === null ? null : typeof x === 'boolean' ? x : x)));
    }
    default:
      return toScalar(v ?? null, ctx);
  }
}

/**
 * Registers `name` backed by formulajs function `path` (dotted, e.g. 'NORM.DIST').
 * `kinds` describes each parameter; a trailing '?' marks an optional one.
 */
export function wrap(name: string, path: string, kinds: string, opts: { volatile?: boolean; dateResult?: boolean } = {}): void {
  const fn = lookup(path);
  if (!fn) return;
  const specs = kinds.split(',').filter(Boolean).map((k) => ({ kind: k.replace(/[?*]/, '') as Kind, optional: k.includes('?'), rest: k.includes('*') }));
  const min = specs.filter((s) => !s.optional && !s.rest).length;
  const max = specs.some((s) => s.rest) ? 255 : specs.length;
  define(name, {
    min,
    max,
    volatile: opts.volatile,
    impl: (args, ctx) => {
      const jsArgs: unknown[] = [];
      for (let i = 0; i < args.length; i++) {
        const spec = specs[Math.min(i, specs.length - 1)];
        if (args[i] === null && spec.optional) {
          jsArgs.push(undefined);
          continue;
        }
        const c = convert(args[i], spec.kind, ctx);
        if (c instanceof CellError) return c;
        jsArgs.push(c);
      }
      while (jsArgs.length && jsArgs[jsArgs.length - 1] === undefined) jsArgs.pop();
      let out: unknown;
      try {
        out = fn(...jsArgs);
      } catch {
        return Errors.VALUE;
      }
      return fromJs(out);
    },
  });
}
