import { getFormatInfo } from 'numfmt';
import { cellName } from '../../refs';
import { CellError, Errors, type Scalar } from '../../values';
import { ArrayValue, RefValue, compareScalars, toBoolean, toNumber, toText, type Thunk, type Value } from '../runtime';
import { define, defineScalar, err, forEachArg, toArray, toScalar, type FnContext } from './helpers';

function isArrayish(v: Value): v is ArrayValue | RefValue {
  return v instanceof ArrayValue || (v instanceof RefValue && !v.isCell());
}

/** Element-wise combination of array arguments with broadcasting (1-wide dimensions repeat). */
export function broadcast(values: Value[], ctx: FnContext, fn: (xs: Scalar[]) => Scalar): Value {
  const arrays = values.map((v) => (isArrayish(v) ? toArray(v, ctx) : null));
  let rows = 1;
  let cols = 1;
  for (const a of arrays) {
    if (!a) continue;
    rows = Math.max(rows, a.rows);
    cols = Math.max(cols, a.cols);
  }
  const scalars = values.map((v, i) => (arrays[i] ? null : toScalar(v, ctx)));
  const out: Scalar[][] = [];
  for (let r = 0; r < rows; r++) {
    const row: Scalar[] = [];
    for (let c = 0; c < cols; c++) {
      const xs = arrays.map((a, i) => {
        if (!a) return scalars[i];
        const rr = a.rows === 1 ? 0 : r;
        const cc = a.cols === 1 ? 0 : c;
        if (rr >= a.rows || cc >= a.cols) return Errors.NA;
        return a.get(rr, cc);
      });
      row.push(fn(xs));
    }
    out.push(row);
  }
  return new ArrayValue(out);
}

function truthy(v: Scalar): boolean | CellError {
  if (typeof v === 'string') {
    const b = toBoolean(v);
    return b instanceof CellError ? Errors.VALUE : b;
  }
  return toBoolean(v);
}

define('TRUE', { min: 0, max: 0, impl: () => true });
define('FALSE', { min: 0, max: 0, impl: () => false });

define('IF', {
  min: 1,
  max: 3,
  lazy: (args: Thunk[], ctx) => {
    const cond = args[0]();
    const pick = (b: boolean): Value => {
      const t = b ? args[1] : args[2];
      if (!t) return b ? true : false;
      const v = t();
      return v === null && !b && args.length < 3 ? false : v;
    };
    if (isArrayish(cond)) {
      const yes = args[1] ? args[1]() : true;
      const no = args[2] ? args[2]() : false;
      return broadcast([cond, yes, no], ctx, ([c, y, n]) => {
        const b = truthy(c);
        if (b instanceof CellError) return b;
        return b ? (y ?? 0) : (n ?? 0);
      });
    }
    const b = truthy(toScalar(cond, ctx));
    if (b instanceof CellError) return b;
    const out = pick(b);
    return out === null ? 0 : out;
  },
});

define('IFS', {
  min: 2,
  max: 254,
  lazy: (args, ctx) => {
    if (args.length % 2) return Errors.VALUE;
    for (let i = 0; i < args.length; i += 2) {
      const b = truthy(toScalar(args[i](), ctx));
      if (b instanceof CellError) return b;
      if (b) {
        const v = args[i + 1]();
        return v === null ? 0 : v;
      }
    }
    return Errors.NA;
  },
});

function errorHandler(test: (e: CellError) => boolean) {
  return (args: Thunk[], ctx: FnContext): Value => {
    const v = args[0]();
    if (isArrayish(v)) {
      const fallback = args[1]();
      return broadcast([v, fallback], ctx, ([x, f]) => (x instanceof CellError && test(x) ? (f ?? 0) : (x ?? 0)));
    }
    const s = toScalar(v, ctx);
    if (s instanceof CellError && test(s)) {
      const f = args[1]();
      return f === null ? 0 : f;
    }
    return v instanceof RefValue ? (s ?? 0) : v === null ? 0 : v;
  };
}
define('IFERROR', { min: 2, max: 2, lazy: errorHandler(() => true) });
define('IFNA', { min: 2, max: 2, lazy: errorHandler((e) => e === Errors.NA) });

define('SWITCH', {
  min: 3,
  max: 254,
  lazy: (args, ctx) => {
    const target = toScalar(args[0](), ctx);
    if (target instanceof CellError) return target;
    const pairsEnd = args.length % 2 === 0 ? args.length - 1 : args.length;
    for (let i = 1; i + 1 < pairsEnd + 1 && i + 1 < args.length; i += 2) {
      const candidate = toScalar(args[i](), ctx);
      if (candidate instanceof CellError) return candidate;
      if (candidate !== null && typeof candidate === typeof target && compareScalars(candidate, target) === 0) {
        const v = args[i + 1]();
        return v === null ? 0 : v;
      }
    }
    if (args.length % 2 === 0) {
      const v = args[args.length - 1]();
      return v === null ? 0 : v;
    }
    return Errors.NA;
  },
});

define('CHOOSE', {
  min: 2,
  max: 255,
  lazy: (args, ctx) => {
    const idx = args[0]();
    if (isArrayish(idx)) {
      const options = args.slice(1).map((t) => t());
      return broadcast([idx], ctx, ([k]) => {
        const n = toNumber(k);
        if (n instanceof CellError) return n;
        const opt = options[Math.trunc(n) - 1];
        if (opt === undefined) return Errors.VALUE;
        return toScalar(opt, ctx);
      });
    }
    const n = toNumber(toScalar(idx, ctx));
    if (n instanceof CellError) return n;
    const k = Math.trunc(n);
    if (k < 1 || k >= args.length) return Errors.VALUE;
    const v = args[k]();
    return v === null ? 0 : v;
  },
});

function logicalReduce(name: string, combine: (acc: boolean, b: boolean) => boolean, init: boolean, xorMode = false): void {
  define(name, {
    min: 1,
    max: 255,
    impl: (args, ctx) => {
      let acc = init;
      let count = 0;
      let error: CellError | undefined;
      let trues = 0;
      forEachArg(args, ctx, (v, fromRef) => {
        if (error) return;
        if (v instanceof CellError) {
          error = v;
          return;
        }
        if (v === null) return;
        if (typeof v === 'string') {
          if (fromRef) return;
          const b = toBoolean(v);
          if (b instanceof CellError) {
            error = b;
            return;
          }
          v = b;
        }
        const b = typeof v === 'boolean' ? v : v !== 0;
        count++;
        if (b) trues++;
        acc = combine(acc, b);
      });
      if (error) return error;
      if (!count) return Errors.VALUE;
      return xorMode ? trues % 2 === 1 : acc;
    },
  });
}
logicalReduce('AND', (a, b) => a && b, true);
logicalReduce('OR', (a, b) => a || b, false);
logicalReduce('XOR', (a, b) => a !== b, false, true);
defineScalar('NOT', 1, 1, ([v]) => {
  const b = truthy(v);
  return b instanceof CellError ? b : !b;
});

// -----------------------------------------------------------------------------------------------
// Information

defineScalar('ISBLANK', 1, 1, ([v]) => v === null);
defineScalar('ISERR', 1, 1, ([v]) => v instanceof CellError && v !== Errors.NA);
defineScalar('ISERROR', 1, 1, ([v]) => v instanceof CellError);
defineScalar('ISNA', 1, 1, ([v]) => v === Errors.NA);
defineScalar('ISLOGICAL', 1, 1, ([v]) => typeof v === 'boolean');
defineScalar('ISNUMBER', 1, 1, ([v]) => typeof v === 'number');
defineScalar('ISTEXT', 1, 1, ([v]) => typeof v === 'string');
defineScalar('ISNONTEXT', 1, 1, ([v]) => typeof v !== 'string');
defineScalar('ISEVEN', 1, 1, ([v]) => {
  const n = toNumber(v);
  return n instanceof CellError ? n : Math.trunc(n) % 2 === 0;
});
defineScalar('ISODD', 1, 1, ([v]) => {
  const n = toNumber(v);
  return n instanceof CellError ? n : Math.abs(Math.trunc(n)) % 2 === 1;
});
defineScalar('N', 1, 1, ([v]) => {
  if (typeof v === 'number') return v;
  if (typeof v === 'boolean') return v ? 1 : 0;
  if (v instanceof CellError) return v;
  return 0;
});
define('NA', { min: 0, max: 0, impl: () => Errors.NA });
define('ISREF', { min: 1, max: 1, impl: ([v]) => v instanceof RefValue });
define('ISFORMULA', {
  min: 1,
  max: 1,
  impl: ([v], ctx) => {
    if (!(v instanceof RefValue)) return Errors.VALUE;
    const { sheet, range } = v.single;
    return ctx.engine.formulaText(sheet, range.r0, range.c0) !== undefined;
  },
});
define('FORMULATEXT', {
  min: 1,
  max: 1,
  impl: ([v], ctx) => {
    if (!(v instanceof RefValue)) return Errors.VALUE;
    const { sheet, range } = v.single;
    const f = ctx.engine.formulaText(sheet, range.r0, range.c0);
    return f === undefined ? Errors.NA : `=${f}`;
  },
});
define('TYPE', {
  min: 1,
  max: 1,
  impl: ([v], ctx) => {
    if (v instanceof ArrayValue || (v instanceof RefValue && !v.isCell())) return 64;
    const s = toScalar(v, ctx);
    if (typeof s === 'number' || s === null) return 1;
    if (typeof s === 'string') return 2;
    if (typeof s === 'boolean') return 4;
    return 16;
  },
});
const ERROR_TYPES: Record<string, number> = {
  '#NULL!': 1, '#DIV/0!': 2, '#VALUE!': 3, '#REF!': 4, '#NAME?': 5, '#NUM!': 6, '#N/A': 7, '#GETTING_DATA': 8, '#SPILL!': 9, '#CALC!': 14,
};
defineScalar('ERROR.TYPE', 1, 1, ([v]) => (v instanceof CellError ? (ERROR_TYPES[v.code] ?? Errors.NA) : Errors.NA));
define('SHEET', {
  min: 0,
  max: 1,
  impl: ([v], ctx) => {
    const wb = ctx.engine.workbook;
    if (v === undefined) return wb.sheets.indexOf(ctx.sheet) + 1;
    if (v instanceof RefValue) return wb.sheets.indexOf(v.single.sheet) + 1;
    const name = toText(toScalar(v, ctx));
    if (name instanceof CellError) return name;
    const sheet = ctx.engine.sheetByName(name);
    return sheet ? wb.sheets.indexOf(sheet) + 1 : Errors.NA;
  },
});
define('SHEETS', {
  min: 0,
  max: 1,
  impl: ([v], ctx) => {
    if (v instanceof RefValue) return new Set(v.areas.map((a) => a.sheet)).size;
    return ctx.engine.workbook.sheets.length;
  },
});
define('CELL', {
  min: 1,
  max: 2,
  impl: ([infoArg, ref], ctx) => {
    const info = toText(toScalar(infoArg, ctx));
    if (info instanceof CellError) return info;
    let sheet = ctx.sheet;
    let r = ctx.row;
    let c = ctx.col;
    if (ref instanceof RefValue) {
      sheet = ref.single.sheet;
      r = ref.single.range.r0;
      c = ref.single.range.c0;
    } else if (ref !== undefined && ref !== null) return Errors.VALUE;
    const value = ctx.engine.cellValue(sheet, r, c);
    const style = ctx.engine.workbook.styles.get(sheet.styleAt(r, c));
    switch (info.toLowerCase()) {
      case 'address':
        return (sheet === ctx.sheet ? '' : `${sheet.name}!`) + cellName(r, c, true, true);
      case 'col':
        return c + 1;
      case 'row':
        return r + 1;
      case 'contents':
        return value ?? 0;
      case 'type':
        return value === null ? 'b' : typeof value === 'string' ? 'l' : 'v';
      case 'format':
        try {
          return getFormatInfo(style.numFmt).code;
        } catch {
          return 'G';
        }
      case 'width':
        return Math.round(sheet.colWidthPx(c) / 7);
      case 'filename':
        return `[${ctx.engine.workbook.props.fileName ?? 'Book1.xlsx'}]${sheet.name}`;
      case 'sheetname':
        return sheet.name;
      case 'prefix': {
        const h = style.align?.horizontal;
        if (typeof value !== 'string') return '';
        return h === 'right' ? '"' : h === 'center' ? '^' : h === 'fill' ? '\\' : "'";
      }
      case 'protect':
        return style.locked === false ? 0 : 1;
      case 'color':
        try {
          return getFormatInfo(style.numFmt).color;
        } catch {
          return 0;
        }
      case 'parentheses':
        try {
          return getFormatInfo(style.numFmt).parentheses;
        } catch {
          return 0;
        }
      default:
        return Errors.VALUE;
    }
  },
});
define('INFO', {
  min: 1,
  max: 1,
  impl: ([v], ctx) => {
    const t = toText(toScalar(v, ctx));
    if (t instanceof CellError) return t;
    switch (t.toLowerCase()) {
      case 'numfile':
        return ctx.engine.workbook.sheets.length;
      case 'recalc':
        return 'Automatic';
      case 'release':
        return '16.0';
      case 'system':
        return 'pcdos';
      case 'osversion':
        return 'Web';
      case 'origin':
        return '$A:$A$1';
      case 'directory':
        return '';
      default:
        return Errors.NA;
    }
  },
});
void err;
