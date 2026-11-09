import { CellError, Errors, type Scalar } from '../../values';
import { ArrayValue, RefValue, toNumber, type Value } from '../runtime';
import { makeCriteria } from '../criteria';
import { alias, collectNumbers, define, err, forEachArg, num, toArray, toScalar, type FnContext } from './helpers';
import { conditionalCells } from './math';
import { kth, mean, median, modeSingle, percentileExc, percentileInc, quartile, stdev, sum, variance } from './numeric';
import { wrap } from './fjs';

function numbersOf(args: Value[], ctx: FnContext, a = false): number[] | CellError {
  return collectNumbers(args, ctx, a ? { countBoolInRefs: true, textAsZero: true } : {});
}

function reducer(name: string, f: (xs: number[]) => number | CellError, aVariant = false, emptyValue?: number): void {
  define(name, {
    min: 1,
    max: 255,
    impl: (args, ctx) => {
      const xs = numbersOf(args, ctx, aVariant);
      if (xs instanceof CellError) return xs;
      if (!xs.length && emptyValue !== undefined) return emptyValue;
      return f(xs);
    },
  });
}

reducer('AVERAGE', mean);
reducer('AVERAGEA', mean, true);
reducer('MAX', (xs) => Math.max(...xs), false, 0);
reducer('MAXA', (xs) => Math.max(...xs), true, 0);
reducer('MIN', (xs) => Math.min(...xs), false, 0);
reducer('MINA', (xs) => Math.min(...xs), true, 0);
reducer('MEDIAN', median);
reducer('MODE.SNGL', modeSingle);
alias('MODE', 'MODE.SNGL');
define('MODE.MULT', {
  min: 1,
  max: 255,
  impl: (args, ctx) => {
    const xs = numbersOf(args, ctx);
    if (xs instanceof CellError) return xs;
    const counts = new Map<number, number>();
    for (const x of xs) counts.set(x, (counts.get(x) ?? 0) + 1);
    const best = Math.max(0, ...counts.values());
    if (best < 2) return Errors.NA;
    const modes = [...counts.entries()].filter(([, c]) => c === best).map(([v]) => v);
    return ArrayValue.column(modes);
  },
});
reducer('STDEV.S', (xs) => stdev(xs, true));
alias('STDEV', 'STDEV.S');
reducer('STDEV.P', (xs) => stdev(xs, false));
alias('STDEVP', 'STDEV.P');
reducer('STDEVA', (xs) => stdev(xs, true), true);
reducer('STDEVPA', (xs) => stdev(xs, false), true);
reducer('VAR.S', (xs) => variance(xs, true));
alias('VAR', 'VAR.S');
reducer('VAR.P', (xs) => variance(xs, false));
alias('VARP', 'VAR.P');
reducer('VARA', (xs) => variance(xs, true), true);
reducer('VARPA', (xs) => variance(xs, false), true);
reducer('AVEDEV', (xs) => {
  if (!xs.length) return Errors.NUM;
  const m = sum(xs) / xs.length;
  return sum(xs.map((x) => Math.abs(x - m))) / xs.length;
});
reducer('DEVSQ', (xs) => {
  if (!xs.length) return Errors.NUM;
  const m = sum(xs) / xs.length;
  return sum(xs.map((x) => (x - m) * (x - m)));
});
reducer('GEOMEAN', (xs) => {
  if (!xs.length || xs.some((x) => x <= 0)) return Errors.NUM;
  return Math.exp(sum(xs.map(Math.log)) / xs.length);
});
reducer('HARMEAN', (xs) => {
  if (!xs.length || xs.some((x) => x <= 0)) return Errors.NUM;
  return xs.length / sum(xs.map((x) => 1 / x));
});
reducer('KURT', (xs) => {
  const n = xs.length;
  if (n < 4) return Errors.DIV0;
  const m = sum(xs) / n;
  const s = stdev(xs, true);
  if (typeof s !== 'number' || s === 0) return Errors.DIV0;
  const k = sum(xs.map((x) => Math.pow((x - m) / s, 4)));
  return ((n * (n + 1)) / ((n - 1) * (n - 2) * (n - 3))) * k - (3 * (n - 1) * (n - 1)) / ((n - 2) * (n - 3));
});
reducer('SKEW', (xs) => {
  const n = xs.length;
  if (n < 3) return Errors.DIV0;
  const m = sum(xs) / n;
  const s = stdev(xs, true);
  if (typeof s !== 'number' || s === 0) return Errors.DIV0;
  return (n / ((n - 1) * (n - 2))) * sum(xs.map((x) => Math.pow((x - m) / s, 3)));
});
reducer('SKEW.P', (xs) => {
  const n = xs.length;
  if (n < 3) return Errors.DIV0;
  const m = sum(xs) / n;
  const s = stdev(xs, false);
  if (typeof s !== 'number' || s === 0) return Errors.DIV0;
  return sum(xs.map((x) => Math.pow((x - m) / s, 3))) / n;
});

// -----------------------------------------------------------------------------------------------
// Counting

define('COUNT', {
  min: 1,
  max: 255,
  impl: (args, ctx) => {
    let n = 0;
    forEachArg(args, ctx, (v, fromRef) => {
      if (typeof v === 'number') n++;
      else if (!fromRef && (typeof v === 'boolean' || (typeof v === 'string' && typeof toNumber(v) === 'number'))) n++;
    });
    return n;
  },
});
define('COUNTA', {
  min: 1,
  max: 255,
  impl: (args, ctx) => {
    let n = 0;
    forEachArg(args, ctx, (v, fromRef) => {
      if (v !== null || !fromRef) n++;
    });
    return n;
  },
});
define('COUNTBLANK', {
  min: 1,
  max: 1,
  impl: ([v], ctx) => {
    if (v instanceof RefValue) {
      let filled = 0;
      let total = 0;
      for (const area of v.areas) {
        total += (area.range.r1 - area.range.r0 + 1) * (area.range.c1 - area.range.c0 + 1);
        ctx.engine.forEachValue(area, (x) => {
          if (x !== null && x !== '') filled++;
        });
      }
      return total - filled;
    }
    const arr = toArray(v, ctx);
    let n = 0;
    for (const row of arr.data) for (const x of row) if (x === null || x === '') n++;
    return n;
  },
});
define('COUNTIF', {
  min: 2,
  max: 2,
  impl: ([range, crit], ctx) => {
    const match = makeCriteria(toScalar(crit, ctx));
    if (range instanceof RefValue) {
      let n = 0;
      const isBlankCrit = match(null);
      let filled = 0;
      let total = 0;
      for (const area of range.areas) {
        total += (area.range.r1 - area.range.r0 + 1) * (area.range.c1 - area.range.c0 + 1);
        ctx.engine.forEachValue(area, (x) => {
          if (x !== null) filled++;
          if (x !== null && match(x)) n++;
        });
      }
      if (isBlankCrit) n += total - filled;
      return n;
    }
    const arr = toArray(range, ctx);
    let n = 0;
    for (const row of arr.data) for (const x of row) if (match(x)) n++;
    return n;
  },
});
define('COUNTIFS', {
  min: 2,
  max: 254,
  impl: (args, ctx) => {
    if (args.length % 2) return Errors.VALUE;
    const pairs: [Value, Value][] = [];
    for (let i = 0; i < args.length; i += 2) pairs.push([args[i], args[i + 1]]);
    let n = 0;
    const e = conditionalCells(ctx, pairs, () => {
      n++;
    });
    return e ?? n;
  },
});

function conditionalNumbers(args: Value[], ctx: FnContext, valuesFirst: boolean): number[] | CellError {
  // AVERAGEIF(range, criteria, [average_range]) vs MAXIFS(max_range, range1, crit1, ...)
  const nums: number[] = [];
  let error: CellError | undefined;
  if (!valuesFirst) {
    const values = toArray(args[2] ?? args[0], ctx);
    const e = conditionalCells(ctx, [[args[0], args[1]]], ([r, c]) => {
      const v = values.get(r, c);
      if (v instanceof CellError) error ??= v;
      else if (typeof v === 'number') nums.push(v);
    });
    if (e) return e;
  } else {
    if ((args.length - 1) % 2) return Errors.VALUE;
    const values = toArray(args[0], ctx);
    const pairs: [Value, Value][] = [];
    for (let i = 1; i < args.length; i += 2) pairs.push([args[i], args[i + 1]]);
    const first = toArray(args[1], ctx);
    if (first.rows !== values.rows || first.cols !== values.cols) return Errors.VALUE;
    const e = conditionalCells(ctx, pairs, ([r, c]) => {
      const v = values.get(r, c);
      if (v instanceof CellError) error ??= v;
      else if (typeof v === 'number') nums.push(v);
    });
    if (e) return e;
  }
  return error ?? nums;
}

define('AVERAGEIF', {
  min: 2,
  max: 3,
  impl: (args, ctx) => {
    const xs = conditionalNumbers(args, ctx, false);
    if (xs instanceof CellError) return xs;
    return xs.length ? sum(xs) / xs.length : Errors.DIV0;
  },
});
define('AVERAGEIFS', {
  min: 3,
  max: 255,
  impl: (args, ctx) => {
    const xs = conditionalNumbers(args, ctx, true);
    if (xs instanceof CellError) return xs;
    return xs.length ? sum(xs) / xs.length : Errors.DIV0;
  },
});
define('MAXIFS', {
  min: 3,
  max: 255,
  impl: (args, ctx) => {
    const xs = conditionalNumbers(args, ctx, true);
    if (xs instanceof CellError) return xs;
    return xs.length ? Math.max(...xs) : 0;
  },
});
define('MINIFS', {
  min: 3,
  max: 255,
  impl: (args, ctx) => {
    const xs = conditionalNumbers(args, ctx, true);
    if (xs instanceof CellError) return xs;
    return xs.length ? Math.min(...xs) : 0;
  },
});

// -----------------------------------------------------------------------------------------------
// Order statistics

function kthFn(name: string, largest: boolean): void {
  define(name, {
    min: 2,
    max: 2,
    impl: ([arr, k], ctx) => {
      const xs = numbersOf([arr], ctx);
      if (xs instanceof CellError) return xs;
      if (k instanceof ArrayValue || (k instanceof RefValue && !k.isCell())) {
        const ks = toArray(k, ctx);
        return new ArrayValue(ks.data.map((row) => row.map((kv) => {
          const n = toNumber(kv);
          return n instanceof CellError ? n : kth(xs, n, largest);
        })));
      }
      const n = num(k, ctx);
      if (n instanceof CellError) return n;
      return kth(xs, n, largest);
    },
  });
}
kthFn('LARGE', true);
kthFn('SMALL', false);

function percentileFn(name: string, f: (xs: number[], p: number) => number | CellError): void {
  define(name, {
    min: 2,
    max: 2,
    impl: ([arr, p], ctx) => {
      const xs = numbersOf([arr], ctx);
      const k = num(p, ctx);
      const e = err(xs, k);
      if (e) return e;
      return f(xs as number[], k as number);
    },
  });
}
percentileFn('PERCENTILE.INC', percentileInc);
alias('PERCENTILE', 'PERCENTILE.INC');
percentileFn('PERCENTILE.EXC', percentileExc);
percentileFn('QUARTILE.INC', (xs, q) => quartile(xs, q, false));
alias('QUARTILE', 'QUARTILE.INC');
percentileFn('QUARTILE.EXC', (xs, q) => quartile(xs, q, true));

function rankFn(name: string, avg: boolean): void {
  define(name, {
    min: 2,
    max: 3,
    impl: ([x, ref, order], ctx) => {
      const n = num(x, ctx);
      const o = order == null ? 0 : num(order, ctx);
      const xs = numbersOf([ref], ctx);
      const e = err(n, o, xs);
      if (e) return e;
      const list = xs as number[];
      const value = n as number;
      if (!list.includes(value)) return Errors.NA;
      const asc = (o as number) !== 0;
      const better = list.filter((v) => (asc ? v < value : v > value)).length;
      if (!avg) return better + 1;
      const ties = list.filter((v) => v === value).length;
      return better + (ties + 1) / 2;
    },
  });
}
rankFn('RANK.EQ', false);
alias('RANK', 'RANK.EQ');
rankFn('RANK.AVG', true);

function percentRank(xs: number[], x: number, sig: number, exclusive: boolean): number | CellError {
  if (!xs.length) return Errors.NUM;
  const s = [...xs].sort((a, b) => a - b);
  if (x < s[0] || x > s[s.length - 1]) return Errors.NA;
  const n = s.length;
  let pos: number;
  const below = s.filter((v) => v < x).length;
  if (s.includes(x)) pos = below;
  else {
    const lo = s[below - 1];
    const hi = s[below];
    pos = below - 1 + (x - lo) / (hi - lo);
  }
  const r = exclusive ? (pos + 1) / (n + 1) : n === 1 ? 1 : pos / (n - 1);
  const f = Math.pow(10, Math.trunc(sig));
  return Math.floor(Number((r * f).toPrecision(15))) / f;
}

function percentRankFn(name: string, exclusive: boolean): void {
  define(name, {
    min: 2,
    max: 3,
    impl: ([arr, x, sig], ctx) => {
      const xs = numbersOf([arr], ctx);
      const v = num(x, ctx);
      const s = sig == null ? 3 : num(sig, ctx);
      const e = err(xs, v, s);
      if (e) return e;
      if ((s as number) < 1) return Errors.NUM;
      return percentRank(xs as number[], v as number, s as number, exclusive);
    },
  });
}
percentRankFn('PERCENTRANK.INC', false);
alias('PERCENTRANK', 'PERCENTRANK.INC');
percentRankFn('PERCENTRANK.EXC', true);

define('TRIMMEAN', {
  min: 2,
  max: 2,
  impl: ([arr, p], ctx) => {
    const xs = numbersOf([arr], ctx);
    const pct = num(p, ctx);
    const e = err(xs, pct);
    if (e) return e;
    const list = [...(xs as number[])].sort((a, b) => a - b);
    if ((pct as number) < 0 || (pct as number) >= 1 || !list.length) return Errors.NUM;
    const k = Math.floor((list.length * (pct as number)) / 2);
    const kept = list.slice(k, list.length - k);
    return sum(kept) / kept.length;
  },
});

define('FREQUENCY', {
  min: 2,
  max: 2,
  impl: ([data, bins], ctx) => {
    const xs = numbersOf([data], ctx);
    const bs = numbersOf([bins], ctx);
    const e = err(xs, bs);
    if (e) return e;
    const sorted = [...(bs as number[])].sort((a, b) => a - b);
    const counts = new Array(sorted.length + 1).fill(0);
    for (const x of xs as number[]) {
      let i = 0;
      while (i < sorted.length && x > sorted[i]) i++;
      counts[i]++;
    }
    // Results follow the order of the bins as given.
    const order = (bs as number[]).map((b) => sorted.indexOf(b));
    const out: Scalar[] = order.map((i) => counts[i]);
    out.push(counts[sorted.length]);
    return ArrayValue.column(out);
  },
});

// -----------------------------------------------------------------------------------------------
// Paired data

function pairs(ys: Value, xs: Value, ctx: FnContext): [number[], number[]] | CellError {
  const a = toArray(ys, ctx);
  const b = toArray(xs, ctx);
  if (a.rows * a.cols !== b.rows * b.cols) return Errors.NA;
  const outY: number[] = [];
  const outX: number[] = [];
  const fa = a.data.flat();
  const fb = b.data.flat();
  for (let i = 0; i < fa.length; i++) {
    const y = fa[i];
    const x = fb[i];
    if (y instanceof CellError) return y;
    if (x instanceof CellError) return x;
    if (typeof y === 'number' && typeof x === 'number') {
      outY.push(y);
      outX.push(x);
    }
  }
  return [outY, outX];
}

function regression(ys: number[], xs: number[]): { slope: number; intercept: number; r2: number; n: number; ssr: number } | CellError {
  const n = ys.length;
  if (n < 2) return Errors.DIV0;
  const mx = sum(xs) / n;
  const my = sum(ys) / n;
  let sxy = 0;
  let sxx = 0;
  let syy = 0;
  for (let i = 0; i < n; i++) {
    sxy += (xs[i] - mx) * (ys[i] - my);
    sxx += (xs[i] - mx) * (xs[i] - mx);
    syy += (ys[i] - my) * (ys[i] - my);
  }
  if (sxx === 0) return Errors.DIV0;
  const slope = sxy / sxx;
  const intercept = my - slope * mx;
  const r2 = syy === 0 ? 1 : (sxy * sxy) / (sxx * syy);
  const ssr = syy - slope * sxy;
  return { slope, intercept, r2, n, ssr };
}

function pairFn(name: string, f: (ys: number[], xs: number[]) => number | CellError): void {
  define(name, {
    min: 2,
    max: 2,
    impl: ([a, b], ctx) => {
      const p = pairs(a, b, ctx);
      if (p instanceof CellError) return p;
      return f(p[0], p[1]);
    },
  });
}

pairFn('SLOPE', (ys, xs) => {
  const r = regression(ys, xs);
  return r instanceof CellError ? r : r.slope;
});
pairFn('INTERCEPT', (ys, xs) => {
  const r = regression(ys, xs);
  return r instanceof CellError ? r : r.intercept;
});
pairFn('RSQ', (ys, xs) => {
  const r = regression(ys, xs);
  return r instanceof CellError ? r : r.r2;
});
pairFn('STEYX', (ys, xs) => {
  const r = regression(ys, xs);
  if (r instanceof CellError) return r;
  if (r.n < 3) return Errors.DIV0;
  return Math.sqrt(r.ssr / (r.n - 2));
});
pairFn('CORREL', (a, b) => {
  const n = a.length;
  if (n < 2) return Errors.DIV0;
  const ma = sum(a) / n;
  const mb = sum(b) / n;
  let sab = 0;
  let saa = 0;
  let sbb = 0;
  for (let i = 0; i < n; i++) {
    sab += (a[i] - ma) * (b[i] - mb);
    saa += (a[i] - ma) ** 2;
    sbb += (b[i] - mb) ** 2;
  }
  if (saa === 0 || sbb === 0) return Errors.DIV0;
  return sab / Math.sqrt(saa * sbb);
});
alias('PEARSON', 'CORREL');
pairFn('COVARIANCE.P', (a, b) => {
  const n = a.length;
  if (!n) return Errors.DIV0;
  const ma = sum(a) / n;
  const mb = sum(b) / n;
  return sum(a.map((x, i) => (x - ma) * (b[i] - mb))) / n;
});
alias('COVAR', 'COVARIANCE.P');
pairFn('COVARIANCE.S', (a, b) => {
  const n = a.length;
  if (n < 2) return Errors.DIV0;
  const ma = sum(a) / n;
  const mb = sum(b) / n;
  return sum(a.map((x, i) => (x - ma) * (b[i] - mb))) / (n - 1);
});

define('FORECAST.LINEAR', {
  min: 3,
  max: 3,
  impl: ([x, ys, xs], ctx) => {
    const v = num(x, ctx);
    if (v instanceof CellError) return v;
    const p = pairs(ys, xs, ctx);
    if (p instanceof CellError) return p;
    const r = regression(p[0], p[1]);
    return r instanceof CellError ? r : r.intercept + r.slope * v;
  },
});
alias('FORECAST', 'FORECAST.LINEAR');

function trendLike(name: string, exponential: boolean): void {
  define(name, {
    min: 1,
    max: 4,
    impl: ([knownY, knownX, newX, constArg], ctx) => {
      const ya = toArray(knownY, ctx);
      const yflat = ya.data.flat();
      const xflat = knownX == null ? yflat.map((_, i) => i + 1) : toArray(knownX, ctx).data.flat();
      if (xflat.length !== yflat.length) return Errors.REF;
      const ys: number[] = [];
      const xs: number[] = [];
      for (let i = 0; i < yflat.length; i++) {
        const y = toNumber(yflat[i]);
        const x = toNumber(xflat[i] as Scalar);
        if (y instanceof CellError || x instanceof CellError) return Errors.VALUE;
        if (exponential && y <= 0) return Errors.NUM;
        ys.push(exponential ? Math.log(y) : y);
        xs.push(x);
      }
      const useConst = constArg == null ? true : toScalar(constArg, ctx) !== false && toNumber(toScalar(constArg, ctx)) !== 0;
      let slope: number;
      let intercept: number;
      if (useConst) {
        const r = regression(ys, xs);
        if (r instanceof CellError) return r;
        slope = r.slope;
        intercept = r.intercept;
      } else {
        const sxy = sum(xs.map((x, i) => x * ys[i]));
        const sxx = sum(xs.map((x) => x * x));
        if (sxx === 0) return Errors.DIV0;
        slope = sxy / sxx;
        intercept = 0;
      }
      const targets = newX == null ? new ArrayValue(xs.map((x) => [x])) : toArray(newX, ctx);
      return new ArrayValue(targets.data.map((row) => row.map((t) => {
        const x = toNumber(t);
        if (x instanceof CellError) return x;
        const y = intercept + slope * x;
        return exponential ? Math.exp(y) : y;
      })));
    },
  });
}
trendLike('TREND', false);
trendLike('GROWTH', true);

function linestLike(name: string, exponential: boolean): void {
  define(name, {
    min: 1,
    max: 4,
    impl: ([knownY, knownX, constArg, statsArg], ctx) => {
      const yflat = toArray(knownY, ctx).data.flat();
      const xflat = knownX == null ? yflat.map((_, i) => i + 1) : toArray(knownX, ctx).data.flat();
      if (xflat.length !== yflat.length) return Errors.REF;
      const ys: number[] = [];
      const xs: number[] = [];
      for (let i = 0; i < yflat.length; i++) {
        const y = toNumber(yflat[i]);
        const x = toNumber(xflat[i] as Scalar);
        if (y instanceof CellError || x instanceof CellError) return Errors.VALUE;
        if (exponential && y <= 0) return Errors.NUM;
        ys.push(exponential ? Math.log(y) : y);
        xs.push(x);
      }
      const useConst = constArg == null ? true : toScalar(constArg, ctx) !== false && toNumber(toScalar(constArg, ctx)) !== 0;
      const stats = statsArg != null && (toScalar(statsArg, ctx) === true || toNumber(toScalar(statsArg, ctx)) === 1);
      const n = ys.length;
      let slope: number;
      let intercept: number;
      if (useConst) {
        const r = regression(ys, xs);
        if (r instanceof CellError) return r;
        slope = r.slope;
        intercept = r.intercept;
      } else {
        const sxx = sum(xs.map((x) => x * x));
        if (sxx === 0) return Errors.DIV0;
        slope = sum(xs.map((x, i) => x * ys[i])) / sxx;
        intercept = 0;
      }
      const head = exponential ? [Math.exp(slope), Math.exp(intercept)] : [slope, intercept];
      if (!stats) return new ArrayValue([head]);
      const fitted = xs.map((x) => intercept + slope * x);
      const my = useConst ? sum(ys) / n : 0;
      const ssResid = sum(ys.map((y, i) => (y - fitted[i]) ** 2));
      const ssReg = sum(fitted.map((f) => (f - my) ** 2));
      const df = n - (useConst ? 2 : 1);
      const seY = df > 0 ? Math.sqrt(ssResid / df) : Errors.NUM;
      const mx = useConst ? sum(xs) / n : 0;
      const sxx = sum(xs.map((x) => (x - mx) ** 2));
      const seSlope = typeof seY === 'number' ? seY / Math.sqrt(sxx) : seY;
      const seIntercept = typeof seY === 'number' && useConst ? seY * Math.sqrt(1 / n + (mx * mx) / sxx) : Errors.NA;
      const r2 = ssReg + ssResid === 0 ? 1 : ssReg / (ssReg + ssResid);
      const f = ssResid === 0 ? Errors.NUM : ssReg / 1 / (ssResid / df);
      return new ArrayValue([
        head,
        [seSlope, seIntercept],
        [r2, seY],
        [f, df],
        [ssReg, ssResid],
      ]);
    },
  });
}
linestLike('LINEST', false);
linestLike('LOGEST', true);

define('STANDARDIZE', {
  min: 3,
  max: 3,
  scalar: 'all',
  impl: ([x, m, s], ctx) => {
    const a = num(x, ctx);
    const b = num(m, ctx);
    const c = num(s, ctx);
    const e = err(a, b, c);
    if (e) return e;
    if ((c as number) <= 0) return Errors.NUM;
    return ((a as number) - (b as number)) / (c as number);
  },
});
define('FISHER', {
  min: 1,
  max: 1,
  scalar: 'all',
  impl: ([x], ctx) => {
    const v = num(x, ctx);
    if (v instanceof CellError) return v;
    if (v <= -1 || v >= 1) return Errors.NUM;
    return 0.5 * Math.log((1 + v) / (1 - v));
  },
});
define('FISHERINV', {
  min: 1,
  max: 1,
  scalar: 'all',
  impl: ([x], ctx) => {
    const v = num(x, ctx);
    if (v instanceof CellError) return v;
    const e2 = Math.exp(2 * v);
    return (e2 - 1) / (e2 + 1);
  },
});

// -----------------------------------------------------------------------------------------------
// Distributions and tests (formulajs)

wrap('NORM.DIST', 'NORM.DIST', 'n,n,n,b');
wrap('NORMDIST', 'NORM.DIST', 'n,n,n,b');
wrap('NORM.INV', 'NORM.INV', 'n,n,n');
wrap('NORMINV', 'NORM.INV', 'n,n,n');
wrap('NORM.S.DIST', 'NORM.S.DIST', 'n,b');
wrap('NORMSDIST', 'NORMSDIST', 'n');
wrap('NORM.S.INV', 'NORM.S.INV', 'n');
wrap('NORMSINV', 'NORM.S.INV', 'n');
wrap('LOGNORM.DIST', 'LOGNORM.DIST', 'n,n,n,b');
wrap('LOGNORMDIST', 'LOGNORMDIST', 'n,n,n');
wrap('LOGNORM.INV', 'LOGNORM.INV', 'n,n,n');
wrap('LOGINV', 'LOGNORM.INV', 'n,n,n');
wrap('EXPON.DIST', 'EXPON.DIST', 'n,n,b');
wrap('EXPONDIST', 'EXPON.DIST', 'n,n,b');
wrap('BINOM.DIST', 'BINOM.DIST', 'n,n,n,b');
wrap('BINOMDIST', 'BINOM.DIST', 'n,n,n,b');
wrap('BINOM.DIST.RANGE', 'BINOM.DIST.RANGE', 'n,n,n,n?');
wrap('BINOM.INV', 'BINOM.INV', 'n,n,n');
wrap('CRITBINOM', 'BINOM.INV', 'n,n,n');
wrap('NEGBINOM.DIST', 'NEGBINOM.DIST', 'n,n,n,b');
wrap('NEGBINOMDIST', 'NEGBINOM.DIST', 'n,n,n');
wrap('POISSON.DIST', 'POISSON.DIST', 'n,n,b');
wrap('POISSON', 'POISSON.DIST', 'n,n,b');
wrap('GAMMA', 'GAMMA', 'n');
wrap('GAMMA.DIST', 'GAMMA.DIST', 'n,n,n,b');
wrap('GAMMADIST', 'GAMMA.DIST', 'n,n,n,b');
wrap('GAMMA.INV', 'GAMMA.INV', 'n,n,n');
wrap('GAMMAINV', 'GAMMA.INV', 'n,n,n');
wrap('GAMMALN', 'GAMMALN', 'n');
wrap('GAMMALN.PRECISE', 'GAMMALN', 'n');
wrap('BETA.DIST', 'BETA.DIST', 'n,n,n,b,n?,n?');
wrap('BETADIST', 'BETADIST', 'n,n,n,n?,n?');
wrap('BETA.INV', 'BETA.INV', 'n,n,n,n?,n?');
wrap('BETAINV', 'BETA.INV', 'n,n,n,n?,n?');
wrap('CHISQ.DIST', 'CHISQ.DIST', 'n,n,b');
wrap('CHISQ.DIST.RT', 'CHISQ.DIST.RT', 'n,n');
wrap('CHIDIST', 'CHISQ.DIST.RT', 'n,n');
wrap('CHISQ.INV', 'CHISQ.INV', 'n,n');
wrap('CHISQ.INV.RT', 'CHISQ.INV.RT', 'n,n');
wrap('CHIINV', 'CHISQ.INV.RT', 'n,n');
wrap('CHISQ.TEST', 'CHISQ.TEST', 'm,m');
wrap('CHITEST', 'CHISQ.TEST', 'm,m');
wrap('F.DIST', 'F.DIST', 'n,n,n,b');
wrap('F.DIST.RT', 'F.DIST.RT', 'n,n,n');
wrap('FDIST', 'F.DIST.RT', 'n,n,n');
wrap('F.INV', 'F.INV', 'n,n,n');
wrap('F.INV.RT', 'F.INV.RT', 'n,n,n');
wrap('FINV', 'F.INV.RT', 'n,n,n');
wrap('F.TEST', 'F.TEST', 'a,a');
wrap('FTEST', 'F.TEST', 'a,a');
wrap('T.DIST', 'T.DIST', 'n,n,b');
wrap('T.DIST.2T', 'T.DIST.2T', 'n,n');
wrap('T.DIST.RT', 'T.DIST.RT', 'n,n');
wrap('TDIST', 'TDIST', 'n,n,n');
wrap('T.INV', 'T.INV', 'n,n');
wrap('T.INV.2T', 'T.INV.2T', 'n,n');
wrap('TINV', 'T.INV.2T', 'n,n');
wrap('T.TEST', 'T.TEST', 'a,a,n,n');
wrap('TTEST', 'T.TEST', 'a,a,n,n');
wrap('WEIBULL.DIST', 'WEIBULL.DIST', 'n,n,n,b');
wrap('WEIBULL', 'WEIBULL.DIST', 'n,n,n,b');
wrap('Z.TEST', 'Z.TEST', 'a,n,n?');
wrap('ZTEST', 'Z.TEST', 'a,n,n?');
wrap('CONFIDENCE.NORM', 'CONFIDENCE.NORM', 'n,n,n');
wrap('CONFIDENCE', 'CONFIDENCE.NORM', 'n,n,n');
wrap('CONFIDENCE.T', 'CONFIDENCE.T', 'n,n,n');
wrap('HYPGEOM.DIST', 'HYPGEOM.DIST', 'n,n,n,n,b');
wrap('HYPGEOMDIST', 'HYPGEOMDIST', 'n,n,n,n');
wrap('PHI', 'PHI', 'n');
wrap('GAUSS', 'GAUSS', 'n');
wrap('PROB', 'PROB', 'a,a,n,n?');
wrap('ERF', 'ERF', 'n,n?');
wrap('ERF.PRECISE', 'ERF', 'n');
wrap('ERFC', 'ERFC', 'n');
wrap('ERFC.PRECISE', 'ERFC', 'n');
