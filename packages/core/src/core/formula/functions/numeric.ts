/** Numeric building blocks shared by math, statistics and aggregate functions. */
import { Errors, type CellError } from '../../values';

/** Rounds half away from zero at `digits` decimals, based on the decimal representation. */
export function roundHalfAway(x: number, digits: number): number {
  if (!Number.isFinite(x)) return x;
  const d = Math.trunc(digits);
  const sign = x < 0 ? -1 : 1;
  const ax = Math.abs(Number(x.toPrecision(15)));
  if (d >= 0) {
    const r = Math.round(Number(`${ax}e${d}`));
    return sign * Number(`${r}e-${d}`);
  }
  const f = Math.pow(10, -d);
  return sign * Math.round(ax / f) * f;
}

export function roundUp(x: number, digits: number): number {
  const d = Math.trunc(digits);
  const sign = x < 0 ? -1 : 1;
  const ax = Math.abs(Number(x.toPrecision(15)));
  if (d >= 0) {
    const scaled = Number(Number(`${ax}e${d}`).toPrecision(15));
    return sign * Number(`${Math.ceil(scaled)}e-${d}`);
  }
  const f = Math.pow(10, -d);
  return sign * Math.ceil(Number((ax / f).toPrecision(15))) * f;
}

export function roundDown(x: number, digits: number): number {
  const d = Math.trunc(digits);
  const sign = x < 0 ? -1 : 1;
  const ax = Math.abs(Number(x.toPrecision(15)));
  if (d >= 0) {
    const scaled = Number(Number(`${ax}e${d}`).toPrecision(15));
    return sign * Number(`${Math.floor(scaled)}e-${d}`);
  }
  const f = Math.pow(10, -d);
  return sign * Math.floor(Number((ax / f).toPrecision(15))) * f;
}

export function sum(xs: number[]): number {
  // Neumaier summation keeps 0.1+0.2+... closer to Excel's results.
  let s = 0;
  let c = 0;
  for (const x of xs) {
    const t = s + x;
    if (Math.abs(s) >= Math.abs(x)) c += s - t + x;
    else c += x - t + s;
    s = t;
  }
  return s + c;
}

export function mean(xs: number[]): number | CellError {
  if (!xs.length) return Errors.DIV0;
  return sum(xs) / xs.length;
}

export function variance(xs: number[], sample: boolean): number | CellError {
  const n = xs.length;
  if (n === 0 || (sample && n === 1)) return Errors.DIV0;
  const m = sum(xs) / n;
  let ss = 0;
  for (const x of xs) ss += (x - m) * (x - m);
  return ss / (sample ? n - 1 : n);
}

export function stdev(xs: number[], sample: boolean): number | CellError {
  const v = variance(xs, sample);
  return typeof v === 'number' ? Math.sqrt(v) : v;
}

export function median(xs: number[]): number | CellError {
  if (!xs.length) return Errors.NUM;
  const s = [...xs].sort((a, b) => a - b);
  const mid = s.length >> 1;
  return s.length % 2 ? s[mid] : (s[mid - 1] + s[mid]) / 2;
}

export function modeSingle(xs: number[]): number | CellError {
  const counts = new Map<number, number>();
  let best: number | null = null;
  let bestCount = 1;
  for (const x of xs) {
    const c = (counts.get(x) ?? 0) + 1;
    counts.set(x, c);
    if (c > bestCount) {
      bestCount = c;
      best = x;
    }
  }
  return best === null ? Errors.NA : best;
}

export function percentileInc(xs: number[], p: number): number | CellError {
  if (!xs.length || p < 0 || p > 1) return Errors.NUM;
  const s = [...xs].sort((a, b) => a - b);
  const h = (s.length - 1) * p;
  const lo = Math.floor(h);
  return s[lo] + (h - lo) * ((s[lo + 1] ?? s[lo]) - s[lo]);
}

export function percentileExc(xs: number[], p: number): number | CellError {
  const n = xs.length;
  if (!n || p <= 0 || p >= 1) return Errors.NUM;
  const s = [...xs].sort((a, b) => a - b);
  const h = (n + 1) * p;
  if (h < 1 || h > n) return Errors.NUM;
  const lo = Math.floor(h);
  return s[lo - 1] + (h - lo) * ((s[lo] ?? s[lo - 1]) - s[lo - 1]);
}

export function quartile(xs: number[], q: number, exclusive: boolean): number | CellError {
  const k = Math.trunc(q);
  if (exclusive) {
    if (k <= 0 || k >= 4) return Errors.NUM;
    return percentileExc(xs, k / 4);
  }
  if (k < 0 || k > 4) return Errors.NUM;
  return percentileInc(xs, k / 4);
}

export function kth(xs: number[], k: number, largest: boolean): number | CellError {
  const n = Math.ceil(k);
  if (!xs.length || n < 1 || n > xs.length) return Errors.NUM;
  const s = [...xs].sort((a, b) => (largest ? b - a : a - b));
  return s[n - 1];
}

export function gcd2(a: number, b: number): number {
  a = Math.abs(Math.trunc(a));
  b = Math.abs(Math.trunc(b));
  while (b) [a, b] = [b, a % b];
  return a;
}

export function factorial(n: number): number {
  const k = Math.trunc(n);
  if (k < 0) return NaN;
  let r = 1;
  for (let i = 2; i <= k; i++) r *= i;
  return r;
}

export function combin(n: number, k: number): number {
  n = Math.trunc(n);
  k = Math.trunc(k);
  if (n < 0 || k < 0 || k > n) return NaN;
  k = Math.min(k, n - k);
  let r = 1;
  for (let i = 1; i <= k; i++) r = (r * (n - k + i)) / i;
  return Math.round(r);
}
