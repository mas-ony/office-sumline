import { CellError, Errors, type Scalar } from '../../values';
import { toNumber, type Value } from '../runtime';
import { daysInMonth, isLeapYear, serialFromYMD, ymdFromSerial } from '../dates';
import { define, err, forEachArg, num, type FnContext } from './helpers';
import { wrap } from './fjs';

wrap('CUMIPMT', 'CUMIPMT', 'n,n,n,n,n,n');
wrap('CUMPRINC', 'CUMPRINC', 'n,n,n,n,n,n');
wrap('DB', 'DB', 'n,n,n,n,n?');
wrap('DDB', 'DDB', 'n,n,n,n,n?');
wrap('DISC', 'DISC', 'd,d,n,n,n?');
wrap('DOLLARDE', 'DOLLARDE', 'n,n');
wrap('DOLLARFR', 'DOLLARFR', 'n,n');
wrap('EFFECT', 'EFFECT', 'n,n');
wrap('FV', 'FV', 'n,n,n,n?,n?');
wrap('FVSCHEDULE', 'FVSCHEDULE', 'n,a');
wrap('IPMT', 'IPMT', 'n,n,n,n,n?,n?');
wrap('IRR', 'IRR', 'a,n?');
wrap('ISPMT', 'ISPMT', 'n,n,n,n');
wrap('MIRR', 'MIRR', 'a,n,n');
wrap('NOMINAL', 'NOMINAL', 'n,n');
wrap('NPER', 'NPER', 'n,n,n,n?,n?');
wrap('PDURATION', 'PDURATION', 'n,n,n');
wrap('PMT', 'PMT', 'n,n,n,n?,n?');
wrap('PPMT', 'PPMT', 'n,n,n,n,n?,n?');
wrap('PRICEDISC', 'PRICEDISC', 'd,d,n,n,n?');
wrap('PV', 'PV', 'n,n,n,n?,n?');
wrap('RATE', 'RATE', 'n,n,n,n?,n?,n?');
wrap('RRI', 'RRI', 'n,n,n');
wrap('SLN', 'SLN', 'n,n,n');
wrap('SYD', 'SYD', 'n,n,n,n');
wrap('TBILLEQ', 'TBILLEQ', 'd,d,n');
wrap('TBILLPRICE', 'TBILLPRICE', 'd,d,n');
wrap('TBILLYIELD', 'TBILLYIELD', 'd,d,n');
wrap('XIRR', 'XIRR', 'a,a,n?');
wrap('XNPV', 'XNPV', 'n,a,a');

define('NPV', {
  min: 2,
  max: 255,
  impl: (args, ctx) => {
    const rate = num(args[0], ctx);
    if (rate instanceof CellError) return rate;
    let total = 0;
    let i = 1;
    for (const a of args.slice(1)) {
      const xs = collectAll(a, ctx);
      if (xs instanceof CellError) return xs;
      for (const x of xs) total += x / Math.pow(1 + rate, i++);
    }
    return total;
  },
});

function collectAll(v: Value, ctx: FnContext): number[] | CellError {
  const out: number[] = [];
  let error: CellError | undefined;
  forEachArg([v], ctx, (x, fromRef) => {
    if (error) return;
    if (x instanceof CellError) {
      error = x;
      return;
    }
    if (typeof x === 'number') out.push(x);
    else if (!fromRef && x !== null) {
      const n = toNumber(x);
      if (n instanceof CellError) error = n;
      else out.push(n);
    }
  });
  return error ?? out;
}

// -----------------------------------------------------------------------------------------------
// Coupons and bonds

/** Day count between two serials for an Excel basis (0 US 30/360, 1-3 actual, 4 European 30/360). */
function dayCount(a: number, b: number, basis: number): number {
  if (basis === 0 || basis === 4) {
    let [y1, m1, d1] = ymdFromSerial(a);
    let [y2, m2, d2] = ymdFromSerial(b);
    if (basis === 4) {
      if (d1 === 31) d1 = 30;
      if (d2 === 31) d2 = 30;
    } else {
      const lastFeb = (y: number, m: number, d: number) => m === 2 && d === daysInMonth(y, 2);
      if (lastFeb(y1, m1, d1)) {
        if (lastFeb(y2, m2, d2)) d2 = 30;
        d1 = 30;
      }
      if (d2 === 31 && d1 >= 30) d2 = 30;
      if (d1 === 31) d1 = 30;
    }
    void y1;
    return (y2 - y1) * 360 + (m2 - m1) * 30 + (d2 - d1);
  }
  return b - a;
}

function yearBasis(settlement: number, basis: number): number {
  if (basis === 0 || basis === 2 || basis === 4) return 360;
  if (basis === 3) return 365;
  return isLeapYear(ymdFromSerial(settlement)[0]) ? 366 : 365;
}

function addMonthsEnd(serial: number, months: number, keepEnd: boolean): number {
  const [y, m, d] = ymdFromSerial(serial);
  const total = y * 12 + (m - 1) + months;
  const ny = Math.floor(total / 12);
  const nm = (total % 12) + 1;
  const dim = daysInMonth(ny, nm);
  const endOfMonth = keepEnd && d === daysInMonth(y, m);
  return serialFromYMD(ny, nm, endOfMonth ? dim : Math.min(d, dim));
}

interface CouponArgs {
  settlement: number;
  maturity: number;
  frequency: number;
  basis: number;
}

function couponArgs(args: Value[], ctx: FnContext, si: number, mi: number, fi: number, bi: number): CouponArgs | CellError {
  const s = num(args[si], ctx);
  const m = num(args[mi], ctx);
  const f = num(args[fi], ctx);
  const b = args[bi] == null ? 0 : num(args[bi], ctx);
  const e = err(s, m, f, b);
  if (e) return e;
  const settlement = Math.floor(s as number);
  const maturity = Math.floor(m as number);
  const frequency = Math.trunc(f as number);
  const basis = Math.trunc(b as number);
  if (![1, 2, 4].includes(frequency) || basis < 0 || basis > 4 || settlement >= maturity) return Errors.NUM;
  return { settlement, maturity, frequency, basis };
}

function previousCoupon(a: CouponArgs): number {
  const step = 12 / a.frequency;
  let n = 1;
  let d = addMonthsEnd(a.maturity, -step, true);
  while (d > a.settlement) {
    n++;
    d = addMonthsEnd(a.maturity, -step * n, true);
  }
  return d;
}

function nextCoupon(a: CouponArgs): number {
  return addMonthsEnd(previousCoupon(a), 12 / a.frequency, true);
}

function couponCount(a: CouponArgs): number {
  const step = 12 / a.frequency;
  let n = 0;
  let d = a.maturity;
  while (d > a.settlement) {
    n++;
    d = addMonthsEnd(a.maturity, -step * n, true);
  }
  return n;
}

function couponDays(a: CouponArgs): number {
  if (a.basis === 1) return nextCoupon(a) - previousCoupon(a);
  return (a.basis === 3 ? 365 : 360) / a.frequency;
}

function couponDaysBS(a: CouponArgs): number {
  return dayCount(previousCoupon(a), a.settlement, a.basis);
}

function couponDaysNC(a: CouponArgs): number {
  if (a.basis === 0) return couponDays(a) - couponDaysBS(a);
  return dayCount(a.settlement, nextCoupon(a), a.basis);
}

function couponFn(name: string, f: (a: CouponArgs) => number): void {
  define(name, {
    min: 3,
    max: 4,
    impl: (args, ctx) => {
      const a = couponArgs(args, ctx, 0, 1, 2, 3);
      return a instanceof CellError ? a : f(a);
    },
  });
}
couponFn('COUPPCD', previousCoupon);
couponFn('COUPNCD', nextCoupon);
couponFn('COUPNUM', couponCount);
couponFn('COUPDAYS', couponDays);
couponFn('COUPDAYBS', couponDaysBS);
couponFn('COUPDAYSNC', couponDaysNC);

function bondPrice(a: CouponArgs, rate: number, yld: number, redemption: number): number {
  const N = couponCount(a);
  const E = couponDays(a);
  const A = couponDaysBS(a);
  const DSC = couponDaysNC(a);
  const f = a.frequency;
  const coupon = (100 * rate) / f;
  if (N === 1) {
    return (redemption + coupon) / (1 + ((DSC / E) * yld) / f) - (A / E) * coupon;
  }
  let price = redemption / Math.pow(1 + yld / f, N - 1 + DSC / E);
  for (let k = 1; k <= N; k++) price += coupon / Math.pow(1 + yld / f, k - 1 + DSC / E);
  return price - (A / E) * coupon;
}

define('PRICE', {
  min: 6,
  max: 7,
  impl: (args, ctx) => {
    const a = couponArgs(args, ctx, 0, 1, 5, 6);
    if (a instanceof CellError) return a;
    const rate = num(args[2], ctx);
    const yld = num(args[3], ctx);
    const red = num(args[4], ctx);
    const e = err(rate, yld, red);
    if (e) return e;
    if ((rate as number) < 0 || (yld as number) < 0 || (red as number) <= 0) return Errors.NUM;
    return bondPrice(a, rate as number, yld as number, red as number);
  },
});

define('YIELD', {
  min: 6,
  max: 7,
  impl: (args, ctx) => {
    const a = couponArgs(args, ctx, 0, 1, 5, 6);
    if (a instanceof CellError) return a;
    const rate = num(args[2], ctx);
    const pr = num(args[3], ctx);
    const red = num(args[4], ctx);
    const e = err(rate, pr, red);
    if (e) return e;
    if ((rate as number) < 0 || (pr as number) <= 0 || (red as number) <= 0) return Errors.NUM;
    const N = couponCount(a);
    if (N === 1) {
      const A = couponDaysBS(a);
      const E = couponDays(a);
      const DSR = E - A;
      const f = a.frequency;
      const r = rate as number;
      const num1 = (red as number) / 100 + r / f - ((pr as number) / 100 + (A / E) * (r / f));
      const den = (pr as number) / 100 + (A / E) * (r / f);
      return (num1 / den) * ((f * E) / DSR);
    }
    let y = rate as number || 0.05;
    for (let i = 0; i < 100; i++) {
      const p = bondPrice(a, rate as number, y, red as number) - (pr as number);
      const dp = (bondPrice(a, rate as number, y + 1e-7, red as number) - (pr as number) - p) / 1e-7;
      if (dp === 0) break;
      const next = y - p / dp;
      if (Math.abs(next - y) < 1e-12) return next;
      y = next;
    }
    return Number.isFinite(y) ? y : Errors.NUM;
  },
});

function duration(a: CouponArgs, coupon: number, yld: number): number {
  const N = couponCount(a);
  const E = couponDays(a);
  const DSC = couponDaysNC(a);
  const f = a.frequency;
  const c = (100 * coupon) / f;
  let weighted = 0;
  let total = 0;
  for (let k = 1; k <= N; k++) {
    const t = k - 1 + DSC / E;
    const cf = c + (k === N ? 100 : 0);
    const pv = cf / Math.pow(1 + yld / f, t);
    weighted += (t / f) * pv;
    total += pv;
  }
  return weighted / total;
}

function durationFn(name: string, modified: boolean): void {
  define(name, {
    min: 5,
    max: 6,
    impl: (args, ctx) => {
      const a = couponArgs(args, ctx, 0, 1, 4, 5);
      if (a instanceof CellError) return a;
      const c = num(args[2], ctx);
      const y = num(args[3], ctx);
      const e = err(c, y);
      if (e) return e;
      if ((c as number) < 0 || (y as number) < 0) return Errors.NUM;
      const d = duration(a, c as number, y as number);
      return modified ? d / (1 + (y as number) / a.frequency) : d;
    },
  });
}
durationFn('DURATION', false);
durationFn('MDURATION', true);

function simpleSecurity(args: Value[], ctx: FnContext, basisIndex: number): { s: number; m: number; basis: number } | CellError {
  const s = num(args[0], ctx);
  const m = num(args[1], ctx);
  const b = args[basisIndex] == null ? 0 : num(args[basisIndex], ctx);
  const e = err(s, m, b);
  if (e) return e;
  const basis = Math.trunc(b as number);
  if (basis < 0 || basis > 4 || (s as number) >= (m as number)) return Errors.NUM;
  return { s: Math.floor(s as number), m: Math.floor(m as number), basis };
}

define('INTRATE', {
  min: 4,
  max: 5,
  impl: (args, ctx) => {
    const sec = simpleSecurity(args, ctx, 4);
    if (sec instanceof CellError) return sec;
    const inv = num(args[2], ctx);
    const red = num(args[3], ctx);
    const e = err(inv, red);
    if (e) return e;
    if ((inv as number) <= 0 || (red as number) <= 0) return Errors.NUM;
    const dim = dayCount(sec.s, sec.m, sec.basis);
    return (((red as number) - (inv as number)) / (inv as number)) * (yearBasis(sec.s, sec.basis) / dim);
  },
});
define('RECEIVED', {
  min: 4,
  max: 5,
  impl: (args, ctx) => {
    const sec = simpleSecurity(args, ctx, 4);
    if (sec instanceof CellError) return sec;
    const inv = num(args[2], ctx);
    const disc = num(args[3], ctx);
    const e = err(inv, disc);
    if (e) return e;
    const dim = dayCount(sec.s, sec.m, sec.basis);
    const d = 1 - ((disc as number) * dim) / yearBasis(sec.s, sec.basis);
    if (d <= 0) return Errors.NUM;
    return (inv as number) / d;
  },
});
define('YIELDDISC', {
  min: 4,
  max: 5,
  impl: (args, ctx) => {
    const sec = simpleSecurity(args, ctx, 4);
    if (sec instanceof CellError) return sec;
    const pr = num(args[2], ctx);
    const red = num(args[3], ctx);
    const e = err(pr, red);
    if (e) return e;
    if ((pr as number) <= 0 || (red as number) <= 0) return Errors.NUM;
    const dsm = dayCount(sec.s, sec.m, sec.basis);
    return (((red as number) - (pr as number)) / (pr as number)) * (yearBasis(sec.s, sec.basis) / dsm);
  },
});
define('PRICEMAT', {
  min: 5,
  max: 6,
  impl: (args, ctx) => {
    const sec = simpleSecurity(args, ctx, 5);
    if (sec instanceof CellError) return sec;
    const issue = num(args[2], ctx);
    const rate = num(args[3], ctx);
    const yld = num(args[4], ctx);
    const e = err(issue, rate, yld);
    if (e) return e;
    const B = yearBasis(sec.s, sec.basis);
    const dim = dayCount(issue as number, sec.m, sec.basis);
    const dsm = dayCount(sec.s, sec.m, sec.basis);
    const A = dayCount(issue as number, sec.s, sec.basis);
    return (100 + (dim / B) * (rate as number) * 100) / (1 + (dsm / B) * (yld as number)) - (A / B) * (rate as number) * 100;
  },
});
define('YIELDMAT', {
  min: 5,
  max: 6,
  impl: (args, ctx) => {
    const sec = simpleSecurity(args, ctx, 5);
    if (sec instanceof CellError) return sec;
    const issue = num(args[2], ctx);
    const rate = num(args[3], ctx);
    const pr = num(args[4], ctx);
    const e = err(issue, rate, pr);
    if (e) return e;
    const B = yearBasis(sec.s, sec.basis);
    const dim = dayCount(issue as number, sec.m, sec.basis);
    const dsm = dayCount(sec.s, sec.m, sec.basis);
    const A = dayCount(issue as number, sec.s, sec.basis);
    const top = 1 + (dim / B) * (rate as number) - ((pr as number) / 100 + (A / B) * (rate as number));
    return (top / ((pr as number) / 100 + (A / B) * (rate as number))) * (B / dsm);
  },
});

define('VDB', {
  min: 5,
  max: 7,
  impl: (args, ctx) => {
    const cost = num(args[0], ctx);
    const salvage = num(args[1], ctx);
    const life = num(args[2], ctx);
    const start = num(args[3], ctx);
    const end = num(args[4], ctx);
    const factor = args[5] == null ? 2 : num(args[5], ctx);
    const noSwitch = args[6] == null ? false : Boolean(toNumber(args[6] as Scalar));
    const e = err(cost, salvage, life, start, end, factor);
    if (e) return e;
    const c = cost as number;
    const s = salvage as number;
    const l = life as number;
    const a = start as number;
    const b = end as number;
    const fct = factor as number;
    if (c < 0 || s < 0 || l <= 0 || a < 0 || b < a || b > l || fct <= 0) return Errors.NUM;
    let value = c;
    let total = 0;
    const periods = Math.ceil(b);
    let switched = false;
    for (let p = 0; p < periods; p++) {
      let dep: number;
      const ddb = Math.min((value * fct) / l, Math.max(0, value - s));
      const sln = (value - s) / (l - p);
      if (!noSwitch && (switched || sln > ddb)) {
        switched = true;
        dep = Math.max(0, sln);
      } else dep = ddb;
      const from = Math.max(a, p);
      const to = Math.min(b, p + 1);
      if (to > from) total += dep * (to - from);
      value -= dep;
    }
    return total;
  },
});
