import { CellError, Errors, type Scalar } from '../../values';
import { textToNumber, toNumber, toText, type Value } from '../runtime';
import {
  daysInMonth, isLeapYear, nowSerial, parseDateText, parseTimeText, serialFromYMD, timeParts, todaySerial, weekdayOf, ymdFromSerial,
} from '../dates';
import { collectNumbers, define, defineScalar, err, num, toScalar, type FnContext } from './helpers';
import { wrap } from './fjs';

function serialArg(v: Scalar | undefined): number | CellError {
  if (v === undefined || v === null) return 0;
  const n = toNumber(v);
  if (n instanceof CellError) return n;
  if (n < 0) return Errors.NUM;
  return n;
}

defineScalar('DATE', 3, 3, ([y, m, d]) => {
  const yy = toNumber(y);
  const mm = toNumber(m);
  const dd = toNumber(d);
  const e = err(yy, mm, dd);
  if (e) return e;
  let year = Math.trunc(yy as number);
  if (year < 0 || year >= 10000) return Errors.NUM;
  if (year < 1900) year += 1900;
  const s = serialFromYMD(year, Math.trunc(mm as number), Math.trunc(dd as number));
  return s < 0 ? Errors.NUM : s;
});
defineScalar('TIME', 3, 3, ([h, m, s]) => {
  const hh = toNumber(h);
  const mm = toNumber(m);
  const ss = toNumber(s);
  const e = err(hh, mm, ss);
  if (e) return e;
  const total = Math.trunc(hh as number) * 3600 + Math.trunc(mm as number) * 60 + Math.trunc(ss as number);
  if (total < 0) return Errors.NUM;
  return (total % 86400) / 86400;
});
defineScalar('YEAR', 1, 1, ([v]) => {
  const s = serialArg(v);
  return s instanceof CellError ? s : ymdFromSerial(s)[0];
});
defineScalar('MONTH', 1, 1, ([v]) => {
  const s = serialArg(v);
  return s instanceof CellError ? s : ymdFromSerial(s)[1];
});
defineScalar('DAY', 1, 1, ([v]) => {
  const s = serialArg(v);
  return s instanceof CellError ? s : ymdFromSerial(s)[2];
});
defineScalar('HOUR', 1, 1, ([v]) => {
  const s = serialArg(v);
  return s instanceof CellError ? s : timeParts(s)[0];
});
defineScalar('MINUTE', 1, 1, ([v]) => {
  const s = serialArg(v);
  return s instanceof CellError ? s : timeParts(s)[1];
});
defineScalar('SECOND', 1, 1, ([v]) => {
  const s = serialArg(v);
  return s instanceof CellError ? s : timeParts(s)[2];
});
define('NOW', { min: 0, max: 0, volatile: true, impl: () => nowSerial() });
define('TODAY', { min: 0, max: 0, volatile: true, impl: () => todaySerial() });
defineScalar('DATEVALUE', 1, 1, ([v]) => {
  if (typeof v !== 'string') return Errors.VALUE;
  const d = parseDateText(v);
  if (d === null) {
    const n = textToNumber(v);
    return n === null ? Errors.VALUE : Errors.VALUE;
  }
  return Math.floor(d);
});
defineScalar('TIMEVALUE', 1, 1, ([v]) => {
  if (typeof v !== 'string') return Errors.VALUE;
  const t = parseTimeText(v);
  if (t !== null) return t - Math.floor(t);
  const d = parseDateText(v);
  return d === null ? Errors.VALUE : d - Math.floor(d);
});
defineScalar('WEEKDAY', 1, 2, ([v, type]) => {
  const s = serialArg(v);
  const t = type === undefined ? 1 : toNumber(type);
  const e = err(s, t);
  if (e) return e;
  const dow = weekdayOf(s as number); // 0 = Sunday
  switch (Math.trunc(t as number)) {
    case 1:
    case 17:
      return dow + 1;
    case 2:
    case 11:
      return ((dow + 6) % 7) + 1;
    case 3:
      return (dow + 6) % 7;
    case 12:
      return ((dow + 5) % 7) + 1;
    case 13:
      return ((dow + 4) % 7) + 1;
    case 14:
      return ((dow + 3) % 7) + 1;
    case 15:
      return ((dow + 2) % 7) + 1;
    case 16:
      return ((dow + 1) % 7) + 1;
    default:
      return Errors.NUM;
  }
});

function isoWeek(serial: number): number {
  const [y, m, d] = ymdFromSerial(serial);
  const date = new Date(Date.UTC(y, m - 1, d));
  const day = (date.getUTCDay() + 6) % 7; // Monday = 0
  date.setUTCDate(date.getUTCDate() - day + 3);
  const firstThursday = new Date(Date.UTC(date.getUTCFullYear(), 0, 4));
  return 1 + Math.round(((date.getTime() - firstThursday.getTime()) / 86400000 - 3 + ((firstThursday.getUTCDay() + 6) % 7)) / 7);
}

defineScalar('ISOWEEKNUM', 1, 1, ([v]) => {
  const s = serialArg(v);
  return s instanceof CellError ? s : isoWeek(s);
});
defineScalar('WEEKNUM', 1, 2, ([v, type]) => {
  const s = serialArg(v);
  const t = type === undefined ? 1 : toNumber(type);
  const e = err(s, t);
  if (e) return e;
  const kind = Math.trunc(t as number);
  if (kind === 21) return isoWeek(s as number);
  const startMap: Record<number, number> = { 1: 0, 17: 0, 2: 1, 11: 1, 12: 2, 13: 3, 14: 4, 15: 5, 16: 6 };
  const start = startMap[kind];
  if (start === undefined) return Errors.NUM;
  const [y] = ymdFromSerial(s as number);
  const jan1 = serialFromYMD(y, 1, 1);
  const offset = (weekdayOf(jan1) - start + 7) % 7;
  return Math.floor((Math.floor(s as number) - jan1 + offset) / 7) + 1;
});

function addMonths(serial: number, months: number, endOfMonth: boolean): number | CellError {
  const [y, m, d] = ymdFromSerial(serial);
  const total = y * 12 + (m - 1) + Math.trunc(months);
  const ny = Math.floor(total / 12);
  const nm = (total % 12) + 1;
  if (ny < 1900 || ny > 9999) return Errors.NUM;
  const dim = daysInMonth(ny, nm);
  return serialFromYMD(ny, nm, endOfMonth ? dim : Math.min(d, dim));
}
defineScalar('EDATE', 2, 2, ([v, months]) => {
  const s = serialArg(v);
  const k = toNumber(months);
  const e = err(s, k);
  return e ?? addMonths(Math.floor(s as number), k as number, false);
});
defineScalar('EOMONTH', 2, 2, ([v, months]) => {
  const s = serialArg(v);
  const k = toNumber(months);
  const e = err(s, k);
  return e ?? addMonths(Math.floor(s as number), k as number, true);
});
defineScalar('DAYS', 2, 2, ([end, start]) => {
  const a = serialArg(end);
  const b = serialArg(start);
  const e = err(a, b);
  return e ?? Math.floor(a as number) - Math.floor(b as number);
});
defineScalar('DAYS360', 2, 3, ([start, end, method]) => {
  const a = serialArg(start);
  const b = serialArg(end);
  const e = err(a, b);
  if (e) return e;
  const european = method !== undefined && (method === true || toNumber(method) === 1);
  let [y1, m1, d1] = ymdFromSerial(a as number);
  let [y2, m2, d2] = ymdFromSerial(b as number);
  if (european) {
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
  void isLeapYear;
  return (y2 - y1) * 360 + (m2 - m1) * 30 + (d2 - d1);
});
defineScalar('DATEDIF', 3, 3, ([start, end, unit]) => {
  const a = serialArg(start);
  const b = serialArg(end);
  const u = toText(unit ?? null);
  const e = err(a, b, u);
  if (e) return e;
  const s1 = Math.floor(a as number);
  const s2 = Math.floor(b as number);
  if (s1 > s2) return Errors.NUM;
  const [y1, m1, d1] = ymdFromSerial(s1);
  const [y2, m2, d2] = ymdFromSerial(s2);
  switch ((u as string).toUpperCase()) {
    case 'Y':
      return y2 - y1 - (m2 < m1 || (m2 === m1 && d2 < d1) ? 1 : 0);
    case 'M':
      return (y2 - y1) * 12 + (m2 - m1) - (d2 < d1 ? 1 : 0);
    case 'D':
      return s2 - s1;
    case 'MD': {
      if (d2 >= d1) return d2 - d1;
      const pm = m2 === 1 ? 12 : m2 - 1;
      const py = m2 === 1 ? y2 - 1 : y2;
      return daysInMonth(py, pm) - d1 + d2;
    }
    case 'YM':
      return (m2 - m1 + 12 - (d2 < d1 ? 1 : 0)) % 12;
    case 'YD': {
      let start2 = serialFromYMD(y2, m1, d1);
      if (start2 > s2) start2 = serialFromYMD(y2 - 1, m1, d1);
      return s2 - start2;
    }
    default:
      return Errors.NUM;
  }
});

// -----------------------------------------------------------------------------------------------
// Working days

function weekendMask(arg: Value | undefined, ctx: FnContext): boolean[] | CellError {
  // Index 0 = Monday .. 6 = Sunday.
  if (arg === undefined || arg === null) return [false, false, false, false, false, true, true];
  const v = toScalar(arg, ctx);
  if (typeof v === 'string') {
    if (!/^[01]{7}$/.test(v) || v === '1111111') return Errors.VALUE;
    return [...v].map((c) => c === '1');
  }
  const n = toNumber(v);
  if (n instanceof CellError) return n;
  const code = Math.trunc(n);
  const pairs: Record<number, [number, number]> = { 1: [5, 6], 2: [6, 0], 3: [0, 1], 4: [1, 2], 5: [2, 3], 6: [3, 4], 7: [4, 5] };
  const mask = [false, false, false, false, false, false, false];
  if (pairs[code]) {
    mask[pairs[code][0]] = true;
    mask[pairs[code][1]] = true;
    return mask;
  }
  if (code >= 11 && code <= 17) {
    mask[code - 11] = true;
    return mask;
  }
  return Errors.NUM;
}

function holidaySet(arg: Value | undefined, ctx: FnContext): Set<number> | CellError {
  if (arg === undefined || arg === null) return new Set();
  const xs = collectNumbers([arg], ctx);
  if (xs instanceof CellError) return xs;
  return new Set(xs.map((x) => Math.floor(x)));
}

function isWorkday(serial: number, mask: boolean[], holidays: Set<number>): boolean {
  const mondayIndex = (weekdayOf(serial) + 6) % 7;
  return !mask[mondayIndex] && !holidays.has(serial);
}

function networkDays(start: number, end: number, mask: boolean[], holidays: Set<number>): number {
  const a = Math.floor(Math.min(start, end));
  const b = Math.floor(Math.max(start, end));
  let count = 0;
  // Whole weeks first for speed.
  const workPerWeek = mask.filter((x) => !x).length;
  const days = b - a + 1;
  const weeks = Math.floor(days / 7);
  count += weeks * workPerWeek;
  for (let s = a + weeks * 7; s <= b; s++) if (isWorkday(s, mask, new Set())) count++;
  for (const h of holidays) if (h >= a && h <= b && isWorkday(h, mask, new Set())) count--;
  return start <= end ? count : -count;
}

function networkDaysImpl(intl: boolean) {
  return (args: Value[], ctx: FnContext): Value => {
    const a = num(args[0], ctx);
    const b = num(args[1], ctx);
    const e = err(a, b);
    if (e) return e;
    const mask = intl ? weekendMask(args[2], ctx) : weekendMask(undefined, ctx);
    if (mask instanceof CellError) return mask;
    const hol = holidaySet(intl ? args[3] : args[2], ctx);
    if (hol instanceof CellError) return hol;
    return networkDays(a as number, b as number, mask, hol);
  };
}
define('NETWORKDAYS', { min: 2, max: 3, impl: networkDaysImpl(false) });
define('NETWORKDAYS.INTL', { min: 2, max: 4, impl: networkDaysImpl(true) });

function workdayImpl(intl: boolean) {
  return (args: Value[], ctx: FnContext): Value => {
    const a = num(args[0], ctx);
    const n = num(args[1], ctx);
    const e = err(a, n);
    if (e) return e;
    const mask = intl ? weekendMask(args[2], ctx) : weekendMask(undefined, ctx);
    if (mask instanceof CellError) return mask;
    const hol = holidaySet(intl ? args[3] : args[2], ctx);
    if (hol instanceof CellError) return hol;
    let s = Math.floor(a as number);
    let left = Math.trunc(n as number);
    const step = left >= 0 ? 1 : -1;
    while (left !== 0) {
      s += step;
      if (isWorkday(s, mask, hol)) left -= step;
      if (s < 0 || s > 2958465) return Errors.NUM;
    }
    return s;
  };
}
define('WORKDAY', { min: 2, max: 3, impl: workdayImpl(false) });
define('WORKDAY.INTL', { min: 2, max: 4, impl: workdayImpl(true) });

defineScalar('YEARFRAC', 2, 3, ([start, end, basisArg]) => {
  const a = serialArg(start);
  const b = serialArg(end);
  const basis = basisArg === undefined ? 0 : toNumber(basisArg);
  const e = err(a, b, basis);
  if (e) return e;
  let s1 = Math.floor(a as number);
  let s2 = Math.floor(b as number);
  if (s1 > s2) [s1, s2] = [s2, s1];
  const [y1, m1, d1] = ymdFromSerial(s1);
  const [y2, m2, d2] = ymdFromSerial(s2);
  switch (Math.trunc(basis as number)) {
    case 0: {
      let dd1 = d1;
      let dd2 = d2;
      const lastFeb = (y: number, m: number, d: number) => m === 2 && d === daysInMonth(y, 2);
      if (dd1 === 31 && dd2 === 31) {
        dd1 = 30;
        dd2 = 30;
      } else if (dd1 === 31) dd1 = 30;
      else if (dd1 === 30 && dd2 === 31) dd2 = 30;
      else if (lastFeb(y1, m1, d1) && lastFeb(y2, m2, d2)) {
        dd1 = 30;
        dd2 = 30;
      } else if (lastFeb(y1, m1, d1)) dd1 = 30;
      return ((y2 - y1) * 360 + (m2 - m1) * 30 + (dd2 - dd1)) / 360;
    }
    case 1: {
      const days = s2 - s1;
      if (y1 === y2 || (y2 === y1 + 1 && (m1 > m2 || (m1 === m2 && d1 >= d2)))) {
        let year = 365;
        if (y1 === y2 && isLeapYear(y1)) year = 366;
        else if (y1 !== y2) {
          const feb29 = (y: number) => serialFromYMD(y, 2, 29);
          if ((isLeapYear(y1) && s1 <= feb29(y1)) || (isLeapYear(y2) && s2 >= feb29(y2))) year = 366;
        }
        return days / year;
      }
      let total = 0;
      for (let y = y1; y <= y2; y++) total += isLeapYear(y) ? 366 : 365;
      return days / (total / (y2 - y1 + 1));
    }
    case 2:
      return (s2 - s1) / 360;
    case 3:
      return (s2 - s1) / 365;
    case 4: {
      const dd1 = Math.min(d1, 30);
      const dd2 = Math.min(d2, 30);
      return ((y2 - y1) * 360 + (m2 - m1) * 30 + (dd2 - dd1)) / 360;
    }
    default:
      return Errors.NUM;
  }
});

// Financial functions with date arguments are implemented by formulajs.
wrap('ACCRINT', 'ACCRINT', 'd,d,d,n,n,n,n?,b?');
void wrap;
