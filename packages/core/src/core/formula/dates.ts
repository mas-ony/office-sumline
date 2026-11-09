/** Excel serial dates (1900 system with the Lotus leap-year bug, and the 1904 system). */
import { parseDate, parseTime } from 'numfmt';

const MS_PER_DAY = 86400000;
/** Days between 1899-12-30 and 1970-01-01. */
const EPOCH_OFFSET = 25569;

let use1904 = false;

/** Sets the date system of the workbook being calculated. */
export function setDate1904(on: boolean): void {
  use1904 = on;
}

export function isDate1904(): boolean {
  return use1904;
}

/** Serial number for a calendar date (month and day may overflow like DATE()). */
export function serialFromYMD(y: number, m: number, d: number): number {
  const ms = Date.UTC(y, m - 1, 1) + (d - 1) * MS_PER_DAY;
  let serial = Math.round(ms / MS_PER_DAY) + EPOCH_OFFSET;
  if (use1904) return serial - 1462;
  // 1900-01-01 .. 1900-02-28 are one lower because Excel counts the fake 1900-02-29.
  if (serial < 61) serial -= 1;
  return serial;
}

/** [year, month (1-12), day] of a serial number. */
export function ymdFromSerial(serial: number): [number, number, number] {
  let s = Math.floor(serial);
  if (use1904) s += 1462;
  else {
    if (s === 60) return [1900, 2, 29];
    if (s === 0) return [1900, 1, 0];
    if (s < 60) s += 1;
  }
  const date = new Date((s - EPOCH_OFFSET) * MS_PER_DAY);
  return [date.getUTCFullYear(), date.getUTCMonth() + 1, date.getUTCDate()];
}

/** Day of week 0 = Sunday for a serial number (Excel treats serial 0..59 consistently). */
export function weekdayOf(serial: number): number {
  const s = Math.floor(serial) + (use1904 ? 1462 : 0);
  // Serial 1 (1900-01-01) is a Sunday in Excel's (buggy) calendar.
  return ((s % 7) + 6) % 7;
}

export function timeParts(serial: number): [number, number, number] {
  let frac = serial - Math.floor(serial);
  // Round to the nearest second to avoid 59.9999 artefacts.
  let secs = Math.round(frac * 86400);
  if (secs >= 86400) secs = 0;
  frac = secs;
  const h = Math.floor(frac / 3600);
  const m = Math.floor((frac % 3600) / 60);
  const s = frac % 60;
  return [h, m, s];
}

export function serialFromJsDate(d: Date): number {
  const ms = Date.UTC(d.getFullYear(), d.getMonth(), d.getDate(), d.getHours(), d.getMinutes(), d.getSeconds(), d.getMilliseconds());
  let serial = ms / MS_PER_DAY + EPOCH_OFFSET;
  if (use1904) serial -= 1462;
  else if (serial < 61) serial -= 1;
  return serial;
}

/** A JS Date (local time fields) for a serial number. */
export function jsDateFromSerial(serial: number): Date {
  const [y, m, d] = ymdFromSerial(serial);
  const [h, mi, s] = timeParts(serial);
  return new Date(y, m - 1, d, h, mi, s);
}

export function daysInMonth(y: number, m: number): number {
  return new Date(Date.UTC(y, m, 0)).getUTCDate();
}

export function isLeapYear(y: number): boolean {
  return (y % 4 === 0 && y % 100 !== 0) || y % 400 === 0;
}

export function parseDateText(text: string): number | null {
  const r = parseDate(text) ?? null;
  if (!r || typeof r.v !== 'number') return null;
  return use1904 ? r.v - 1462 : r.v;
}

export function parseTimeText(text: string): number | null {
  const r = parseTime(text) ?? null;
  if (!r || typeof r.v !== 'number') return null;
  return r.v;
}

export function todaySerial(): number {
  const now = new Date();
  return serialFromYMD(now.getFullYear(), now.getMonth() + 1, now.getDate());
}

export function nowSerial(): number {
  return serialFromJsDate(new Date());
}
