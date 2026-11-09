/** Display text of cell values under Excel number formats. */
import { format, formatColor, getFormatInfo, isDateFormat } from 'numfmt';
import { CellError, numberToString, type Scalar } from '../values';

export interface Formatted {
  text: string;
  /** Color from the format ([Red] etc.), as CSS. */
  color?: string;
  /** True when the formatted value is a number (right aligned by default). */
  numeric: boolean;
}

const COLOR_NAMES: Record<string, string> = {
  black: '#000000',
  blue: '#0000FF',
  cyan: '#00FFFF',
  green: '#00FF00',
  magenta: '#FF00FF',
  red: '#FF0000',
  white: '#FFFFFF',
  yellow: '#FFFF00',
};

const infoCache = new Map<string, ReturnType<typeof getFormatInfo> | null>();

export function formatInfo(pattern: string): ReturnType<typeof getFormatInfo> | null {
  let info = infoCache.get(pattern);
  if (info === undefined) {
    try {
      info = getFormatInfo(pattern);
    } catch {
      info = null;
    }
    if (infoCache.size > 2000) infoCache.clear();
    infoCache.set(pattern, info);
  }
  return info;
}

export function isDatePattern(pattern: string): boolean {
  try {
    return isDateFormat(pattern);
  } catch {
    return false;
  }
}

/** General format: up to 11 characters like Excel's default column width allows. */
export function formatGeneral(n: number): string {
  if (!Number.isFinite(n)) return '#NUM!';
  if (n === 0) return '0';
  const abs = Math.abs(n);
  if (abs >= 1e11 || abs < 1e-9) {
    const s = n.toExponential(5).replace(/\.?0+e/, 'e');
    const [m, e] = s.split('e');
    const exp = Number(e);
    return `${m}E${exp < 0 ? '-' : '+'}${String(Math.abs(exp)).padStart(2, '0')}`;
  }
  // 10 significant digits fit the default 8.43 character column.
  const digits = abs >= 1 ? Math.max(1, 10 - Math.floor(Math.log10(abs))) : 10;
  let s = n.toFixed(Math.min(digits, 10));
  if (s.includes('.')) s = s.replace(/0+$/, '').replace(/\.$/, '');
  if (s.replace('-', '').replace('.', '').length > 11) return numberToString(Number(n.toPrecision(10)));
  return s;
}

export function formatValue(value: Scalar, pattern: string, date1904 = false): Formatted {
  if (value === null) return { text: '', numeric: false };
  if (value instanceof CellError) return { text: value.code, numeric: false };
  if (typeof value === 'boolean') return { text: value ? 'TRUE' : 'FALSE', numeric: false };
  const p = pattern || 'General';
  if (typeof value === 'number') {
    if (p === 'General') return { text: formatGeneral(value), numeric: true };
    let v = value;
    if (date1904 && isDatePattern(p)) v += 1462;
    try {
      const text = format(p, v, { throws: false, leap1900: true });
      const c = formatColor(p, v, { throws: false });
      const color = typeof c === 'string' ? COLOR_NAMES[c.toLowerCase()] ?? (c.startsWith('#') ? c : undefined) : undefined;
      return { text, color, numeric: true };
    } catch {
      return { text: formatGeneral(value), numeric: true };
    }
  }
  if (p === 'General' || p === '@') return { text: value, numeric: false };
  try {
    // Text sections of a format (4th section or '@') apply to strings.
    const text = format(p, value, { throws: false });
    return { text, numeric: false };
  } catch {
    return { text: value, numeric: false };
  }
}

/** The text shown in the formula bar / cell editor for a cell value (not formatted). */
export function editText(value: Scalar, pattern: string, date1904 = false): string {
  if (value === null) return '';
  if (value instanceof CellError) return value.code;
  if (typeof value === 'boolean') return value ? 'TRUE' : 'FALSE';
  if (typeof value === 'string') return value;
  const info = formatInfo(pattern);
  if (info?.isDate || info?.type === 'time' || info?.type === 'datetime') {
    const v = date1904 ? value + 1462 : value;
    const hasTime = Math.abs(v - Math.floor(v)) > 1e-9;
    const datePart = info.type === 'time' ? '' : 'm/d/yyyy';
    const timePart = hasTime || info.type === 'time' ? 'h:mm:ss AM/PM' : '';
    const p = [datePart, timePart].filter(Boolean).join(' ');
    try {
      return format(p, v, { leap1900: true });
    } catch {
      return numberToString(value);
    }
  }
  if (info?.isPercent) {
    const s = numberToString(Number((value * 100).toPrecision(15)));
    return `${s}%`;
  }
  return numberToString(value);
}
