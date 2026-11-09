import { format as numfmtFormat } from 'numfmt';
import { CellError, Errors, numberToString, type Scalar } from '../../values';
import { ArrayValue, textToNumber, toBoolean, toNumber, toText, type Value } from '../runtime';
import { wildcardToRegExp } from '../criteria';
import { isDate1904 } from '../dates';
import { alias, define, defineScalar, err, forEachArg, toArray, toScalar, type FnContext } from './helpers';
import { roundHalfAway } from './numeric';

/** Windows-1252 characters for codes 128..159 (CHAR/CODE follow the ANSI code page on Windows). */
const CP1252: Record<number, number> = {
  128: 0x20ac, 130: 0x201a, 131: 0x0192, 132: 0x201e, 133: 0x2026, 134: 0x2020, 135: 0x2021, 136: 0x02c6, 137: 0x2030,
  138: 0x0160, 139: 0x2039, 140: 0x0152, 142: 0x017d, 145: 0x2018, 146: 0x2019, 147: 0x201c, 148: 0x201d, 149: 0x2022,
  150: 0x2013, 151: 0x2014, 152: 0x02dc, 153: 0x2122, 154: 0x0161, 155: 0x203a, 156: 0x0153, 158: 0x017e, 159: 0x0178,
};
const CP1252_REVERSE = new Map(Object.entries(CP1252).map(([k, v]) => [v, Number(k)]));

function str(v: Scalar | undefined): string | CellError {
  return toText(v ?? null);
}

function n(v: Scalar | undefined, dflt?: number): number | CellError {
  if (v === undefined) return dflt ?? Errors.VALUE;
  return toNumber(v);
}

/** Formats a value with an Excel number format (also used by the TEXT function). */
export function formatWithPattern(value: Scalar, pattern: string): string {
  let v: Scalar = value;
  if (typeof v === 'string') {
    const parsed = textToNumber(v);
    if (parsed !== null) v = parsed;
  }
  if (typeof v === 'number' && isDate1904()) {
    // numfmt formats 1900-system serials.
    v = v + 1462;
  }
  try {
    return numfmtFormat(pattern, v instanceof CellError ? v.code : v, { throws: false, leap1900: true });
  } catch {
    return typeof v === 'number' ? numberToString(v) : String(v ?? '');
  }
}

defineScalar('CHAR', 1, 1, ([x]) => {
  const c = n(x);
  if (c instanceof CellError) return c;
  const code = Math.trunc(c);
  if (code < 1 || code > 255) return Errors.VALUE;
  return String.fromCharCode(CP1252[code] ?? code);
});
defineScalar('CODE', 1, 1, ([x]) => {
  const s = str(x);
  if (s instanceof CellError) return s;
  if (!s) return Errors.VALUE;
  const cp = s.charCodeAt(0);
  return CP1252_REVERSE.get(cp) ?? (cp > 255 ? 63 : cp);
});
defineScalar('UNICHAR', 1, 1, ([x]) => {
  const c = n(x);
  if (c instanceof CellError) return c;
  const code = Math.trunc(c);
  if (code < 1 || code > 0x10ffff) return Errors.VALUE;
  try {
    return String.fromCodePoint(code);
  } catch {
    return Errors.VALUE;
  }
});
defineScalar('UNICODE', 1, 1, ([x]) => {
  const s = str(x);
  if (s instanceof CellError) return s;
  if (!s) return Errors.VALUE;
  return s.codePointAt(0)!;
});
defineScalar('CLEAN', 1, 1, ([x]) => {
  const s = str(x);
  // eslint-disable-next-line no-control-regex
  return s instanceof CellError ? s : s.replace(/[\x00-\x1f]/g, '');
});
defineScalar('CONCATENATE', 1, 255, (args) => {
  let out = '';
  for (const a of args) {
    const s = str(a);
    if (s instanceof CellError) return s;
    out += s;
  }
  return out;
});
define('CONCAT', {
  min: 1,
  max: 254,
  impl: (args, ctx) => {
    let out = '';
    let error: CellError | undefined;
    forEachArg(args, ctx, (v) => {
      if (error) return;
      const s = toText(v);
      if (s instanceof CellError) error = s;
      else out += s;
    });
    if (error) return error;
    return out.length > 32767 ? Errors.VALUE : out;
  },
});
define('TEXTJOIN', {
  min: 3,
  max: 255,
  impl: (args, ctx) => {
    const delims = toArray(args[0], ctx).data.flat().map((d) => toText(d));
    const e0 = delims.find((d) => d instanceof CellError);
    if (e0) return e0 as CellError;
    const ignore = toBoolean(toScalar(args[1], ctx));
    if (ignore instanceof CellError) return ignore;
    const parts: string[] = [];
    let error: CellError | undefined;
    forEachArg(args.slice(2), ctx, (v) => {
      if (error) return;
      const s = toText(v);
      if (s instanceof CellError) error = s;
      else if (!(ignore && s === '')) parts.push(s);
    });
    if (error) return error;
    let out = '';
    parts.forEach((p, i) => {
      if (i > 0) out += (delims[(i - 1) % delims.length] as string) ?? '';
      out += p;
    });
    return out.length > 32767 ? Errors.VALUE : out;
  },
});
defineScalar('DOLLAR', 1, 2, ([x, d]) => {
  const v = n(x);
  const dec = n(d, 2);
  const e = err(v, dec);
  if (e) return e;
  return formatFixed(v as number, dec as number, true, true);
});
defineScalar('FIXED', 1, 3, ([x, d, noCommas]) => {
  const v = n(x);
  const dec = n(d, 2);
  const nc = noCommas === undefined ? false : toBoolean(noCommas);
  const e = err(v, dec, nc);
  if (e) return e;
  return formatFixed(v as number, dec as number, !nc, false);
});

function formatFixed(v: number, decimals: number, commas: boolean, currency: boolean): string {
  const d = Math.trunc(decimals);
  const r = roundHalfAway(v, d);
  const abs = Math.abs(r);
  let s = abs.toFixed(Math.max(0, Math.min(d, 99)));
  if (commas) {
    const [ip, fp] = s.split('.');
    s = ip.replace(/\B(?=(\d{3})+(?!\d))/g, ',') + (fp !== undefined ? `.${fp}` : '');
  }
  if (currency) return r < 0 ? `($${s})` : `$${s}`;
  return r < 0 ? `-${s}` : s;
}

defineScalar('EXACT', 2, 2, ([a, b]) => {
  const x = str(a);
  const y = str(b);
  const e = err(x, y);
  return e ?? x === y;
});

function findImpl(caseSensitive: boolean) {
  return ([needle, hay, start]: Scalar[]) => {
    const f = str(needle);
    const h = str(hay);
    const s = n(start, 1);
    const e = err(f, h, s);
    if (e) return e;
    const from = Math.trunc(s as number);
    if (from < 1 || from > (h as string).length + 1) return Errors.VALUE;
    if (caseSensitive) {
      const idx = (h as string).indexOf(f as string, from - 1);
      return idx < 0 ? Errors.VALUE : idx + 1;
    }
    if ((f as string) === '') return from;
    const re = wildcardToRegExp(f as string, false);
    const sub = (h as string).slice(from - 1);
    const m = re.exec(sub);
    return m ? m.index + from : Errors.VALUE;
  };
}
defineScalar('FIND', 2, 3, findImpl(true));
defineScalar('FINDB', 2, 3, findImpl(true));
defineScalar('SEARCH', 2, 3, findImpl(false));
defineScalar('SEARCHB', 2, 3, findImpl(false));

defineScalar('LEFT', 1, 2, ([t, k]) => {
  const s = str(t);
  const c = n(k, 1);
  const e = err(s, c);
  if (e) return e;
  if ((c as number) < 0) return Errors.VALUE;
  return [...(s as string)].slice(0, Math.trunc(c as number)).join('');
});
defineScalar('RIGHT', 1, 2, ([t, k]) => {
  const s = str(t);
  const c = n(k, 1);
  const e = err(s, c);
  if (e) return e;
  if ((c as number) < 0) return Errors.VALUE;
  const chars = [...(s as string)];
  const count = Math.trunc(c as number);
  return count === 0 ? '' : chars.slice(Math.max(0, chars.length - count)).join('');
});
defineScalar('MID', 3, 3, ([t, st, k]) => {
  const s = str(t);
  const a = n(st);
  const c = n(k);
  const e = err(s, a, c);
  if (e) return e;
  if ((a as number) < 1 || (c as number) < 0) return Errors.VALUE;
  return [...(s as string)].slice(Math.trunc(a as number) - 1, Math.trunc(a as number) - 1 + Math.trunc(c as number)).join('');
});
defineScalar('LEN', 1, 1, ([t]) => {
  const s = str(t);
  return s instanceof CellError ? s : (s as string).length;
});
defineScalar('LOWER', 1, 1, ([t]) => {
  const s = str(t);
  return s instanceof CellError ? s : s.toLowerCase();
});
defineScalar('UPPER', 1, 1, ([t]) => {
  const s = str(t);
  return s instanceof CellError ? s : s.toUpperCase();
});
defineScalar('PROPER', 1, 1, ([t]) => {
  const s = str(t);
  if (s instanceof CellError) return s;
  return s.toLowerCase().replace(/(^|[^\p{L}\p{N}'])(\p{L})/gu, (_m, pre: string, ch: string) => pre + ch.toUpperCase())
    .replace(/(\p{L})'(\p{L})/gu, (_m, a: string, b2: string) => `${a}'${b2.toLowerCase()}`);
});
defineScalar('REPLACE', 4, 4, ([t, st, k, nt]) => {
  const s = str(t);
  const a = n(st);
  const c = n(k);
  const r = str(nt);
  const e = err(s, a, c, r);
  if (e) return e;
  if ((a as number) < 1 || (c as number) < 0) return Errors.VALUE;
  const chars = [...(s as string)];
  const i = Math.trunc(a as number) - 1;
  return chars.slice(0, i).join('') + r + chars.slice(i + Math.trunc(c as number)).join('');
});
defineScalar('REPT', 2, 2, ([t, k]) => {
  const s = str(t);
  const c = n(k);
  const e = err(s, c);
  if (e) return e;
  const times = Math.trunc(c as number);
  if (times < 0 || (s as string).length * times > 32767) return Errors.VALUE;
  return (s as string).repeat(times);
});
defineScalar('SUBSTITUTE', 3, 4, ([t, oldT, newT, inst]) => {
  const s = str(t);
  const o = str(oldT);
  const r = str(newT);
  const e = err(s, o, r);
  if (e) return e;
  if (!(o as string)) return s as string;
  if (inst === undefined) return (s as string).split(o as string).join(r as string);
  const k = toNumber(inst);
  if (k instanceof CellError) return k;
  if (k < 1) return Errors.VALUE;
  let idx = -1;
  for (let i = 0; i < Math.trunc(k); i++) {
    idx = (s as string).indexOf(o as string, idx + 1);
    if (idx < 0) return s as string;
  }
  return (s as string).slice(0, idx) + r + (s as string).slice(idx + (o as string).length);
});
defineScalar('T', 1, 1, ([v]) => (typeof v === 'string' ? v : v instanceof CellError ? v : ''));
defineScalar('TRIM', 1, 1, ([t]) => {
  const s = str(t);
  return s instanceof CellError ? s : s.replace(/ +/g, ' ').replace(/^ | $/g, '');
});
defineScalar('VALUE', 1, 1, ([t]) => {
  if (typeof t === 'number') return t;
  if (t === null) return 0;
  if (typeof t === 'boolean') return Errors.VALUE;
  if (t instanceof CellError) return t;
  const v = textToNumber(t);
  return v === null ? Errors.VALUE : v;
});
defineScalar('NUMBERVALUE', 1, 3, ([t, dec, grp]) => {
  const s = str(t);
  const d = dec === undefined ? '.' : str(dec);
  const g = grp === undefined ? ',' : str(grp);
  const e = err(s, d, g);
  if (e) return e;
  let x = (s as string).replace(/\s+/g, '');
  if (!x) return 0;
  const dc = (d as string)[0] ?? '.';
  const gc = (g as string)[0] ?? ',';
  if (dc === gc) return Errors.VALUE;
  const decIdx = x.indexOf(dc);
  if (decIdx >= 0 && x.indexOf(gc, decIdx) >= 0) return Errors.VALUE;
  x = x.split(gc).join('');
  if (dc !== '.') x = x.replace(dc, '.');
  let pct = 0;
  while (x.endsWith('%')) {
    pct++;
    x = x.slice(0, -1);
  }
  const v = Number(x);
  if (!Number.isFinite(v) || x === '') return Errors.VALUE;
  return v / Math.pow(100, pct);
});
defineScalar('TEXT', 2, 2, ([v, f]) => {
  if (v instanceof CellError) return v;
  const pattern = str(f);
  if (pattern instanceof CellError) return pattern;
  return formatWithPattern(v, pattern);
});
defineScalar('VALUETOTEXT', 1, 2, ([v, fmt]) => {
  if (v instanceof CellError) return v.code;
  const strict = fmt !== undefined && toNumber(fmt) === 1;
  if (typeof v === 'string') return strict ? `"${v.replace(/"/g, '""')}"` : v;
  return toText(v) as string;
});
define('ARRAYTOTEXT', {
  min: 1,
  max: 2,
  impl: ([v, fmt], ctx) => {
    const strict = fmt != null && toNumber(toScalar(fmt, ctx)) === 1;
    const arr = toArray(v, ctx);
    const cell = (x: Scalar) => (x instanceof CellError ? x.code : typeof x === 'string' && strict ? `"${x.replace(/"/g, '""')}"` : (toText(x) as string));
    if (!strict) return arr.data.flat().map(cell).join(', ');
    return `{${arr.data.map((row) => row.map(cell).join(',')).join(';')}}`;
  },
});

// TEXTBEFORE / TEXTAFTER / TEXTSPLIT

function findDelimiter(s: string, delims: string[], instance: number, caseInsensitive: boolean, fromEnd: boolean): { start: number; len: number } | null {
  const hay = caseInsensitive ? s.toLowerCase() : s;
  const ds = delims.map((d) => (caseInsensitive ? d.toLowerCase() : d));
  const matches: { start: number; len: number }[] = [];
  for (let i = 0; i <= hay.length; ) {
    let found: { start: number; len: number } | null = null;
    for (const d of ds) {
      if (d === '') continue;
      if (hay.startsWith(d, i) && (!found || d.length > found.len)) found = { start: i, len: d.length };
    }
    if (found) {
      matches.push(found);
      i += found.len;
    } else i++;
  }
  if (ds.every((d) => d === '')) return { start: fromEnd ? s.length : 0, len: 0 };
  if (instance > 0) return matches[instance - 1] ?? (fromEnd && instance === matches.length + 1 ? { start: s.length, len: 0 } : null);
  return matches[matches.length + instance] ?? (fromEnd && -instance === matches.length + 1 ? { start: 0, len: 0 } : null);
}

function textAround(after: boolean) {
  return (args: Value[], ctx: FnContext): Value => {
    const s = toText(toScalar(args[0], ctx));
    if (s instanceof CellError) return s;
    const delims = toArray(args[1], ctx).data.flat().map((d) => toText(d));
    const ed = delims.find((d) => d instanceof CellError);
    if (ed) return ed as CellError;
    const inst = args[2] == null ? 1 : toNumber(toScalar(args[2], ctx));
    const mode = args[3] == null ? 0 : toNumber(toScalar(args[3], ctx));
    const matchEnd = args[4] == null ? 0 : toNumber(toScalar(args[4], ctx));
    const e = err(inst, mode, matchEnd);
    if (e) return e;
    const k = Math.trunc(inst as number);
    if (k === 0 || Math.abs(k) > s.length + 1) return Errors.VALUE;
    const hit = findDelimiter(s, delims as string[], k, (mode as number) === 1, (matchEnd as number) === 1);
    if (!hit) return args[5] != null ? toScalar(args[5], ctx) : Errors.NA;
    return after ? s.slice(hit.start + hit.len) : s.slice(0, hit.start);
  };
}
define('TEXTBEFORE', { min: 2, max: 6, impl: textAround(false) });
define('TEXTAFTER', { min: 2, max: 6, impl: textAround(true) });
define('TEXTSPLIT', {
  min: 2,
  max: 6,
  impl: (args, ctx) => {
    const s = toText(toScalar(args[0], ctx));
    if (s instanceof CellError) return s;
    const colDelims = args[1] == null ? [] : (toArray(args[1], ctx).data.flat().map((d) => toText(d)) as string[]);
    const rowDelims = args[2] == null ? [] : (toArray(args[2], ctx).data.flat().map((d) => toText(d)) as string[]);
    const ignoreEmpty = args[3] == null ? false : toBoolean(toScalar(args[3], ctx)) === true;
    const ci = args[4] != null && toNumber(toScalar(args[4], ctx)) === 1;
    const pad = args[5] == null ? Errors.NA : toScalar(args[5], ctx);
    const split = (text: string, delims: string[]): string[] => {
      const ds = delims.filter((d) => d !== '');
      if (!ds.length) return [text];
      const re = new RegExp(ds.map((d) => d.replace(/[.*+?^${}()|[\]\\]/g, '\\$&')).join('|'), ci ? 'gi' : 'g');
      let parts = text.split(re);
      if (ignoreEmpty) parts = parts.filter((p) => p !== '');
      return parts;
    };
    const rows = split(s, rowDelims).map((r) => split(r, colDelims));
    const width = Math.max(1, ...rows.map((r) => r.length));
    return new ArrayValue(rows.map((r) => Array.from({ length: width }, (_, i) => (i < r.length ? r[i] : pad))));
  },
});

// Regular expressions (Excel 365)
function regexFrom(pattern: string, flags: string): RegExp | CellError {
  try {
    return new RegExp(pattern, flags);
  } catch {
    return Errors.VALUE;
  }
}
defineScalar('REGEXTEST', 2, 3, ([t, p, cs]) => {
  const s = str(t);
  const pat = str(p);
  const e = err(s, pat);
  if (e) return e;
  const re = regexFrom(pat as string, cs !== undefined && toNumber(cs) === 1 ? 'i' : '');
  return re instanceof CellError ? re : re.test(s as string);
});
define('REGEXEXTRACT', {
  min: 2,
  max: 4,
  impl: (args, ctx) => {
    const s = toText(toScalar(args[0], ctx));
    const pat = toText(toScalar(args[1], ctx));
    const e = err(s, pat);
    if (e) return e;
    const mode = args[2] == null ? 0 : toNumber(toScalar(args[2], ctx));
    const ci = args[3] != null && toNumber(toScalar(args[3], ctx)) === 1;
    const re = regexFrom(pat as string, ci ? 'gi' : 'g');
    if (re instanceof CellError) return re;
    const all = [...(s as string).matchAll(re)];
    if (!all.length) return Errors.NA;
    if (mode === 1) return ArrayValue.column(all.map((m) => m[0]));
    if (mode === 2) return ArrayValue.row(all[0].slice(1).map((g) => g ?? ''));
    return all[0][0];
  },
});
defineScalar('REGEXREPLACE', 3, 5, ([t, p, r, occ, cs]) => {
  const s = str(t);
  const pat = str(p);
  const rep = str(r);
  const e = err(s, pat, rep);
  if (e) return e;
  const k = occ === undefined ? 0 : toNumber(occ);
  if (k instanceof CellError) return k;
  const re = regexFrom(pat as string, `g${cs !== undefined && toNumber(cs) === 1 ? 'i' : ''}`);
  if (re instanceof CellError) return re;
  if (k === 0) return (s as string).replace(re, rep as string);
  const matches = [...(s as string).matchAll(re)];
  const idx = k > 0 ? k - 1 : matches.length + k;
  const m = matches[idx];
  if (!m) return s as string;
  return (s as string).slice(0, m.index) + (m[0].replace(new RegExp(re.source, re.flags.replace('g', '')), rep as string)) + (s as string).slice(m.index! + m[0].length);
});

defineScalar('ASC', 1, 1, ([t]) => str(t));
defineScalar('DBCS', 1, 1, ([t]) => str(t));
defineScalar('PHONETIC', 1, 1, ([t]) => str(t));
defineScalar('HYPERLINK', 1, 2, ([link, name]) => (name === undefined || name === null ? str(link) : name));

// Byte variants behave like the character variants for single-byte languages.
alias('LEFTB', 'LEFT');
alias('RIGHTB', 'RIGHT');
alias('MIDB', 'MID');
alias('LENB', 'LEN');
alias('REPLACEB', 'REPLACE');
