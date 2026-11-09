/**
 * Excel formula tokenizer. Tokens keep their source text and position so formulas can be
 * rewritten (reference shifting, sheet renames) without reformatting them.
 */
import { colIndex, MAX_COLS, MAX_ROWS } from '../refs';

export type TokenType =
  | 'number'
  | 'string'
  | 'bool'
  | 'error'
  | 'ref'
  | 'name'
  | 'func'
  | 'structured'
  | 'op'
  | 'open'
  | 'close'
  | 'brace-open'
  | 'brace-close'
  | 'comma'
  | 'semicolon'
  | 'colon'
  | 'spill'
  | 'at'
  | 'ws'
  | 'unknown';

export interface RefParts {
  /** Sheet name (unquoted). */
  sheet?: string;
  /** Last sheet of a 3D reference (Sheet1:Sheet3!A1). */
  sheet2?: string;
  /** External workbook index or name ([1]Sheet1!A1). */
  book?: string;
  kind: 'cell' | 'range' | 'cols' | 'rows' | 'error';
  r0: number;
  c0: number;
  r1: number;
  c1: number;
  absR0: boolean;
  absC0: boolean;
  absR1: boolean;
  absC1: boolean;
}

export interface Token {
  type: TokenType;
  text: string;
  start: number;
  end: number;
  /** Parsed reference for 'ref' tokens. */
  ref?: RefParts;
  /** Sheet qualifier of a sheet-scoped name token (Sheet1!MyName). */
  sheet?: string;
  /** Normalised value for literals and names. */
  value?: string | number | boolean;
}

const ERRORS = ['#NULL!', '#DIV/0!', '#VALUE!', '#REF!', '#NAME?', '#NUM!', '#N/A', '#GETTING_DATA', '#SPILL!', '#CALC!'];

function isIdentStart(ch: string): boolean {
  return /[A-Za-z_\\¡-￿]/.test(ch);
}

function isIdentChar(ch: string): boolean {
  return /[A-Za-z0-9_.?\\¡-￿]/.test(ch);
}

function isSheetChar(ch: string): boolean {
  return /[A-Za-z0-9_.¡-￿]/.test(ch);
}

const CELL_PART = /^(\$?)([A-Za-z]{1,3})(\$?)(\d{1,7})/;
const COL_PART = /^(\$?)([A-Za-z]{1,3})(?![A-Za-z0-9_(.])/;
const ROW_PART = /^(\$?)(\d{1,7})(?![0-9.A-Za-z_(])/;

interface CellPart {
  len: number;
  r: number;
  c: number;
  absR: boolean;
  absC: boolean;
}

function matchCell(src: string): CellPart | null {
  const m = CELL_PART.exec(src);
  if (!m) return null;
  const next = src[m[0].length];
  if (next !== undefined && /[A-Za-z0-9_(.]/.test(next)) return null;
  const c = colIndex(m[2]);
  const r = Number(m[4]) - 1;
  if (c < 0 || c >= MAX_COLS || r < 0 || r >= MAX_ROWS) return null;
  return { len: m[0].length, r, c, absC: m[1] === '$', absR: m[3] === '$' };
}

/** Matches a reference body (no sheet prefix) at the start of `src`. */
function matchRefBody(src: string): { len: number; ref: Omit<RefParts, 'sheet' | 'sheet2' | 'book'> } | null {
  // Cell or cell range.
  const a = matchCell(src);
  if (a) {
    if (src[a.len] === ':') {
      const b = matchCell(src.slice(a.len + 1));
      if (b) {
        return {
          len: a.len + 1 + b.len,
          ref: { kind: 'range', r0: a.r, c0: a.c, r1: b.r, c1: b.c, absR0: a.absR, absC0: a.absC, absR1: b.absR, absC1: b.absC },
        };
      }
    }
    return { len: a.len, ref: { kind: 'cell', r0: a.r, c0: a.c, r1: a.r, c1: a.c, absR0: a.absR, absC0: a.absC, absR1: a.absR, absC1: a.absC } };
  }
  // Column range A:C.
  const ca = /^(\$?)([A-Za-z]{1,3}):(\$?)([A-Za-z]{1,3})(?![A-Za-z0-9_(.])/.exec(src);
  if (ca) {
    const c0 = colIndex(ca[2]);
    const c1 = colIndex(ca[4]);
    if (c0 >= 0 && c1 >= 0) {
      return {
        len: ca[0].length,
        ref: { kind: 'cols', r0: 0, c0, r1: MAX_ROWS - 1, c1, absR0: false, absC0: ca[1] === '$', absR1: false, absC1: ca[3] === '$' },
      };
    }
  }
  // Row range 1:3.
  const ra = /^(\$?)(\d{1,7}):(\$?)(\d{1,7})(?![0-9.A-Za-z_(])/.exec(src);
  if (ra) {
    const r0 = Number(ra[2]) - 1;
    const r1 = Number(ra[4]) - 1;
    if (r0 >= 0 && r1 >= 0 && r0 < MAX_ROWS && r1 < MAX_ROWS) {
      return {
        len: ra[0].length,
        ref: { kind: 'rows', r0, c0: 0, r1, c1: MAX_COLS - 1, absR0: ra[1] === '$', absC0: false, absR1: ra[3] === '$', absC1: false },
      };
    }
  }
  return null;
}

void COL_PART;
void ROW_PART;

/** Reads a sheet prefix ('My Sheet'!, Sheet1!, Sheet1:Sheet3!, [1]Sheet1!) at `i`. */
function readSheetPrefix(src: string, i: number): { end: number; sheet: string; sheet2?: string; book?: string } | null {
  let j = i;
  let book: string | undefined;
  if (src[j] === '[') {
    const close = src.indexOf(']', j);
    if (close < 0) return null;
    book = src.slice(j + 1, close);
    j = close + 1;
  }
  if (src[j] === "'") {
    let k = j + 1;
    let name = '';
    while (k < src.length) {
      if (src[k] === "'") {
        if (src[k + 1] === "'") {
          name += "'";
          k += 2;
          continue;
        }
        break;
      }
      name += src[k++];
    }
    if (src[k] !== "'" || src[k + 1] !== '!') return null;
    let sheet = name;
    let sheet2: string | undefined;
    if (!book && sheet.startsWith('[')) {
      const close = sheet.indexOf(']');
      if (close > 0) {
        book = sheet.slice(1, close);
        sheet = sheet.slice(close + 1);
      }
    }
    const colon = sheet.indexOf(':');
    if (colon > 0) {
      sheet2 = sheet.slice(colon + 1);
      sheet = sheet.slice(0, colon);
    }
    return { end: k + 2, sheet, sheet2, book };
  }
  let k = j;
  while (k < src.length && isSheetChar(src[k])) k++;
  if (k === j) return null;
  let sheet = src.slice(j, k);
  let sheet2: string | undefined;
  if (src[k] === ':' && isSheetChar(src[k + 1] ?? '')) {
    let m = k + 1;
    while (m < src.length && isSheetChar(src[m])) m++;
    if (src[m] === '!') {
      sheet2 = src.slice(k + 1, m);
      k = m;
    }
  }
  if (src[k] !== '!') return null;
  if (/^\d/.test(sheet) && !book) return null;
  return { end: k + 1, sheet, sheet2, book };
}

export function tokenize(formula: string): Token[] {
  const src = formula;
  const out: Token[] = [];
  let i = 0;
  const n = src.length;
  const push = (type: TokenType, start: number, end: number, extra: Partial<Token> = {}) => {
    out.push({ type, text: src.slice(start, end), start, end, ...(extra as object) });
  };

  const lastSignificant = (): Token | undefined => {
    for (let k = out.length - 1; k >= 0; k--) if (out[k].type !== 'ws') return out[k];
    return undefined;
  };

  while (i < n) {
    const ch = src[i];
    // Whitespace.
    if (ch === ' ' || ch === '\n' || ch === '\r' || ch === '\t') {
      let j = i + 1;
      while (j < n && /[ \n\r\t]/.test(src[j])) j++;
      push('ws', i, j);
      i = j;
      continue;
    }
    // Strings.
    if (ch === '"') {
      let j = i + 1;
      let value = '';
      while (j < n) {
        if (src[j] === '"') {
          if (src[j + 1] === '"') {
            value += '"';
            j += 2;
            continue;
          }
          break;
        }
        value += src[j++];
      }
      push('string', i, Math.min(j + 1, n), { value });
      i = Math.min(j + 1, n);
      continue;
    }
    // Errors (#REF! etc).
    if (ch === '#') {
      const prev = lastSignificant();
      const rest = src.slice(i).toUpperCase();
      const err = ERRORS.find((e) => rest.startsWith(e));
      if (err) {
        push('error', i, i + err.length, { value: err });
        i += err.length;
        continue;
      }
      if (prev && (prev.type === 'ref' || prev.type === 'name' || prev.type === 'close') && prev.end === i) {
        push('spill', i, i + 1);
        i++;
        continue;
      }
      push('unknown', i, i + 1);
      i++;
      continue;
    }
    // Sheet-qualified references or names.
    if (ch === "'" || ch === '[' || isSheetChar(ch)) {
      const prefix = readSheetPrefix(src, i);
      if (prefix) {
        const rest = src.slice(prefix.end);
        const body = matchRefBody(rest);
        if (body) {
          const end = prefix.end + body.len;
          push('ref', i, end, { ref: { ...body.ref, sheet: prefix.sheet, sheet2: prefix.sheet2, book: prefix.book } });
          i = end;
          continue;
        }
        const errRef = /^#REF!/i.exec(rest);
        if (errRef) {
          const end = prefix.end + 5;
          push('ref', i, end, {
            ref: { kind: 'error', sheet: prefix.sheet, sheet2: prefix.sheet2, book: prefix.book, r0: 0, c0: 0, r1: 0, c1: 0, absR0: false, absC0: false, absR1: false, absC1: false },
          });
          i = end;
          continue;
        }
        if (isIdentStart(rest[0] ?? '')) {
          let j = prefix.end + 1;
          while (j < n && isIdentChar(src[j])) j++;
          push('name', i, j, { sheet: prefix.sheet, value: src.slice(prefix.end, j) });
          i = j;
          continue;
        }
      }
    }
    // Unqualified references.
    if (ch === '$' || /[A-Za-z0-9]/.test(ch)) {
      const body = matchRefBody(src.slice(i));
      if (body) {
        // A bare number like "1" is not a row reference unless followed by ':'.
        push('ref', i, i + body.len, { ref: { ...body.ref } });
        i += body.len;
        continue;
      }
    }
    // Numbers.
    if (/[0-9]/.test(ch) || (ch === '.' && /[0-9]/.test(src[i + 1] ?? ''))) {
      const m = /^(\d+\.?\d*|\.\d+)([eE][+-]?\d+)?/.exec(src.slice(i));
      if (m) {
        push('number', i, i + m[0].length, { value: Number(m[0]) });
        i += m[0].length;
        continue;
      }
    }
    // Identifiers: functions, names, booleans, structured references.
    if (isIdentStart(ch)) {
      let j = i + 1;
      while (j < n && isIdentChar(src[j])) j++;
      const word = src.slice(i, j);
      if (src[j] === '(') {
        push('func', i, j, { value: word });
        i = j;
        continue;
      }
      if (src[j] === '[') {
        const end = matchBrackets(src, j);
        push('structured', i, end, { value: word });
        i = end;
        continue;
      }
      const upper = word.toUpperCase();
      if (upper === 'TRUE' || upper === 'FALSE') {
        push('bool', i, j, { value: upper === 'TRUE' });
        i = j;
        continue;
      }
      push('name', i, j, { value: word });
      i = j;
      continue;
    }
    if (ch === '[') {
      const end = matchBrackets(src, i);
      push('structured', i, end, { value: '' });
      i = end;
      continue;
    }
    // Operators and punctuation.
    const two = src.slice(i, i + 2);
    if (two === '<>' || two === '<=' || two === '>=') {
      push('op', i, i + 2, { value: two });
      i += 2;
      continue;
    }
    switch (ch) {
      case '+':
      case '-':
      case '*':
      case '/':
      case '^':
      case '&':
      case '=':
      case '<':
      case '>':
      case '%':
        push('op', i, i + 1, { value: ch });
        break;
      case '(':
        push('open', i, i + 1);
        break;
      case ')':
        push('close', i, i + 1);
        break;
      case '{':
        push('brace-open', i, i + 1);
        break;
      case '}':
        push('brace-close', i, i + 1);
        break;
      case ',':
        push('comma', i, i + 1);
        break;
      case ';':
        push('semicolon', i, i + 1);
        break;
      case ':':
        push('colon', i, i + 1);
        break;
      case '@':
        push('at', i, i + 1);
        break;
      default:
        push('unknown', i, i + 1);
    }
    i++;
  }
  return out;
}

function matchBrackets(src: string, start: number): number {
  let depth = 0;
  let i = start;
  while (i < src.length) {
    const ch = src[i];
    if (ch === "'" && src[i + 1]) {
      i += 2;
      continue;
    }
    if (ch === '[') depth++;
    else if (ch === ']') {
      depth--;
      if (depth === 0) return i + 1;
    }
    i++;
  }
  return src.length;
}

/** Formats a reference back to text (used after shifting references). */
export function refToText(ref: RefParts, quote: (sheet: string) => string): string {
  let prefix = '';
  if (ref.sheet !== undefined) {
    const sheets = ref.sheet2 !== undefined ? `${ref.sheet}:${ref.sheet2}` : ref.sheet;
    const book = ref.book !== undefined ? `[${ref.book}]` : '';
    const needsQuote = quote(sheets) !== sheets || (ref.book !== undefined && /[^A-Za-z0-9_.]/.test(sheets));
    prefix = needsQuote ? `'${book}${sheets.replace(/'/g, "''")}'!` : `${book}${sheets}!`;
    if (ref.sheet2 !== undefined) {
      const q1 = quote(ref.sheet);
      const q2 = quote(ref.sheet2);
      if (q1 !== ref.sheet || q2 !== ref.sheet2) prefix = `'${book}${ref.sheet.replace(/'/g, "''")}:${ref.sheet2.replace(/'/g, "''")}'!`;
    }
  }
  if (ref.kind === 'error') return `${prefix}#REF!`;
  const col = (c: number, abs: boolean) => `${abs ? '$' : ''}${colLetters(c)}`;
  const row = (r: number, abs: boolean) => `${abs ? '$' : ''}${r + 1}`;
  switch (ref.kind) {
    case 'cell':
      return `${prefix}${col(ref.c0, ref.absC0)}${row(ref.r0, ref.absR0)}`;
    case 'range':
      return `${prefix}${col(ref.c0, ref.absC0)}${row(ref.r0, ref.absR0)}:${col(ref.c1, ref.absC1)}${row(ref.r1, ref.absR1)}`;
    case 'cols':
      return `${prefix}${col(ref.c0, ref.absC0)}:${col(ref.c1, ref.absC1)}`;
    case 'rows':
      return `${prefix}${row(ref.r0, ref.absR0)}:${row(ref.r1, ref.absR1)}`;
  }
}

function colLetters(c: number): string {
  let s = '';
  let n = c + 1;
  while (n > 0) {
    const rem = (n - 1) % 26;
    s = String.fromCharCode(65 + rem) + s;
    n = Math.floor((n - 1) / 26);
  }
  return s;
}
