/** A1 notation, cell addresses and rectangular ranges. Rows and columns are 0-based. */

export const MAX_ROWS = 1048576;
export const MAX_COLS = 16384;

export interface Address {
  r: number;
  c: number;
}

/** An inclusive rectangle, always normalised so that r0 <= r1 and c0 <= c1. */
export interface Range {
  r0: number;
  c0: number;
  r1: number;
  c1: number;
}

export function colName(c: number): string {
  let s = '';
  let n = c + 1;
  while (n > 0) {
    const rem = (n - 1) % 26;
    s = String.fromCharCode(65 + rem) + s;
    n = Math.floor((n - 1) / 26);
  }
  return s;
}

/** 'A' -> 0, 'AA' -> 26. Returns -1 for invalid input. */
export function colIndex(name: string): number {
  let n = 0;
  for (let i = 0; i < name.length; i++) {
    const ch = name.charCodeAt(i) & ~32; // upper-case
    if (ch < 65 || ch > 90) return -1;
    n = n * 26 + (ch - 64);
    if (n > MAX_COLS) return -1;
  }
  return n - 1;
}

export function cellName(r: number, c: number, absR = false, absC = false): string {
  return `${absC ? '$' : ''}${colName(c)}${absR ? '$' : ''}${r + 1}`;
}

export function addressName(a: Address): string {
  return cellName(a.r, a.c);
}

const CELL_RE = /^(\$?)([A-Za-z]{1,3})(\$?)([1-9][0-9]{0,6})$/;

export interface CellRef extends Address {
  absR: boolean;
  absC: boolean;
}

export function parseCellRef(text: string): CellRef | null {
  const m = CELL_RE.exec(text.trim());
  if (!m) return null;
  const c = colIndex(m[2]);
  const r = Number(m[4]) - 1;
  if (c < 0 || r >= MAX_ROWS) return null;
  return { r, c, absC: m[1] === '$', absR: m[3] === '$' };
}

export function makeRange(r0: number, c0: number, r1 = r0, c1 = c0): Range {
  return { r0: Math.min(r0, r1), c0: Math.min(c0, c1), r1: Math.max(r0, r1), c1: Math.max(c0, c1) };
}

export function cellRange(a: Address): Range {
  return { r0: a.r, c0: a.c, r1: a.r, c1: a.c };
}

/** Parses 'A1', 'A1:B5', 'A:C', '2:4' (also with $ signs). */
export function parseRange(text: string): Range | null {
  const t = text.trim();
  const idx = t.indexOf(':');
  if (idx < 0) {
    const ref = parseCellRef(t);
    return ref ? cellRange(ref) : null;
  }
  const a = t.slice(0, idx).replace(/\$/g, '');
  const b = t.slice(idx + 1).replace(/\$/g, '');
  if (/^[A-Za-z]{1,3}$/.test(a) && /^[A-Za-z]{1,3}$/.test(b)) {
    const c0 = colIndex(a);
    const c1 = colIndex(b);
    if (c0 < 0 || c1 < 0) return null;
    return makeRange(0, c0, MAX_ROWS - 1, c1);
  }
  if (/^\d+$/.test(a) && /^\d+$/.test(b)) {
    const r0 = Number(a) - 1;
    const r1 = Number(b) - 1;
    if (r0 < 0 || r1 < 0 || r0 >= MAX_ROWS || r1 >= MAX_ROWS) return null;
    return makeRange(r0, 0, r1, MAX_COLS - 1);
  }
  const ra = parseCellRef(a);
  const rb = parseCellRef(b);
  if (!ra || !rb) return null;
  return makeRange(ra.r, ra.c, rb.r, rb.c);
}

export function isFullColumns(r: Range): boolean {
  return r.r0 === 0 && r.r1 === MAX_ROWS - 1;
}

export function isFullRows(r: Range): boolean {
  return r.c0 === 0 && r.c1 === MAX_COLS - 1;
}

export function rangeName(r: Range): string {
  if (isFullColumns(r) && !isFullRows(r)) return `${colName(r.c0)}:${colName(r.c1)}`;
  if (isFullRows(r) && !isFullColumns(r)) return `${r.r0 + 1}:${r.r1 + 1}`;
  const a = cellName(r.r0, r.c0);
  if (r.r0 === r.r1 && r.c0 === r.c1) return a;
  return `${a}:${cellName(r.r1, r.c1)}`;
}

export function isSingleCell(r: Range): boolean {
  return r.r0 === r.r1 && r.c0 === r.c1;
}

export function rangeContains(r: Range, row: number, col: number): boolean {
  return row >= r.r0 && row <= r.r1 && col >= r.c0 && col <= r.c1;
}

export function rangeContainsRange(outer: Range, inner: Range): boolean {
  return inner.r0 >= outer.r0 && inner.r1 <= outer.r1 && inner.c0 >= outer.c0 && inner.c1 <= outer.c1;
}

export function rangesIntersect(a: Range, b: Range): boolean {
  return a.r0 <= b.r1 && b.r0 <= a.r1 && a.c0 <= b.c1 && b.c0 <= a.c1;
}

export function intersect(a: Range, b: Range): Range | null {
  if (!rangesIntersect(a, b)) return null;
  return { r0: Math.max(a.r0, b.r0), c0: Math.max(a.c0, b.c0), r1: Math.min(a.r1, b.r1), c1: Math.min(a.c1, b.c1) };
}

export function union(a: Range, b: Range): Range {
  return { r0: Math.min(a.r0, b.r0), c0: Math.min(a.c0, b.c0), r1: Math.max(a.r1, b.r1), c1: Math.max(a.c1, b.c1) };
}

export function rangeEquals(a: Range, b: Range): boolean {
  return a.r0 === b.r0 && a.c0 === b.c0 && a.r1 === b.r1 && a.c1 === b.c1;
}

export function rangeRows(r: Range): number {
  return r.r1 - r.r0 + 1;
}

export function rangeCols(r: Range): number {
  return r.c1 - r.c0 + 1;
}

/** Parses an Excel `sqref` (space separated list of ranges). */
export function parseSqref(text: string): Range[] {
  const out: Range[] = [];
  for (const part of text.split(/\s+/)) {
    if (!part) continue;
    const r = parseRange(part);
    if (r) out.push(r);
  }
  return out;
}

export function sqrefString(ranges: Range[]): string {
  return ranges.map(rangeName).join(' ');
}

const BARE_SHEET_RE = /^[A-Za-z_¡-￿][A-Za-z0-9_.¡-￿]*$/;

/** Sheet name as written in a formula, quoted when needed. */
export function quoteSheetName(name: string): string {
  if (BARE_SHEET_RE.test(name) && !parseCellRef(name) && !/^(TRUE|FALSE)$/i.test(name) && !/^R\d*C\d*$/i.test(name)) {
    return name;
  }
  return `'${name.replace(/'/g, "''")}'`;
}

/** Splits "'My Sheet'!A1:B2" or "Sheet1!A1" into sheet name and range text. */
export function splitSheetRef(text: string): { sheet?: string; ref: string } {
  const t = text.trim();
  if (t.startsWith("'")) {
    let i = 1;
    let name = '';
    while (i < t.length) {
      if (t[i] === "'") {
        if (t[i + 1] === "'") {
          name += "'";
          i += 2;
          continue;
        }
        break;
      }
      name += t[i++];
    }
    if (t[i + 1] === '!') return { sheet: name, ref: t.slice(i + 2) };
    return { ref: t };
  }
  const bang = t.lastIndexOf('!');
  if (bang > 0) return { sheet: t.slice(0, bang), ref: t.slice(bang + 1) };
  return { ref: t };
}

/** Numeric key for a cell inside one sheet. */
export function cellKey(r: number, c: number): number {
  return r * MAX_COLS + c;
}

export function keyRow(key: number): number {
  return Math.floor(key / MAX_COLS);
}

export function keyCol(key: number): number {
  return key % MAX_COLS;
}

export function clampRange(r: Range): Range {
  return {
    r0: Math.max(0, Math.min(MAX_ROWS - 1, r.r0)),
    c0: Math.max(0, Math.min(MAX_COLS - 1, r.c0)),
    r1: Math.max(0, Math.min(MAX_ROWS - 1, r.r1)),
    c1: Math.max(0, Math.min(MAX_COLS - 1, r.c1)),
  };
}
