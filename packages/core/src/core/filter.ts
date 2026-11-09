/**
 * Auto filter logic: which rows of a filtered area match the column criteria, and the distinct
 * values offered in a column's filter list. Values are compared as displayed (like Excel), numbers
 * numerically for the ordering conditions.
 */
import type { AutoFilter, ColumnFilter, FilterCondition, Sheet } from './model/sheet';
import type { Scalar } from './values';

export type DisplayText = (r: number, c: number) => string;
export type CellValue = (r: number, c: number) => Scalar;

function compareCondition(cond: FilterCondition, text: string, value: Scalar): boolean {
  const t = text.toLowerCase();
  const v = cond.value.toLowerCase();
  const num = Number(cond.value);
  const numeric = typeof value === 'number' && cond.value.trim() !== '' && Number.isFinite(num);
  switch (cond.op) {
    case 'contains':
      return wildcard(`*${v}*`).test(t);
    case 'notContains':
      return !wildcard(`*${v}*`).test(t);
    case 'beginsWith':
      return wildcard(`${v}*`).test(t);
    case 'endsWith':
      return wildcard(`*${v}`).test(t);
    case 'equal':
      return numeric ? value === num : wildcard(v).test(t);
    case 'notEqual':
      return numeric ? value !== num : !wildcard(v).test(t);
    case 'greaterThan':
      return numeric ? (value as number) > num : t > v;
    case 'lessThan':
      return numeric ? (value as number) < num : t < v;
    case 'greaterThanOrEqual':
      return numeric ? (value as number) >= num : t >= v;
    case 'lessThanOrEqual':
      return numeric ? (value as number) <= num : t <= v;
  }
}

const wildcardCache = new Map<string, RegExp>();

/** Excel wildcards: * any text, ? one character, ~ escapes. */
function wildcard(pattern: string): RegExp {
  let re = wildcardCache.get(pattern);
  if (re) return re;
  let src = '';
  for (let i = 0; i < pattern.length; i++) {
    const ch = pattern[i];
    if (ch === '~' && i + 1 < pattern.length) src += pattern[++i].replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
    else if (ch === '*') src += '[\\s\\S]*';
    else if (ch === '?') src += '[\\s\\S]';
    else src += ch.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
  }
  re = new RegExp(`^${src}$`, 'i');
  if (wildcardCache.size > 200) wildcardCache.clear();
  wildcardCache.set(pattern, re);
  return re;
}

export function cellMatches(filter: ColumnFilter, text: string, value: Scalar): boolean {
  if (filter.values) {
    if (text === '') return !!filter.blanks;
    const set = new Set(filter.values.map((v) => v.toLowerCase()));
    if (!set.has(text.toLowerCase())) return false;
  }
  if (filter.conditions?.length) {
    const results = filter.conditions.map((c) => compareCondition(c, text, value));
    return filter.and ? results.every(Boolean) : results.some(Boolean);
  }
  return true;
}

/** Rows (below the header) of the filter area that do not match the criteria. */
export function rowsToHide(af: AutoFilter, display: DisplayText, value: CellValue): Set<number> {
  const hidden = new Set<number>();
  if (!af.columns.size) return hidden;
  const cols = [...af.columns.entries()].filter(([c]) => c >= af.ref.c0 && c <= af.ref.c1);
  for (let r = af.ref.r0 + 1; r <= af.ref.r1; r++) {
    for (const [c, f] of cols) {
      if (!cellMatches(f, display(r, c), value(r, c))) {
        hidden.add(r);
        break;
      }
    }
  }
  return hidden;
}

/** Distinct displayed values of a column below the header, with counts, sorted. */
export function columnValues(af: AutoFilter, col: number, display: DisplayText, value: CellValue): { text: string; count: number; numeric: boolean }[] {
  const map = new Map<string, { text: string; count: number; numeric: boolean; sort: number | string }>();
  for (let r = af.ref.r0 + 1; r <= af.ref.r1; r++) {
    const text = display(r, col);
    const key = text.toLowerCase();
    const v = value(r, col);
    const hit = map.get(key);
    if (hit) hit.count++;
    else map.set(key, { text, count: 1, numeric: typeof v === 'number', sort: typeof v === 'number' ? v : key });
  }
  return [...map.values()]
    .sort((a, b) => {
      if (a.text === '') return 1;
      if (b.text === '') return -1;
      if (typeof a.sort === 'number' && typeof b.sort === 'number') return a.sort - b.sort;
      if (typeof a.sort === 'number') return -1;
      if (typeof b.sort === 'number') return 1;
      return String(a.sort).localeCompare(String(b.sort));
    })
    .map(({ text, count, numeric }) => ({ text, count, numeric }));
}

/**
 * The area an auto filter should cover around the active cell: the contiguous block of data (like
 * Excel's current region), or the selection when several cells are selected.
 */
export function filterRegion(sheet: Sheet, sel: { r0: number; c0: number; r1: number; c1: number }, active: { r: number; c: number }): { r0: number; c0: number; r1: number; c1: number } | null {
  if (sel.r1 > sel.r0 && sel.r1 - sel.r0 < 1_000_000 && sel.c1 - sel.c0 < 16384) {
    const used = sheet.usedRange(false);
    if (!used) return null;
    return { r0: sel.r0, c0: sel.c0, r1: Math.min(sel.r1, used.r1), c1: Math.min(sel.c1, used.c1) };
  }
  const filled = (r: number, c: number) => {
    const v = sheet.getCell(r, c)?.v;
    return v !== null && v !== undefined && v !== '';
  };
  let r0 = active.r;
  let r1 = active.r;
  let c0 = active.c;
  let c1 = active.c;
  // Grow the box while any neighbouring row or column has data.
  for (let changed = true, guard = 0; changed && guard < 10000; guard++) {
    changed = false;
    const rowHas = (r: number) => {
      for (let c = Math.max(0, c0 - 1); c <= c1 + 1; c++) if (filled(r, c)) return true;
      return false;
    };
    const colHas = (c: number) => {
      for (let r = Math.max(0, r0 - 1); r <= r1 + 1; r++) if (filled(r, c)) return true;
      return false;
    };
    if (r0 > 0 && rowHas(r0 - 1)) { r0--; changed = true; }
    if (rowHas(r1 + 1)) { r1++; changed = true; }
    if (c0 > 0 && colHas(c0 - 1)) { c0--; changed = true; }
    if (colHas(c1 + 1)) { c1++; changed = true; }
  }
  if (r0 === r1 && c0 === c1 && !filled(r0, c0)) return null;
  return { r0, c0, r1, c1 };
}
