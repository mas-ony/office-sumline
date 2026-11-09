/**
 * Token-based formula rewriting: reference shifting on row/column insertion and deletion,
 * relative translation for copy/fill, sheet renames and cut-and-paste moves. Formatting of the
 * original formula (spaces, case of function names) is kept.
 */
import { MAX_COLS, MAX_ROWS, quoteSheetName, type Range } from '../refs';
import { refToText, tokenize, type RefParts, type Token } from './tokenizer';

function rebuild(tokens: Token[], replace: (t: Token) => string | null): string {
  let out = '';
  for (const t of tokens) {
    const r = replace(t);
    out += r === null ? t.text : r;
  }
  return out;
}

function sameSheet(a: string | undefined, b: string): boolean {
  return a !== undefined && a.toLowerCase() === b.toLowerCase();
}

/** Whether a reference token points at `sheet`, given the sheet the formula lives on. */
function targets(ref: RefParts, formulaSheet: string, sheet: string): boolean {
  if (ref.book !== undefined) return false;
  if (ref.sheet === undefined) return sameSheet(formulaSheet, sheet);
  if (ref.sheet2 !== undefined) return false;
  return sameSheet(ref.sheet, sheet);
}

function errorRef(ref: RefParts): string {
  return refToText({ ...ref, kind: 'error' }, quoteSheetName);
}

/**
 * Adjusts the references of a formula for an insertion (delta > 0 at index `from`) or a deletion
 * (delta < 0 of the lines `from + delta .. from - 1`) of rows or columns in `sheet`.
 */
export function shiftFormula(formula: string, formulaSheet: string, sheet: string, axis: 'row' | 'col', from: number, delta: number): string {
  const tokens = tokenize(formula);
  let changed = false;
  const out = rebuild(tokens, (t) => {
    if (t.type !== 'ref' || !t.ref || t.ref.kind === 'error' || !targets(t.ref, formulaSheet, sheet)) return null;
    const ref = t.ref;
    const isRow = axis === 'row';
    if ((isRow && ref.kind === 'cols') || (!isRow && ref.kind === 'rows')) return null;
    const lo = isRow ? ref.r0 : ref.c0;
    const hi = isRow ? ref.r1 : ref.c1;
    const a = Math.min(lo, hi);
    const b = Math.max(lo, hi);
    const max = (isRow ? MAX_ROWS : MAX_COLS) - 1;
    let na = a;
    let nb = b;
    if (delta > 0) {
      if (a >= from) na = a + delta;
      if (b >= from) nb = b + delta;
      if (nb > max || na > max) {
        changed = true;
        return errorRef(ref);
      }
    } else {
      const delStart = from + delta;
      const delEnd = from - 1;
      if (a >= delStart && b <= delEnd) {
        changed = true;
        return errorRef(ref);
      }
      na = a > delEnd ? a + delta : a >= delStart ? delStart : a;
      nb = b > delEnd ? b + delta : b >= delStart ? delStart - 1 : b;
    }
    if (na === a && nb === b) return null;
    changed = true;
    const next: RefParts = { ...ref };
    if (isRow) {
      next.r0 = lo <= hi ? na : nb;
      next.r1 = lo <= hi ? nb : na;
    } else {
      next.c0 = lo <= hi ? na : nb;
      next.c1 = lo <= hi ? nb : na;
    }
    if (next.kind === 'range' && next.r0 === next.r1 && next.c0 === next.c1 && ref.kind === 'cell') next.kind = 'cell';
    return refToText(next, quoteSheetName);
  });
  return changed ? out : formula;
}

/**
 * Translates relative references by (dr, dc), as when a formula is copied or filled.
 * References that leave the sheet become #REF!.
 */
export function translateFormula(formula: string, dr: number, dc: number): string {
  if (!dr && !dc) return formula;
  const tokens = tokenize(formula);
  return rebuild(tokens, (t) => {
    if (t.type !== 'ref' || !t.ref || t.ref.kind === 'error' || t.ref.book !== undefined) return null;
    const ref = t.ref;
    const next: RefParts = { ...ref };
    if (ref.kind !== 'cols') {
      if (!ref.absR0) next.r0 = ref.r0 + dr;
      if (!ref.absR1) next.r1 = ref.r1 + dr;
    }
    if (ref.kind !== 'rows') {
      if (!ref.absC0) next.c0 = ref.c0 + dc;
      if (!ref.absC1) next.c1 = ref.c1 + dc;
    }
    if (next.r0 < 0 || next.r1 < 0 || next.c0 < 0 || next.c1 < 0 || next.r0 >= MAX_ROWS || next.r1 >= MAX_ROWS || next.c0 >= MAX_COLS || next.c1 >= MAX_COLS) {
      return errorRef(ref);
    }
    if (next.r0 === ref.r0 && next.r1 === ref.r1 && next.c0 === ref.c0 && next.c1 === ref.c1) return null;
    return refToText(next, quoteSheetName);
  });
}

/** Rewrites sheet names in references and qualified names after a rename. */
export function renameSheetInFormula(formula: string, oldName: string, newName: string): string {
  const tokens = tokenize(formula);
  let changed = false;
  const out = rebuild(tokens, (t) => {
    if (t.type === 'ref' && t.ref && t.ref.book === undefined) {
      const ref = t.ref;
      let hit = false;
      const next = { ...ref };
      if (sameSheet(ref.sheet, oldName)) {
        next.sheet = newName;
        hit = true;
      }
      if (sameSheet(ref.sheet2, oldName)) {
        next.sheet2 = newName;
        hit = true;
      }
      if (!hit) return null;
      changed = true;
      return refToText(next, quoteSheetName);
    }
    if (t.type === 'name' && sameSheet(t.sheet, oldName)) {
      changed = true;
      return `${quoteSheetName(newName)}!${String(t.value)}`;
    }
    return null;
  });
  return changed ? out : formula;
}

/** References to a deleted sheet become #REF!. */
export function deleteSheetInFormula(formula: string, name: string): string {
  const tokens = tokenize(formula);
  let changed = false;
  const out = rebuild(tokens, (t) => {
    if (t.type === 'ref' && t.ref && t.ref.book === undefined && (sameSheet(t.ref.sheet, name) || sameSheet(t.ref.sheet2, name))) {
      changed = true;
      return '#REF!';
    }
    if (t.type === 'name' && sameSheet(t.sheet, name)) {
      changed = true;
      return '#REF!';
    }
    return null;
  });
  return changed ? out : formula;
}

/**
 * After cutting `src` (on sheet `srcSheet`) and pasting it at an offset (and maybe another sheet),
 * references that point entirely inside the moved block follow it.
 */
export function moveRefsInFormula(
  formula: string,
  formulaSheet: string,
  srcSheet: string,
  src: Range,
  dstSheet: string,
  dr: number,
  dc: number,
): string {
  const tokens = tokenize(formula);
  let changed = false;
  const out = rebuild(tokens, (t) => {
    if (t.type !== 'ref' || !t.ref || t.ref.kind === 'error' || !targets(t.ref, formulaSheet, srcSheet)) return null;
    const ref = t.ref;
    if (ref.kind === 'cols' || ref.kind === 'rows') return null;
    const r0 = Math.min(ref.r0, ref.r1);
    const r1 = Math.max(ref.r0, ref.r1);
    const c0 = Math.min(ref.c0, ref.c1);
    const c1 = Math.max(ref.c0, ref.c1);
    if (r0 < src.r0 || r1 > src.r1 || c0 < src.c0 || c1 > src.c1) return null;
    changed = true;
    const next: RefParts = { ...ref, r0: ref.r0 + dr, r1: ref.r1 + dr, c0: ref.c0 + dc, c1: ref.c1 + dc };
    if (!sameSheet(dstSheet, srcSheet) || (ref.sheet === undefined && !sameSheet(formulaSheet, dstSheet))) next.sheet = dstSheet;
    if (ref.sheet === undefined && sameSheet(formulaSheet, dstSheet)) next.sheet = undefined;
    return refToText(next, quoteSheetName);
  });
  return changed ? out : formula;
}

/** Collects the reference tokens of a formula (for highlighting in the editor). */
export function formulaRefTokens(formula: string): Token[] {
  try {
    return tokenize(formula).filter((t) => t.type === 'ref');
  } catch {
    return [];
  }
}

/** Shifts an area-only reference text (defined names, print areas, sqref formulas). */
export function shiftReferenceText(text: string, formulaSheet: string, sheet: string, axis: 'row' | 'col', from: number, delta: number): string {
  return shiftFormula(text, formulaSheet, sheet, axis, from, delta);
}
