/**
 * Fast parser for the <sheetData> element of a worksheet. It builds cells directly instead of a
 * DOM, which keeps memory use low for large sheets (important on phones).
 */
import { decodeEntities } from '../xml/xml';
import { CellError, ERROR_CODES, type ErrorCode, type Scalar } from '../values';
import { cellKey, colIndex, parseRange } from '../refs';
import type { Cell, RichTextRun, RowInfo, Sheet } from '../model/sheet';
import { translateFormula } from '../formula/adjust';
import { serialFromDate } from './util';
import { formulaFromFile } from './formulaText';

export interface SharedString {
  text: string;
  runs?: RichTextRun[];
}

export interface SheetDataContext {
  shared: SharedString[];
  /** Cell metadata indexes (cm) that mark dynamic array formulas. */
  dynamicCm: Set<number>;
  /** Parses rich text runs (<r>) of an inline string. */
  parseRuns: (xml: string) => SharedString;
}

function readAttrs(tag: string): Record<string, string> {
  const attrs: Record<string, string> = {};
  const re = /([\w:.-]+)\s*=\s*("([^"]*)"|'([^']*)')/g;
  let m: RegExpExecArray | null;
  while ((m = re.exec(tag))) attrs[m[1].replace(/^\w+:(?=\w)/, (p) => (p === 'x:' ? '' : p))] = decodeEntities(m[3] ?? m[4] ?? '');
  return attrs;
}

/** Splits a cell reference like "AB12" into [row, col] (0-based). */
function splitRef(ref: string): [number, number] | null {
  let i = 0;
  while (i < ref.length && ref.charCodeAt(i) >= 65) i++;
  const col = colIndex(ref.slice(0, i));
  const row = Number(ref.slice(i)) - 1;
  if (col < 0 || !(row >= 0)) return null;
  return [row, col];
}

const ELEMENT_RE: Record<string, RegExp> = {};

function elementText(xml: string, name: string): string | undefined {
  // Matches <name ...>text</name> (any prefix); undefined if absent, '' if empty.
  let re = ELEMENT_RE[name];
  if (!re) re = ELEMENT_RE[name] = new RegExp(`<(?:\\w+:)?${name}(\\s[^>]*)?(/>|>([\\s\\S]*?)</(?:\\w+:)?${name}>)`);
  const m = re.exec(xml);
  if (!m) return undefined;
  return m[2] === '/>' ? '' : m[3] ?? '';
}

interface SharedFormula {
  text: string;
  r: number;
  c: number;
}

export function parseSheetData(xml: string, sheet: Sheet, ctx: SheetDataContext): void {
  const shared = new Map<string, SharedFormula>();
  const rowRe = /<(?:\w+:)?row(\s[^>]*?)?(\/>|>)/g;
  let rowMatch: RegExpExecArray | null;
  let nextRow = 0;
  const arrayAnchors: { r: number; c: number; range: string }[] = [];
  while ((rowMatch = rowRe.exec(xml))) {
    const rowAttrs = readAttrs(rowMatch[1] ?? '');
    const r = rowAttrs.r ? Number(rowAttrs.r) - 1 : nextRow;
    nextRow = r + 1;
    const info: RowInfo = {};
    if (rowAttrs.ht !== undefined && (rowAttrs.customHeight === '1' || rowAttrs.customHeight === 'true')) {
      info.h = Number(rowAttrs.ht);
      info.customHeight = true;
    } else if (rowAttrs.ht !== undefined) info.h = Number(rowAttrs.ht);
    if (rowAttrs.hidden === '1' || rowAttrs.hidden === 'true') info.hidden = true;
    if ((rowAttrs.customFormat === '1' || rowAttrs.customFormat === 'true') && rowAttrs.s !== undefined) info.s = Number(rowAttrs.s);
    if (rowAttrs.outlineLevel) info.outline = Number(rowAttrs.outlineLevel);
    if (rowAttrs.collapsed === '1') info.collapsed = true;
    if (Object.keys(info).length) sheet.setRowInfo(r, info);
    if (rowMatch[2] === '/>') continue;
    const end = xml.indexOf('</', rowRe.lastIndex);
    // Find the matching </row> (cells contain </v>, </f>... so search for the row close tag).
    const close = xml.slice(end).search(/<\/(?:\w+:)?row>/);
    const body = close < 0 ? xml.slice(rowRe.lastIndex) : xml.slice(rowRe.lastIndex, end + close);
    rowRe.lastIndex = close < 0 ? xml.length : end + close;
    parseRow(body, r, sheet, ctx, shared, arrayAnchors);
  }
  for (const a of arrayAnchors) {
    const range = parseRange(a.range);
    if (!range) continue;
    const key = cellKey(a.r, a.c);
    for (let rr = range.r0; rr <= range.r1; rr++) {
      for (let cc = range.c0; cc <= range.c1; cc++) {
        if (rr === a.r && cc === a.c) continue;
        const cell = sheet.getCell(rr, cc);
        if (cell && cell.f === undefined) sheet.setCell(rr, cc, { ...cell, spill: key });
        else if (!cell) sheet.setCell(rr, cc, { v: null, spill: key });
      }
    }
  }
}

function parseRow(body: string, r: number, sheet: Sheet, ctx: SheetDataContext, shared: Map<string, SharedFormula>, arrays: { r: number; c: number; range: string }[]): void {
  const cellRe = /<(?:\w+:)?c(\s[^>]*?)?(\/>|>)/g;
  let m: RegExpExecArray | null;
  let nextCol = 0;
  while ((m = cellRe.exec(body))) {
    const attrs = readAttrs(m[1] ?? '');
    let c = nextCol;
    if (attrs.r) {
      const rc = splitRef(attrs.r);
      if (rc) c = rc[1];
    }
    nextCol = c + 1;
    const s = attrs.s !== undefined ? Number(attrs.s) : undefined;
    let inner = '';
    if (m[2] === '>') {
      const end = body.indexOf('</c>', cellRe.lastIndex);
      const endX = end < 0 ? body.search(/<\/\w+:c>/) : end;
      inner = body.slice(cellRe.lastIndex, endX < 0 ? body.length : endX);
      cellRe.lastIndex = endX < 0 ? body.length : endX + 4;
    }
    const cell: Cell = { v: null };
    if (s) cell.s = s;
    const t = attrs.t ?? 'n';
    const vText = inner ? elementText(inner, 'v') : undefined;
    switch (t) {
      case 's': {
        const idx = Number(vText);
        const ss = ctx.shared[idx];
        if (ss) {
          cell.v = ss.text;
          if (ss.runs) cell.rt = ss.runs;
        } else cell.v = '';
        break;
      }
      case 'inlineStr': {
        const is = elementText(inner, 'is');
        if (is !== undefined) {
          const parsed = ctx.parseRuns(is);
          cell.v = parsed.text;
          if (parsed.runs) cell.rt = parsed.runs;
        }
        break;
      }
      case 'str':
        cell.v = vText !== undefined ? decodeEntities(vText) : '';
        break;
      case 'b':
        cell.v = vText === '1' || vText === 'true';
        break;
      case 'e': {
        const code = (vText ?? '').trim() as ErrorCode;
        cell.v = (ERROR_CODES as readonly string[]).includes(code) ? CellError.of(code) : CellError.of('#VALUE!');
        break;
      }
      case 'd':
        cell.v = vText ? serialFromDate(vText) : null;
        break;
      default:
        cell.v = vText !== undefined && vText !== '' ? Number(vText) : null;
        if (typeof cell.v === 'number' && Number.isNaN(cell.v)) cell.v = null;
    }
    // Formula.
    const fMatch = inner ? /<(?:\w+:)?f(\s[^>]*?)?(\/>|>([\s\S]*?)<\/(?:\w+:)?f>)/.exec(inner) : null;
    if (fMatch) {
      const fAttrs = readAttrs(fMatch[1] ?? '');
      let text = fMatch[2] === '/>' ? '' : decodeEntities(fMatch[3] ?? '');
      const ft = fAttrs.t ?? 'normal';
      if (ft === 'shared' && fAttrs.si !== undefined) {
        if (text) shared.set(fAttrs.si, { text, r, c });
        else {
          const master = shared.get(fAttrs.si);
          if (master) text = translateFormula(master.text, r - master.r, c - master.c);
        }
      }
      if (ft === 'dataTable') {
        // Data tables keep their cached values.
      } else if (text) {
        cell.f = formulaFromFile(text);
        if (ft === 'array' && fAttrs.ref) {
          cell.arr = fAttrs.ref;
          const cm = attrs.cm !== undefined ? Number(attrs.cm) : undefined;
          if (cm !== undefined && ctx.dynamicCm.has(cm)) cell.dyn = true;
          if (fAttrs.ref.includes(':')) arrays.push({ r, c, range: fAttrs.ref });
        }
      }
    }
    if (cell.v !== null || cell.f !== undefined || cell.s) sheet.setCell(r, c, cell);
  }
}

export function scalarFromText(t: string): Scalar {
  return t;
}
