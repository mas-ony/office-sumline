/**
 * Reader for Excel binary workbooks (.xlsb, BIFF12 records in a ZIP package): values, formulas
 * (decoded from their tokens; the last calculated value is kept when a formula cannot be
 * decoded), styles, merged cells, sizes, frozen panes, defined names and hyperlinks. The
 * workbook is saved as XLSX.
 */
import { strFromU8, unzipSync } from 'fflate';
import { Workbook, type DefinedName } from '../model/workbook';
import { Sheet, widthCharsToPx, type Cell } from '../model/sheet';
import { BUILTIN_NUMFMTS, BORDER_LINE_STYLES, defaultStyle, type AlignmentStyle, type BorderLine, type BorderStyle, type CellStyle, type FillStyle, type FontStyle, type HAlign, type UnderlineType, type VAlign } from '../model/style';
import type { ColorSpec } from '../model/color';
import { parseError, type Scalar } from '../values';
import { decodeFormula, errorText, type FormulaContext } from '../xls/formula';
import { attr, elements, parseXml } from '../xml/xml';
import type { Range } from '../refs';

interface Rec {
  type: number;
  data: Uint8Array;
  view: DataView;
}

function records(data: Uint8Array): Rec[] {
  const out: Rec[] = [];
  let p = 0;
  while (p < data.length) {
    let b = data[p++];
    let type = b & 0x7f;
    if (b & 0x80) {
      b = data[p++];
      type |= (b & 0x7f) << 7;
    }
    let size = 0;
    for (let i = 0; i < 4; i++) {
      b = data[p++];
      size |= (b & 0x7f) << (7 * i);
      if (!(b & 0x80)) break;
    }
    const body = data.subarray(p, Math.min(data.length, p + size));
    out.push({ type, data: body, view: new DataView(body.buffer, body.byteOffset, body.byteLength) });
    p += size;
  }
  return out;
}

/** XLWideString at an offset: 32-bit count of UTF-16 characters. */
function wide(v: DataView, at: number): { s: string; end: number } {
  if (at + 4 > v.byteLength) return { s: '', end: at };
  const cch = v.getUint32(at, true);
  if (cch === 0xffffffff) return { s: '', end: at + 4 };
  let s = '';
  let p = at + 4;
  for (let i = 0; i < cch && p + 2 <= v.byteLength; i++, p += 2) s += String.fromCharCode(v.getUint16(p, true));
  return { s, end: p };
}

function rkValue(rk: number): number {
  let v: number;
  if (rk & 2) v = rk >> 2;
  else {
    const buf = new DataView(new ArrayBuffer(8));
    buf.setUint32(4, rk & 0xfffffffc, true);
    v = buf.getFloat64(0, true);
  }
  return rk & 1 ? v / 100 : v;
}

function color(v: DataView, at: number): ColorSpec | undefined {
  const kind = v.getUint8(at) >> 1;
  const index = v.getUint8(at + 1);
  const tint = v.getInt16(at + 2, true) / 32767;
  const withTint = <T extends ColorSpec>(c: T): T => (tint ? ({ ...c, tint: Math.round(tint * 1000) / 1000 } as T) : c);
  switch (kind) {
    case 1:
      return index === 64 ? undefined : withTint({ indexed: index });
    case 2: {
      const hex = [v.getUint8(at + 4), v.getUint8(at + 5), v.getUint8(at + 6)].map((x) => x.toString(16).padStart(2, '0')).join('').toUpperCase();
      return withTint({ rgb: `FF${hex}` });
    }
    case 3:
      return withTint({ theme: index });
    default:
      return undefined;
  }
}

const UNDERLINES: Record<number, UnderlineType> = { 1: 'single', 2: 'double', 0x21: 'singleAccounting', 0x22: 'doubleAccounting' };
const H_ALIGN: HAlign[] = ['general', 'left', 'center', 'right', 'fill', 'justify', 'centerContinuous', 'distributed'];
const V_ALIGN: VAlign[] = ['top', 'center', 'bottom', 'justify', 'distributed'];
const PATTERNS = ['none', 'solid', 'mediumGray', 'darkGray', 'lightGray', 'darkHorizontal', 'darkVertical', 'darkDown', 'darkUp', 'darkGrid', 'darkTrellis', 'lightHorizontal', 'lightVertical', 'lightDown', 'lightUp', 'lightGrid', 'lightTrellis', 'gray125', 'gray0625'];
const LINE_STYLES = [null, 'thin', 'medium', 'dashed', 'dotted', 'thick', 'double', 'hair', 'mediumDashed', 'dashDot', 'mediumDashDot', 'dashDotDot', 'mediumDashDotDot', 'slantDashDot'] as const;
const BUILTIN_NAMES: Record<string, string> = { Print_Area: '_xlnm.Print_Area', Print_Titles: '_xlnm.Print_Titles', _FilterDatabase: '_xlnm._FilterDatabase', Criteria: '_xlnm.Criteria', Extract: '_xlnm.Extract' };

const T = {
  ROW: 0, BLANK: 1, RK: 2, ERROR: 3, BOOL: 4, REAL: 5, ST: 6, ISST: 7, FMLA_STRING: 8, FMLA_NUM: 9, FMLA_BOOL: 10, FMLA_ERROR: 11, SST_ITEM: 19,
  NAME: 39, FONT: 43, FMT: 44, FILL: 45, BORDER: 46, XF: 47, COLINFO: 60, BOOKVIEW: 135, WSVIEW: 137, PANE: 151, WBPROP: 153, BUNDLESH: 156,
  MERGE: 176, SUP_SELF: 355, SUP_SAME: 356, SUP_ADDIN: 358, SUP_BOOKSRC: 359, EXTERNSHEET: 362, SHRFMLA: 426, ARRFMLA: 427, WSFMTINFO: 485,
  HLINK: 494, BEGIN_CELLXFS: 617, END_CELLXFS: 618,
} as const;

function relsOf(files: Record<string, Uint8Array>, part: string): Map<string, { type: string; target: string }> {
  const dir = part.slice(0, part.lastIndexOf('/') + 1);
  const relsPath = `${dir}_rels/${part.slice(dir.length)}.rels`;
  const out = new Map<string, { type: string; target: string }>();
  const data = files[relsPath];
  if (!data) return out;
  for (const r of elements(parseXml(strFromU8(data)).root)) {
    const target = attr(r, 'Target') ?? '';
    const resolved = target.startsWith('/') ? target.slice(1) : attr(r, 'TargetMode') === 'External' ? target : normalize(dir + target);
    out.set(attr(r, 'Id') ?? '', { type: attr(r, 'Type') ?? '', target: resolved });
  }
  return out;
}

function normalize(path: string): string {
  const parts: string[] = [];
  for (const p of path.split('/')) {
    if (p === '..') parts.pop();
    else if (p && p !== '.') parts.push(p);
  }
  return parts.join('/');
}

/** Parsed formula of a record body: tokens and their extra data. */
function formulaAt(v: DataView, data: Uint8Array, at: number): { rgce: Uint8Array; extra: Uint8Array } | null {
  if (at + 4 > v.byteLength) return null;
  const cce = v.getUint32(at, true);
  const rgce = data.subarray(at + 4, at + 4 + cce);
  const cbAt = at + 4 + cce;
  const cb = cbAt + 4 <= v.byteLength ? v.getUint32(cbAt, true) : 0;
  return { rgce, extra: data.subarray(cbAt + 4, cbAt + 4 + cb) };
}

export function readXlsb(bytes: Uint8Array): Workbook {
  let files: Record<string, Uint8Array>;
  try {
    files = unzipSync(bytes);
  } catch {
    throw new Error('This is not a valid XLSB file.');
  }
  const wbPath = 'xl/workbook.bin';
  const wbData = files[wbPath];
  if (!wbData) throw new Error('This is not a valid XLSB file.');
  const wbRels = relsOf(files, wbPath);
  const relOfType = (suffix: string) => [...wbRels.values()].find((r) => r.type.endsWith(suffix))?.target;

  // Shared strings.
  const sst: string[] = [];
  const sstData = files[relOfType('/sharedStrings') ?? 'xl/sharedStrings.bin'];
  if (sstData) for (const r of records(sstData)) if (r.type === T.SST_ITEM) sst.push(wide(r.view, 1).s);

  // Styles.
  const fonts: FontStyle[] = [];
  const fills: (FillStyle | undefined)[] = [];
  const borders: (BorderStyle | undefined)[] = [];
  const formats = new Map<number, string>();
  const xfs: { font: number; fmt: number; fill: number; border: number; align: AlignmentStyle; locked: boolean; hidden: boolean }[] = [];
  const stylesData = files[relOfType('/styles') ?? 'xl/styles.bin'];
  if (stylesData) {
    let inCellXfs = false;
    for (const r of records(stylesData)) {
      const v = r.view;
      switch (r.type) {
        case T.FMT:
          formats.set(v.getUint16(0, true), wide(v, 2).s);
          break;
        case T.FONT: {
          if (r.data.length < 25) break;
          const grbit = v.getUint16(2, true);
          const font: FontStyle = { name: wide(v, 21).s || 'Calibri', size: v.getUint16(0, true) / 20 };
          if (v.getUint16(4, true) >= 700) font.bold = true;
          if (grbit & 2) font.italic = true;
          if (grbit & 8) font.strike = true;
          const sss = v.getUint16(6, true);
          if (sss === 1) font.vertAlign = 'superscript';
          else if (sss === 2) font.vertAlign = 'subscript';
          const ul = UNDERLINES[v.getUint8(8)];
          if (ul) font.underline = ul;
          const c = color(v, 12);
          if (c) font.color = c;
          const scheme = v.getUint8(20);
          if (scheme === 1) font.scheme = 'major';
          else if (scheme === 2) font.scheme = 'minor';
          fonts.push(font);
          break;
        }
        case T.FILL: {
          const fls = v.getUint32(0, true);
          const pattern = PATTERNS[fls];
          if (!pattern || pattern === 'none') fills.push(undefined);
          else if (pattern === 'solid') fills.push({ pattern, fg: color(v, 4) ?? { indexed: 64 } });
          else fills.push({ pattern, fg: color(v, 4), bg: color(v, 12) });
          break;
        }
        case T.BORDER: {
          const flags = v.getUint8(0);
          const side = (k: number): BorderLine | undefined => {
            const at = 1 + k * 10;
            if (at + 10 > v.byteLength) return undefined;
            const style = LINE_STYLES[v.getUint8(at)];
            return style && (BORDER_LINE_STYLES as readonly string[]).includes(style) ? { style, color: color(v, at + 2) } : undefined;
          };
          const b: BorderStyle = {};
          const [top, bottom, left, right, diag] = [0, 1, 2, 3, 4].map(side);
          if (top) b.top = top;
          if (bottom) b.bottom = bottom;
          if (left) b.left = left;
          if (right) b.right = right;
          if (diag && flags & 3) {
            b.diagonal = diag;
            if (flags & 1) b.diagonalDown = true;
            if (flags & 2) b.diagonalUp = true;
          }
          borders.push(Object.keys(b).length ? b : undefined);
          break;
        }
        case T.BEGIN_CELLXFS:
          inCellXfs = true;
          break;
        case T.END_CELLXFS:
          inCellXfs = false;
          break;
        case T.XF: {
          if (!inCellXfs) break;
          const flags = v.getUint16(12, true);
          const align: AlignmentStyle = {};
          const h = H_ALIGN[flags & 7];
          if (h && h !== 'general') align.horizontal = h;
          const va = V_ALIGN[(flags >> 3) & 7];
          if (va && va !== 'bottom') align.vertical = va;
          if (flags & 0x40) align.wrap = true;
          if (flags & 0x100) align.shrink = true;
          const rot = v.getUint8(10);
          if (rot) align.rotation = rot;
          const indent = v.getUint8(11);
          if (indent) align.indent = indent;
          xfs.push({ fmt: v.getUint16(2, true), font: v.getUint16(4, true), fill: v.getUint16(6, true), border: v.getUint16(8, true), align, locked: (flags & 0x1000) !== 0, hidden: (flags & 0x2000) !== 0 });
          break;
        }
        default:
          break;
      }
    }
  }
  const toStyle = (xf: (typeof xfs)[number] | undefined): CellStyle => {
    if (!xf) return defaultStyle(fonts[0] ? { ...fonts[0] } : undefined);
    const style: CellStyle = { font: { ...(fonts[xf.font] ?? fonts[0] ?? { name: 'Calibri', size: 11 }) }, numFmt: formats.get(xf.fmt) ?? BUILTIN_NUMFMTS[xf.fmt] ?? 'General' };
    const fill = fills[xf.fill];
    if (fill) style.fill = fill;
    const border = borders[xf.border];
    if (border) style.border = border;
    if (Object.keys(xf.align).length) style.align = xf.align;
    if (!xf.locked) style.locked = false;
    if (xf.hidden) style.hidden = true;
    return style;
  };
  const wb = new Workbook(toStyle(xfs[0]));
  const styleOf = xfs.map((xf) => wb.styles.intern(toStyle(xf)));

  // Workbook: sheets, names, external sheet references.
  const bundles: { name: string; rel: string; state: number }[] = [];
  const names: { name: string; itab: number; formula: { rgce: Uint8Array; extra: Uint8Array } | null; hidden: boolean }[] = [];
  const supbooks: boolean[] = [];
  const xtis: { book: number; first: number; last: number }[] = [];
  let activeTab = 0;
  for (const r of records(wbData)) {
    const v = r.view;
    switch (r.type) {
      case T.WBPROP:
        wb.date1904 = (v.getUint32(0, true) & 1) !== 0;
        break;
      case T.BUNDLESH: {
        const rel = wide(v, 8);
        bundles.push({ state: v.getUint32(0, true), rel: rel.s, name: wide(v, rel.end).s });
        break;
      }
      case T.NAME: {
        const flags = v.getUint32(0, true);
        const itab = v.getUint32(5, true);
        const nm = wide(v, 9);
        const builtin = (flags & 0x20) !== 0;
        names.push({ name: builtin && BUILTIN_NAMES[nm.s] ? BUILTIN_NAMES[nm.s] : nm.s, itab: itab === 0xffffffff ? -1 : itab, formula: formulaAt(v, r.data, nm.end), hidden: (flags & 1) !== 0 });
        break;
      }
      case T.SUP_SELF:
      case T.SUP_SAME:
        supbooks.push(true);
        break;
      case T.SUP_ADDIN:
      case T.SUP_BOOKSRC:
        supbooks.push(false);
        break;
      case T.EXTERNSHEET: {
        const n = v.getUint32(0, true);
        for (let k = 0; k < n; k++) xtis.push({ book: v.getUint32(4 + k * 12, true), first: v.getInt32(8 + k * 12, true), last: v.getInt32(12 + k * 12, true) });
        break;
      }
      case T.BOOKVIEW:
        if (r.data.length >= 28) activeTab = v.getUint32(24, true);
        break;
      default:
        break;
    }
  }
  const sheetNames = bundles.map((b) => b.name);
  const ctxFor = (row: number, col: number, shared?: FormulaContext['shared']): FormulaContext => ({
    format: 'biff12',
    row,
    col,
    sheetOf: (ixti) => {
      const x = xtis[ixti];
      if (!x || !supbooks[x.book] || x.first < 0 || x.last < 0) return null;
      const first = sheetNames[x.first];
      const last = sheetNames[x.last];
      return first && last ? { first, last } : null;
    },
    nameOf: (index) => names[index - 1]?.name ?? null,
    externName: () => null,
    shared,
  });

  const sheets: (Sheet | null)[] = bundles.map((b) => {
    const target = wbRels.get(b.rel);
    if (!target || !/worksheet/i.test(target.type)) return null;
    const sheet = wb.addSheet(b.name || `Sheet${wb.sheets.length + 1}`);
    sheet.state = b.state === 1 ? 'hidden' : b.state === 2 ? 'veryHidden' : 'visible';
    return sheet;
  });
  for (const n of names) {
    if (!n.formula || n.name === '_xlnm._FilterDatabase') continue;
    const formula = decodeFormula(n.formula.rgce, n.formula.extra, ctxFor(0, 0));
    if (formula === null) continue;
    const def: DefinedName = { name: n.name, formula };
    if (n.itab >= 0) {
      const sheet = sheets[n.itab];
      if (!sheet) continue;
      def.localSheet = wb.sheets.indexOf(sheet);
    }
    if (n.hidden) def.hidden = true;
    wb.names.push(def);
  }

  bundles.forEach((b, i) => {
    const sheet = sheets[i];
    const target = wbRels.get(b.rel);
    if (!sheet || !target) return;
    const data = files[target.target];
    if (data) readSheet(sheet, data, relsOf(files, target.target), { sst, styleOf, ctxFor });
  });
  if (!wb.sheets.length) wb.addSheet('Sheet1');
  const active = sheets[activeTab];
  wb.activeSheetIndex = active ? wb.sheets.indexOf(active) : 0;
  return wb;
}

interface SheetEnv {
  sst: string[];
  styleOf: number[];
  ctxFor: (row: number, col: number, shared?: FormulaContext['shared']) => FormulaContext;
}

function readSheet(sheet: Sheet, data: Uint8Array, rels: Map<string, { type: string; target: string }>, env: SheetEnv): void {
  const recs = records(data);
  const shared: { range: Range; rgce: Uint8Array; extra: Uint8Array }[] = [];
  const arrays: { range: Range; rgce: Uint8Array; extra: Uint8Array }[] = [];
  const rfx = (v: DataView, at = 0): Range => ({ r0: v.getUint32(at, true), r1: v.getUint32(at + 4, true), c0: v.getUint32(at + 8, true), c1: v.getUint32(at + 12, true) });
  for (const r of recs) {
    if (r.type === T.SHRFMLA || r.type === T.ARRFMLA) {
      const f = formulaAt(r.view, r.data, r.type === T.SHRFMLA ? 16 : 17);
      if (f) (r.type === T.SHRFMLA ? shared : arrays).push({ range: rfx(r.view), ...f });
    }
  }
  const sharedText = (row: number, col: number) => (anchorRow: number): string | null => {
    const sh = shared.find((s) => s.range.r0 === anchorRow && col >= s.range.c0 && col <= s.range.c1 && row >= s.range.r0 && row <= s.range.r1);
    return sh ? decodeFormula(sh.rgce, sh.extra, env.ctxFor(row, col)) : null;
  };
  let row = 0;
  let frozen = false;
  for (const r of recs) {
    const v = r.view;
    const d = r.data;
    if (r.type >= T.BLANK && r.type <= T.FMLA_ERROR) {
      const col = v.getUint32(0, true);
      const s = env.styleOf[v.getUint32(4, true) & 0xffffff] || undefined;
      const cell: Cell = { v: null };
      if (s) cell.s = s;
      let formulaAtOffset = -1;
      switch (r.type) {
        case T.BLANK:
          if (!s) continue;
          break;
        case T.RK:
          cell.v = rkValue(v.getUint32(8, true));
          break;
        case T.ERROR:
          cell.v = parseError(errorText(d[8]));
          break;
        case T.BOOL:
          cell.v = d[8] !== 0;
          break;
        case T.REAL:
          cell.v = v.getFloat64(8, true);
          break;
        case T.ST:
          cell.v = wide(v, 8).s;
          break;
        case T.ISST:
          cell.v = env.sst[v.getUint32(8, true)] ?? '';
          break;
        case T.FMLA_STRING: {
          const w = wide(v, 8);
          cell.v = w.s;
          formulaAtOffset = w.end + 2;
          break;
        }
        case T.FMLA_NUM:
          cell.v = v.getFloat64(8, true);
          formulaAtOffset = 18;
          break;
        case T.FMLA_BOOL:
          cell.v = d[8] !== 0;
          formulaAtOffset = 11;
          break;
        case T.FMLA_ERROR:
          cell.v = parseError(errorText(d[8])) as Scalar;
          formulaAtOffset = 11;
          break;
      }
      if (formulaAtOffset >= 0) {
        const f = formulaAt(v, d, formulaAtOffset);
        if (f && f.rgce.length) {
          if (f.rgce[0] === 0x01) {
            const array = arrays.find((a) => a.range.r0 === row && a.range.c0 === col);
            if (array) {
              const text = decodeFormula(array.rgce, array.extra, env.ctxFor(row, col));
              if (text !== null) {
                cell.f = text;
                cell.arr = `${a1(array.range.r0, array.range.c0)}:${a1(array.range.r1, array.range.c1)}`;
              }
            } else if (!arrays.some((a) => row >= a.range.r0 && row <= a.range.r1 && col >= a.range.c0 && col <= a.range.c1)) {
              const resolver = sharedText(row, col);
              const text = decodeFormula(f.rgce, f.extra, env.ctxFor(row, col, (anchorRow) => resolver(anchorRow)));
              if (text !== null) cell.f = text;
            }
          } else {
            const text = decodeFormula(f.rgce, f.extra, env.ctxFor(row, col));
            if (text !== null) cell.f = text;
          }
        }
      }
      sheet.setCell(row, col, cell);
      continue;
    }
    switch (r.type) {
      case T.ROW: {
        row = v.getUint32(0, true);
        const height = v.getUint16(8, true);
        const flags = d[11];
        const info = { ...(sheet.rows.get(row) ?? {}) };
        if (flags & 0x20) {
          info.h = height / 20;
          info.customHeight = true;
        }
        if (flags & 0x10) info.hidden = true;
        if (flags & 0x40) {
          const s = env.styleOf[v.getUint32(4, true)];
          if (s) info.s = s;
        }
        if (flags & 7) info.outline = flags & 7;
        if (Object.keys(info).length) sheet.rows.set(row, info);
        break;
      }
      case T.COLINFO: {
        const first = v.getUint32(0, true);
        const last = Math.min(v.getUint32(4, true), first + 16383, 16383);
        const width = v.getUint32(8, true) / 256;
        const s = env.styleOf[v.getUint32(12, true)];
        const flags = v.getUint16(16, true);
        for (let c = first; c <= last; c++) {
          const info: { w?: number; customWidth?: boolean; hidden?: boolean; s?: number; outline?: number } = { w: width, customWidth: true };
          if (flags & 1) info.hidden = true;
          if (s) info.s = s;
          const outline = (flags >> 8) & 7;
          if (outline) info.outline = outline;
          sheet.cols.set(c, info);
        }
        break;
      }
      case T.WSFMTINFO: {
        const dx = v.getUint32(0, true);
        if (dx !== 0xffffffff && dx > 0) sheet.defaultColWidthPx = widthCharsToPx(dx / 256);
        else sheet.defaultColWidthPx = widthCharsToPx(v.getUint16(4, true) + 0.71);
        const h = v.getUint16(6, true);
        if (h) sheet.defaultRowHeightPt = h / 20;
        break;
      }
      case T.MERGE:
        sheet.merges.push(rfx(v));
        break;
      case T.WSVIEW: {
        const flags = v.getUint16(0, true);
        sheet.showGridLines = (flags & 4) !== 0;
        if (flags & 0x20) sheet.rightToLeft = true;
        if (d.length >= 18) {
          const zoom = v.getUint16(16, true);
          if (zoom >= 10 && zoom <= 400) sheet.zoom = zoom / 100;
        }
        break;
      }
      case T.PANE: {
        frozen = (d[28] & 1) !== 0;
        if (frozen) sheet.freeze = { cols: Math.round(v.getFloat64(0, true)), rows: Math.round(v.getFloat64(8, true)) };
        break;
      }
      case T.HLINK: {
        const ref = rfx(v);
        const rel = wide(v, 16);
        const location = wide(v, rel.end);
        const tooltip = wide(v, location.end);
        const display = wide(v, tooltip.end);
        const target = rel.s ? rels.get(rel.s)?.target : undefined;
        if (target || location.s) sheet.hyperlinks.push({ ref, target, location: location.s || undefined, tooltip: tooltip.s || undefined, display: display.s || undefined });
        break;
      }
      default:
        break;
    }
  }
  void frozen;
  sheet.layoutVersion++;
}

function a1(r: number, c: number): string {
  let s = '';
  let n = c + 1;
  while (n > 0) {
    const m = (n - 1) % 26;
    s = String.fromCharCode(65 + m) + s;
    n = Math.floor((n - 1) / 26);
  }
  return `${s}${r + 1}`;
}
