/**
 * Reader for Excel 97-2003 workbooks (.xls, BIFF8): values, formulas (decoded from their
 * tokens; the last calculated value is kept when a formula cannot be decoded), number formats,
 * fonts, fills, borders and alignment, merged cells, row heights and column widths, frozen
 * panes, defined names, hyperlinks and notes. The workbook is saved as XLSX.
 */
import { CompoundFile } from './cfb';
import { decodeFormula, errorText, type FormulaContext } from './formula';
import { Workbook, type DefinedName } from '../model/workbook';
import { Sheet, widthCharsToPx, type Cell, type CommentThread, type Hyperlink } from '../model/sheet';
import { BUILTIN_NUMFMTS, BORDER_LINE_STYLES, DEFAULT_FONT, defaultStyle, type AlignmentStyle, type BorderLine, type BorderStyle, type CellStyle, type FontStyle, type HAlign, type UnderlineType, type VAlign } from '../model/style';
import { INDEXED_COLORS, type ColorSpec } from '../model/color';
import { parseError, type Scalar } from '../values';
import type { Range } from '../refs';

interface Rec {
  type: number;
  data: Uint8Array;
  view: DataView;
  offset: number;
}

function readRecords(stream: Uint8Array): Rec[] {
  const out: Rec[] = [];
  const view = new DataView(stream.buffer, stream.byteOffset, stream.byteLength);
  let pos = 0;
  while (pos + 4 <= stream.length) {
    const type = view.getUint16(pos, true);
    const len = view.getUint16(pos + 2, true);
    const data = stream.subarray(pos + 4, Math.min(stream.length, pos + 4 + len));
    out.push({ type, data, view: new DataView(data.buffer, data.byteOffset, data.byteLength), offset: pos });
    pos += 4 + len;
  }
  return out;
}

const R = {
  BOF: 0x0809, EOF: 0x000a, FILEPASS: 0x002f, CODEPAGE: 0x0042, DATEMODE: 0x0022, FONT: 0x0031, FORMAT: 0x041e, XF: 0x00e0, PALETTE: 0x0092,
  BOUNDSHEET: 0x0085, SST: 0x00fc, CONTINUE: 0x003c, NAME: 0x0018, SUPBOOK: 0x01ae, EXTERNSHEET: 0x0017, EXTERNNAME: 0x0023, WINDOW1: 0x003d,
  ROW: 0x0208, COLINFO: 0x007d, DEFCOLWIDTH: 0x0055, STANDARDWIDTH: 0x0099, DEFAULTROWHEIGHT: 0x0225, NUMBER: 0x0203, RK: 0x027e, MULRK: 0x00bd,
  LABELSST: 0x00fd, LABEL: 0x0204, RSTRING: 0x00d6, BOOLERR: 0x0205, BLANK: 0x0201, MULBLANK: 0x00be, FORMULA: 0x0006, STRING: 0x0207,
  SHRFMLA: 0x04bc, ARRAY: 0x0221, MERGEDCELLS: 0x00e5, WINDOW2: 0x023e, PANE: 0x0041, SCL: 0x00a0, HLINK: 0x01b8, NOTE: 0x001c, OBJ: 0x005d, TXO: 0x01b6,
} as const;

/** Cursor over a record and its CONTINUE records (strings may continue in the next record). */
class Continued {
  private part = 0;
  private pos = 0;
  constructor(private readonly parts: Uint8Array[], start = 0) {
    this.pos = start;
  }
  private ensure(): void {
    while (this.part < this.parts.length && this.pos >= this.parts[this.part].length) {
      this.part++;
      this.pos = 0;
    }
  }
  get done(): boolean {
    this.ensure();
    return this.part >= this.parts.length;
  }
  u8(): number {
    this.ensure();
    return this.parts[this.part]?.[this.pos++] ?? 0;
  }
  u16(): number {
    return this.u8() | (this.u8() << 8);
  }
  i32(): number {
    return this.u8() | (this.u8() << 8) | (this.u8() << 16) | (this.u8() << 24);
  }
  skip(n: number): void {
    for (let i = 0; i < n; i++) this.u8();
  }
  /** XLUnicodeRichExtendedString: characters may continue in the next record with a new flag byte. */
  richString(): string {
    const cch = this.u16();
    const flags = this.u8();
    let high = (flags & 1) !== 0;
    const runs = flags & 0x08 ? this.u16() : 0;
    const ext = flags & 0x04 ? this.i32() : 0;
    let s = '';
    for (let i = 0; i < cch; i++) {
      if (this.part < this.parts.length && this.pos >= this.parts[this.part].length) {
        // The rest of the characters are in the next CONTINUE record, after a new flag byte.
        this.part++;
        this.pos = 0;
        if (this.part >= this.parts.length) break;
        high = (this.u8() & 1) !== 0;
      }
      s += String.fromCharCode(high ? this.u16() : this.u8());
    }
    this.skip(runs * 4 + Math.max(0, ext));
    return s;
  }
}

function str8(data: Uint8Array, pos: number, cchSize: 1 | 2): { s: string; end: number } {
  const cch = cchSize === 1 ? data[pos] : data[pos] | (data[pos + 1] << 8);
  let p = pos + cchSize;
  const high = (data[p++] & 1) !== 0;
  let s = '';
  for (let i = 0; i < cch && p < data.length + 1; i++) {
    if (high) {
      s += String.fromCharCode(data[p] | (data[p + 1] << 8));
      p += 2;
    } else s += String.fromCharCode(data[p++]);
  }
  return { s, end: p };
}

function rkValue(rk: number): number {
  let v: number;
  if (rk & 2) v = rk >> 2;
  else {
    const buf = new DataView(new ArrayBuffer(8));
    buf.setUint32(4, rk & 0xfffffffc, true);
    buf.setUint32(0, 0, true);
    v = buf.getFloat64(0, true);
  }
  return rk & 1 ? v / 100 : v;
}

const UNDERLINES: Record<number, UnderlineType> = { 1: 'single', 2: 'double', 0x21: 'singleAccounting', 0x22: 'doubleAccounting' };
const H_ALIGN: HAlign[] = ['general', 'left', 'center', 'right', 'fill', 'justify', 'centerContinuous', 'distributed'];
const V_ALIGN: VAlign[] = ['top', 'center', 'bottom', 'justify', 'distributed'];
const PATTERNS = ['none', 'solid', 'mediumGray', 'darkGray', 'lightGray', 'darkHorizontal', 'darkVertical', 'darkDown', 'darkUp', 'darkGrid', 'darkTrellis', 'lightHorizontal', 'lightVertical', 'lightDown', 'lightUp', 'lightGrid', 'lightTrellis', 'gray125', 'gray0625'];
const LINE_STYLES = [null, 'thin', 'medium', 'dashed', 'dotted', 'thick', 'double', 'hair', 'mediumDashed', 'dashDot', 'mediumDashDot', 'dashDotDot', 'mediumDashDotDot', 'slantDashDot'] as const;
const BUILTIN_NAMES = ['Consolidate_Area', 'Auto_Open', 'Auto_Close', 'Extract', 'Database', 'Criteria', 'Print_Area', 'Print_Titles', 'Recorder', 'Data_Form', 'Auto_Activate', 'Auto_Deactivate', 'Sheet_Title', '_FilterDatabase'];

function icv(index: number): ColorSpec | undefined {
  if (index === 0x7fff || index === 64 || index === 0x40) return undefined;
  if (index === 65 || index === 0x41) return { indexed: 65 };
  return { indexed: index };
}

interface XfRaw {
  font: number;
  fmt: number;
  style: boolean;
  locked: boolean;
  hidden: boolean;
  align: AlignmentStyle;
  border: BorderStyle;
  pattern: number;
  fore: number;
  back: number;
}

function readXf(d: Uint8Array, v: DataView): XfRaw {
  const prot = v.getUint16(4, true);
  const a = d[6];
  const rot = d[7];
  const ind = d[8];
  const b1 = v.getUint32(10, true);
  const b2 = v.getUint32(14, true);
  const colors = v.getUint16(18, true);
  const line = (dg: number, color: number): BorderLine | undefined => {
    const style = LINE_STYLES[dg];
    return style && (BORDER_LINE_STYLES as readonly string[]).includes(style) ? { style, color: icv(color) } : undefined;
  };
  const border: BorderStyle = {};
  const left = line(b1 & 0xf, (b1 >>> 16) & 0x7f);
  const right = line((b1 >>> 4) & 0xf, (b1 >>> 23) & 0x7f);
  const top = line((b1 >>> 8) & 0xf, b2 & 0x7f);
  const bottom = line((b1 >>> 12) & 0xf, (b2 >>> 7) & 0x7f);
  if (left) border.left = left;
  if (right) border.right = right;
  if (top) border.top = top;
  if (bottom) border.bottom = bottom;
  const diagonal = line((b2 >>> 21) & 0xf, (b2 >>> 14) & 0x7f);
  if (diagonal && (b1 >>> 30) & 3) {
    border.diagonal = diagonal;
    if ((b1 >>> 30) & 1) border.diagonalDown = true;
    if ((b1 >>> 30) & 2) border.diagonalUp = true;
  }
  const align: AlignmentStyle = {};
  const h = H_ALIGN[a & 7];
  if (h !== 'general') align.horizontal = h;
  if (a & 8) align.wrap = true;
  const vv = V_ALIGN[(a >> 4) & 7];
  if (vv && vv !== 'bottom') align.vertical = vv;
  if (rot) align.rotation = rot;
  if (ind & 0xf) align.indent = ind & 0xf;
  if (ind & 0x10) align.shrink = true;
  return {
    font: v.getUint16(0, true),
    fmt: v.getUint16(2, true),
    style: (prot & 4) !== 0,
    locked: (prot & 1) !== 0,
    hidden: (prot & 2) !== 0,
    align,
    border,
    pattern: (b2 >>> 26) & 0x3f,
    fore: colors & 0x7f,
    back: (colors >>> 7) & 0x7f,
  };
}

interface SupBook {
  self: boolean;
  names: string[];
}

export function readXls(bytes: Uint8Array): Workbook {
  const cfb = new CompoundFile(bytes);
  const stream = cfb.stream('Workbook', 'Book');
  if (!stream) throw new Error('This file does not contain an Excel workbook.');
  const recs = readRecords(stream);
  if (!recs.length || recs[0].type !== R.BOF) throw new Error('This file does not contain an Excel workbook.');
  const version = recs[0].view.getUint16(0, true);
  if (version !== 0x0600) throw new Error('Excel 5.0/95 workbooks are not supported; save the file in a newer format.');

  const fonts: FontStyle[] = [];
  const formats = new Map<number, string>();
  const xfs: XfRaw[] = [];
  const palette = INDEXED_COLORS.slice();
  const boundSheets: { name: string; offset: number; state: number; kind: number }[] = [];
  let sst: string[] = [];
  const names: { name: string; rgce: Uint8Array; itab: number; hidden: boolean; builtin: boolean }[] = [];
  const supbooks: SupBook[] = [];
  const xtis: { book: number; first: number; last: number }[] = [];
  let date1904 = false;
  let activeTab = 0;

  // Workbook globals.
  let i = 1;
  for (; i < recs.length; i++) {
    const r = recs[i];
    const d = r.data;
    const v = r.view;
    if (r.type === R.EOF) break;
    switch (r.type) {
      case R.FILEPASS:
        throw new Error('Password-protected workbooks cannot be opened.');
      case R.DATEMODE:
        date1904 = v.getUint16(0, true) === 1;
        break;
      case R.FONT: {
        const grbit = v.getUint16(2, true);
        const color = v.getUint16(4, true);
        const name = str8(d, 14, 1).s;
        const font: FontStyle = { name: name || 'Arial', size: v.getUint16(0, true) / 20 };
        if (v.getUint16(6, true) >= 700) font.bold = true;
        if (grbit & 2) font.italic = true;
        if (grbit & 8) font.strike = true;
        const ul = UNDERLINES[d[10]];
        if (ul) font.underline = ul;
        const sss = v.getUint16(8, true);
        if (sss === 1) font.vertAlign = 'superscript';
        else if (sss === 2) font.vertAlign = 'subscript';
        const c = icv(color);
        if (c) font.color = c;
        fonts.push(font);
        break;
      }
      case R.FORMAT:
        formats.set(v.getUint16(0, true), str8(d, 2, 2).s);
        break;
      case R.XF:
        xfs.push(readXf(d, v));
        break;
      case R.PALETTE: {
        const n = v.getUint16(0, true);
        for (let k = 0; k < n && 8 + k < palette.length; k++) {
          const p = 2 + k * 4;
          palette[8 + k] = [d[p], d[p + 1], d[p + 2]].map((x) => x.toString(16).padStart(2, '0')).join('').toUpperCase();
        }
        break;
      }
      case R.BOUNDSHEET:
        boundSheets.push({ offset: v.getUint32(0, true), state: d[4] & 3, kind: d[5], name: str8(d, 6, 1).s });
        break;
      case R.SST: {
        const parts = [d];
        for (let k = i + 1; k < recs.length && recs[k].type === R.CONTINUE; k++) parts.push(recs[k].data);
        const c = new Continued(parts, 8);
        const unique = v.getUint32(4, true);
        sst = [];
        for (let k = 0; k < unique && !c.done; k++) sst.push(c.richString());
        break;
      }
      case R.NAME: {
        const flags = v.getUint16(0, true);
        const cch = d[3];
        const cce = v.getUint16(4, true);
        const itab = v.getUint16(8, true);
        const high = d[14] & 1;
        let p = 15;
        let name = '';
        for (let k = 0; k < cch; k++) {
          if (high) {
            name += String.fromCharCode(d[p] | (d[p + 1] << 8));
            p += 2;
          } else name += String.fromCharCode(d[p++]);
        }
        const builtin = (flags & 0x20) !== 0;
        if (builtin && name.length === 1) name = `_xlnm.${BUILTIN_NAMES[name.charCodeAt(0)] ?? 'Unknown'}`;
        names.push({ name, rgce: d.subarray(p, p + cce), itab, hidden: (flags & 1) !== 0, builtin });
        break;
      }
      case R.SUPBOOK: {
        const marker = v.getUint16(2, true);
        supbooks.push({ self: marker === 0x0401, names: [] });
        break;
      }
      case R.EXTERNNAME: {
        const book = supbooks[supbooks.length - 1];
        if (book) book.names.push(str8(d, 6, 1).s);
        break;
      }
      case R.EXTERNSHEET: {
        const n = v.getUint16(0, true);
        for (let k = 0; k < n; k++) xtis.push({ book: v.getUint16(2 + k * 6, true), first: v.getInt16(4 + k * 6, true), last: v.getInt16(6 + k * 6, true) });
        break;
      }
      case R.WINDOW1:
        if (d.length >= 12) activeTab = v.getUint16(10, true);
        break;
      default:
        break;
    }
  }

  // Styles: every XF becomes a cell style (font index 4 does not exist in BIFF).
  const fontAt = (index: number) => fonts[index >= 4 ? index - 1 : index] ?? fonts[0] ?? { ...DEFAULT_FONT, name: 'Arial', size: 10 };
  const toStyle = (xf: XfRaw | undefined): CellStyle => {
    if (!xf) return defaultStyle({ ...fontAt(0) });
    const style: CellStyle = { font: { ...fontAt(xf.font) }, numFmt: formats.get(xf.fmt) ?? BUILTIN_NUMFMTS[xf.fmt] ?? 'General' };
    if (Object.keys(xf.align).length) style.align = xf.align;
    if (Object.keys(xf.border).length) style.border = xf.border;
    if (xf.pattern > 0) {
      const pattern = PATTERNS[xf.pattern] ?? 'solid';
      style.fill = pattern === 'solid' ? { pattern, fg: icv(xf.fore) ?? { indexed: 64 } } : { pattern, fg: icv(xf.fore), bg: icv(xf.back) };
    }
    if (!xf.locked) style.locked = false;
    if (xf.hidden) style.hidden = true;
    return style;
  };
  const defaultXf = xfs.findIndex((x, k) => k >= 15 && !x.style);
  const wb = new Workbook(toStyle(xfs[defaultXf >= 0 ? defaultXf : 0]));
  wb.date1904 = date1904;
  wb.theme.indexed = palette;
  const styleOf: number[] = xfs.map((xf) => (xf.style ? 0 : wb.styles.intern(toStyle(xf))));

  const sheets: (Sheet | null)[] = boundSheets.map((b) => {
    if (b.kind !== 0) return null;
    const sheet = wb.addSheet(b.name.replace(/[\\/?*[\]:]/g, '_').slice(0, 31) || `Sheet${wb.sheets.length + 1}`);
    sheet.state = b.state === 1 ? 'hidden' : b.state === 2 ? 'veryHidden' : 'visible';
    return sheet;
  });
  const sheetNames = boundSheets.map((b) => b.name);
  const formulaContext = (row: number, col: number, shared?: FormulaContext['shared']): FormulaContext => ({
    format: 'biff8',
    row,
    col,
    sheetOf: (ixti) => {
      const x = xtis[ixti];
      const book = x ? supbooks[x.book] : undefined;
      if (!x || !book?.self || x.first < 0 || x.last < 0) return null;
      const first = sheetNames[x.first];
      const last = sheetNames[x.last];
      return first && last ? { first, last } : null;
    },
    nameOf: (index) => names[index - 1]?.name ?? null,
    externName: (ixti, index) => {
      const x = xtis[ixti];
      const book = x ? supbooks[x.book] : undefined;
      return book?.names[index - 1] ?? null;
    },
    shared,
  });

  // Defined names.
  for (const n of names) {
    if (n.name.startsWith('_xlnm._FilterDatabase')) continue;
    const formula = decodeFormula(n.rgce, new Uint8Array(0), formulaContext(0, 0));
    if (formula === null) continue;
    const def: DefinedName = { name: n.name, formula };
    if (n.itab > 0) {
      const sheet = sheets[n.itab - 1];
      if (!sheet) continue;
      def.localSheet = wb.sheets.indexOf(sheet);
    }
    if (n.hidden) def.hidden = true;
    wb.names.push(def);
  }

  // Worksheets.
  boundSheets.forEach((b, index) => {
    const sheet = sheets[index];
    if (!sheet) return;
    const start = recs.findIndex((r) => r.offset === b.offset);
    if (start < 0) return;
    readSheet(sheet, recs, start + 1, { sst, styleOf, formulaContext, wb });
  });
  if (!wb.sheets.length) wb.addSheet('Sheet1');
  const active = sheets[activeTab];
  wb.activeSheetIndex = active ? wb.sheets.indexOf(active) : 0;
  return wb;
}

interface SheetEnv {
  sst: string[];
  styleOf: number[];
  formulaContext: (row: number, col: number, shared?: FormulaContext['shared']) => FormulaContext;
  wb: Workbook;
}

function readSheet(sheet: Sheet, recs: Rec[], start: number, env: SheetEnv): void {
  // First pass: shared and array formulas (they follow the first cell that uses them).
  const shared: { range: Range; rgce: Uint8Array; extra: Uint8Array }[] = [];
  const arrays: { range: Range; rgce: Uint8Array; extra: Uint8Array }[] = [];
  let end = start;
  for (; end < recs.length && recs[end].type !== R.EOF; end++) {
    const r = recs[end];
    if (r.type !== R.SHRFMLA && r.type !== R.ARRAY) continue;
    const v = r.view;
    const range = { r0: v.getUint16(0, true), r1: v.getUint16(2, true), c0: r.data[4], c1: r.data[5] };
    const off = r.type === R.SHRFMLA ? 8 : 12;
    const cce = v.getUint16(off, true);
    const entry = { range, rgce: r.data.subarray(off + 2, off + 2 + cce), extra: r.data.subarray(off + 2 + cce) };
    (r.type === R.SHRFMLA ? shared : arrays).push(entry);
  }
  const style = (ixfe: number) => env.styleOf[ixfe] || undefined;
  const put = (r: number, c: number, cell: Cell) => {
    if (cell.s === undefined) delete cell.s;
    sheet.setCell(r, c, cell);
  };
  const objTexts = new Map<number, string>();
  let lastObj = -1;
  for (let k = start; k < end; k++) {
    const rec = recs[k];
    const d = rec.data;
    const v = rec.view;
    switch (rec.type) {
      case R.NUMBER:
        put(v.getUint16(0, true), v.getUint16(2, true), { v: v.getFloat64(6, true), s: style(v.getUint16(4, true)) });
        break;
      case R.RK:
        put(v.getUint16(0, true), v.getUint16(2, true), { v: rkValue(v.getUint32(6, true)), s: style(v.getUint16(4, true)) });
        break;
      case R.MULRK: {
        const row = v.getUint16(0, true);
        const first = v.getUint16(2, true);
        const n = (d.length - 6) / 6;
        for (let j = 0; j < n; j++) put(row, first + j, { v: rkValue(v.getUint32(6 + j * 6, true)), s: style(v.getUint16(4 + j * 6, true)) });
        break;
      }
      case R.LABELSST:
        put(v.getUint16(0, true), v.getUint16(2, true), { v: env.sst[v.getUint32(6, true)] ?? '', s: style(v.getUint16(4, true)) });
        break;
      case R.LABEL:
      case R.RSTRING:
        put(v.getUint16(0, true), v.getUint16(2, true), { v: str8(d, 6, 2).s, s: style(v.getUint16(4, true)) });
        break;
      case R.BOOLERR: {
        const value: Scalar = d[7] ? parseError(errorText(d[6])) : d[6] !== 0;
        put(v.getUint16(0, true), v.getUint16(2, true), { v: value, s: style(v.getUint16(4, true)) });
        break;
      }
      case R.BLANK: {
        const s = style(v.getUint16(4, true));
        if (s) put(v.getUint16(0, true), v.getUint16(2, true), { v: null, s });
        break;
      }
      case R.MULBLANK: {
        const row = v.getUint16(0, true);
        const first = v.getUint16(2, true);
        const n = (d.length - 6) / 2;
        for (let j = 0; j < n; j++) {
          const s = style(v.getUint16(4 + j * 2, true));
          if (s) put(row, first + j, { v: null, s });
        }
        break;
      }
      case R.FORMULA: {
        const row = v.getUint16(0, true);
        const col = v.getUint16(2, true);
        let value: Scalar;
        if (v.getUint16(12, true) === 0xffff) {
          switch (d[6]) {
            case 0: {
              const next = recs[k + 1];
              value = next && next.type === R.STRING ? str8(next.data, 0, 2).s : '';
              break;
            }
            case 1:
              value = d[8] !== 0;
              break;
            case 2:
              value = parseError(errorText(d[8]));
              break;
            default:
              value = '';
          }
        } else value = v.getFloat64(6, true);
        const cce = v.getUint16(20, true);
        const rgce = d.subarray(22, 22 + cce);
        const extra = d.subarray(22 + cce);
        const cell: Cell = { v: value, s: style(v.getUint16(4, true)) };
        if (rgce[0] === 0x01) {
          // A shared or array formula anchored at another cell.
          const ar = rgce.length >= 5 ? (rgce[1] | (rgce[2] << 8)) : -1;
          const ac = rgce.length >= 5 ? (rgce[3] | (rgce[4] << 8)) : -1;
          const array = arrays.find((a) => a.range.r0 === ar && a.range.c0 === ac);
          if (array) {
            if (row === ar && col === ac) {
              const f = decodeFormula(array.rgce, array.extra, env.formulaContext(row, col));
              if (f !== null) {
                cell.f = f;
                cell.arr = `${cellA1(array.range.r0, array.range.c0)}:${cellA1(array.range.r1, array.range.c1)}`;
              }
            }
          } else {
            const sh = shared.find((s) => s.range.r0 === ar && s.range.c0 === ac) ?? shared.find((s) => row >= s.range.r0 && row <= s.range.r1 && col >= s.range.c0 && col <= s.range.c1);
            if (sh) {
              const f = decodeFormula(sh.rgce, sh.extra, env.formulaContext(row, col));
              if (f !== null) cell.f = f;
            }
          }
        } else {
          const f = decodeFormula(rgce, extra, env.formulaContext(row, col));
          if (f !== null) cell.f = f;
        }
        put(row, col, cell);
        break;
      }
      case R.ROW: {
        const row = v.getUint16(0, true);
        const height = v.getUint16(6, true) & 0x7fff;
        const flags = v.getUint16(12, true);
        const info = { ...(sheet.rows.get(row) ?? {}) };
        if (flags & 0x40) {
          info.h = height / 20;
          info.customHeight = true;
        }
        if (flags & 0x20) info.hidden = true;
        if (flags & 0x80) {
          const s = style(v.getUint16(14, true) & 0xfff);
          if (s) info.s = s;
        }
        const outline = flags & 7;
        if (outline) info.outline = outline;
        if (Object.keys(info).length) sheet.rows.set(row, info);
        break;
      }
      case R.COLINFO: {
        const first = v.getUint16(0, true);
        const last = Math.min(v.getUint16(2, true), 16383);
        const width = v.getUint16(4, true) / 256;
        const s = style(v.getUint16(6, true));
        const flags = v.getUint16(8, true);
        for (let c = first; c <= last && c < 256; c++) {
          const info: { w?: number; customWidth?: boolean; hidden?: boolean; s?: number; outline?: number } = { w: width, customWidth: true };
          if (flags & 1) info.hidden = true;
          if (s) info.s = s;
          const outline = (flags >> 8) & 7;
          if (outline) info.outline = outline;
          sheet.cols.set(c, info);
        }
        break;
      }
      case R.DEFCOLWIDTH:
        sheet.defaultColWidthPx = widthCharsToPx(v.getUint16(0, true) + 0.71);
        break;
      case R.STANDARDWIDTH:
        sheet.defaultColWidthPx = widthCharsToPx(v.getUint16(0, true) / 256);
        break;
      case R.DEFAULTROWHEIGHT:
        if (d.length >= 4) sheet.defaultRowHeightPt = Math.max(1, v.getUint16(2, true) / 20);
        break;
      case R.MERGEDCELLS: {
        const n = v.getUint16(0, true);
        for (let j = 0; j < n; j++) {
          const p = 2 + j * 8;
          sheet.merges.push({ r0: v.getUint16(p, true), r1: v.getUint16(p + 2, true), c0: v.getUint16(p + 4, true), c1: v.getUint16(p + 6, true) });
        }
        break;
      }
      case R.WINDOW2: {
        const flags = v.getUint16(0, true);
        sheet.showGridLines = (flags & 2) !== 0;
        if (flags & 0x40) sheet.rightToLeft = true;
        if (d.length >= 14) {
          const zoom = v.getUint16(12, true);
          if (zoom >= 10 && zoom <= 400) sheet.zoom = zoom / 100;
        }
        (sheet as Sheet & { frozenFlag?: boolean }).frozenFlag = (flags & 8) !== 0;
        break;
      }
      case R.SCL: {
        const num = v.getUint16(0, true);
        const den = v.getUint16(2, true);
        if (num && den) sheet.zoom = Math.max(0.1, Math.min(4, num / den));
        break;
      }
      case R.PANE:
        if ((sheet as Sheet & { frozenFlag?: boolean }).frozenFlag) sheet.freeze = { cols: v.getUint16(0, true), rows: v.getUint16(2, true) };
        break;
      case R.HLINK: {
        const link = readHlink(d, v);
        if (link) sheet.hyperlinks.push(link);
        break;
      }
      case R.OBJ:
        // Common object data first: ft (2), cb (2), ot (2), id (2).
        if (d.length >= 8 && v.getUint16(0, true) === 0x15) lastObj = v.getUint16(6, true);
        break;
      case R.TXO: {
        const cch = d.length >= 12 ? v.getUint16(10, true) : 0;
        const parts: Uint8Array[] = [];
        for (let j = k + 1; j < end && recs[j].type === R.CONTINUE; j++) parts.push(recs[j].data);
        if (cch && parts.length) {
          const first = parts[0];
          const high = first[0] & 1;
          let s = '';
          for (let p = 1; s.length < cch && p < first.length; ) {
            if (high) {
              s += String.fromCharCode(first[p] | (first[p + 1] << 8));
              p += 2;
            } else s += String.fromCharCode(first[p++]);
          }
          if (lastObj >= 0) objTexts.set(lastObj, s);
        }
        break;
      }
      case R.NOTE: {
        if (d.length < 8) break;
        const row = v.getUint16(0, true);
        const col = v.getUint16(2, true);
        const id = v.getUint16(6, true);
        const author = d.length > 10 ? str8(d, 8, 2).s : '';
        const text = objTexts.get(id);
        if (text === undefined) break;
        const thread: CommentThread = { id: `xls-note-${row}-${col}`, r: row, c: col, kind: 'note', resolved: false, comments: [{ id: `xls-note-${row}-${col}-0`, author, text: author && text.startsWith(`${author}:`) ? text.slice(author.length + 1).replace(/^\s*\n?/, '') : text }] };
        sheet.comments.push(thread);
        break;
      }
      default:
        break;
    }
  }
  delete (sheet as Sheet & { frozenFlag?: boolean }).frozenFlag;
  sheet.layoutVersion++;
}

function cellA1(r: number, c: number): string {
  let s = '';
  let n = c + 1;
  while (n > 0) {
    const m = (n - 1) % 26;
    s = String.fromCharCode(65 + m) + s;
    n = Math.floor((n - 1) / 26);
  }
  return `${s}${r + 1}`;
}

/** URL and in-document hyperlinks (file monikers are left out). */
function readHlink(d: Uint8Array, v: DataView): Hyperlink | null {
  if (d.length < 32) return null;
  const ref = { r0: v.getUint16(0, true), r1: v.getUint16(2, true), c0: v.getUint16(4, true), c1: v.getUint16(6, true) };
  let p = 8 + 16 + 4;
  const flags = v.getUint32(p, true);
  p += 4;
  const hlString = (): string => {
    const len = v.getUint32(p, true);
    p += 4;
    let s = '';
    for (let i = 0; i < len; i++) {
      const ch = v.getUint16(p, true);
      p += 2;
      if (ch) s += String.fromCharCode(ch);
    }
    return s;
  };
  const link: Hyperlink = { ref };
  try {
    if (flags & 0x10) link.display = hlString();
    if (flags & 0x80) hlString();
    if (flags & 0x01) {
      if (flags & 0x100) link.target = hlString();
      else {
        const clsid = [...d.subarray(p, p + 16)].map((b) => b.toString(16).padStart(2, '0')).join('');
        p += 16;
        if (clsid === 'e0c9ea79f9bace118c8200aa004ba90b') {
          const bytes = v.getUint32(p, true);
          p += 4;
          let s = '';
          for (let i = 0; i + 1 < bytes; i += 2) {
            const ch = v.getUint16(p + i, true);
            if (!ch) break;
            s += String.fromCharCode(ch);
          }
          p += bytes;
          link.target = s;
        } else return null;
      }
    }
    if (flags & 0x08) link.location = hlString();
  } catch {
    // Truncated record: keep what was read.
  }
  return link.target || link.location ? link : null;
}
