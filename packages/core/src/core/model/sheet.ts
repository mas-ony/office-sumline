import { type Range, MAX_COLS, MAX_ROWS, cellKey, keyCol, keyRow, rangeContains, rangesIntersect } from '../refs';
import type { Scalar } from '../values';
import type { ColorSpec } from './color';
import type { FontStyle } from './style';
import type { XmlElement } from '../xml/xml';
import type { SheetChart, SheetShape } from './drawing';

export interface RichTextRun {
  text: string;
  font?: Partial<FontStyle>;
}

/**
 * A cell. Cells are treated as immutable values: every change replaces the object, which keeps
 * undo snapshots cheap.
 */
export interface Cell {
  /** The value; for formula cells the last computed result. */
  v: Scalar;
  /** Formula without the leading '='. */
  f?: string;
  /** Style index into the workbook StyleTable (0 or undefined = default). */
  s?: number;
  /** Rich text runs for string values. */
  rt?: RichTextRun[];
  /** Range of a (legacy or dynamic) array formula; set on its anchor cell. */
  arr?: string;
  /** True when the array formula is a dynamic array (spills). */
  dyn?: boolean;
  /** Anchor key of the array formula this cell's value spilled from. */
  spill?: number;
}

export interface RowInfo {
  /** Height in points. */
  h?: number;
  customHeight?: boolean;
  hidden?: boolean;
  /** Row style index. */
  s?: number;
  outline?: number;
  collapsed?: boolean;
}

export interface ColInfo {
  /** Width in Excel character units (as stored in the file). */
  w?: number;
  customWidth?: boolean;
  hidden?: boolean;
  s?: number;
  outline?: number;
  collapsed?: boolean;
  bestFit?: boolean;
}

export interface Hyperlink {
  ref: Range;
  /** External URL (stored as a relationship). */
  target?: string;
  /** Internal location such as Sheet2!A1. */
  location?: string;
  tooltip?: string;
  display?: string;
}

export interface CommentEntry {
  id: string;
  author: string;
  /** Person id for threaded comments. */
  personId?: string;
  text: string;
  /** ISO date-time. */
  created?: string;
}

export interface CommentThread {
  id: string;
  r: number;
  c: number;
  /** 'note' = legacy Excel note, 'threaded' = Excel 365 threaded comment. */
  kind: 'note' | 'threaded';
  comments: CommentEntry[];
  resolved: boolean;
  /** Raw VML shape of a legacy note, kept when the note is unchanged. */
  vmlShape?: string;
}

export interface AnchorPoint {
  r: number;
  c: number;
  /** Offsets in EMU. */
  dx: number;
  dy: number;
}

export interface SheetImage {
  id: string;
  name: string;
  descr?: string;
  mime: string;
  data: Uint8Array;
  anchorType: 'twoCell' | 'oneCell' | 'absolute';
  from: AnchorPoint;
  to?: AnchorPoint;
  /** Size in EMU (oneCell and absolute anchors). */
  ext?: { cx: number; cy: number };
  /** Absolute position in EMU. */
  pos?: { x: number; y: number };
  editAs?: 'oneCell' | 'twoCell' | 'absolute';
  /** The original anchor element; dropped once the image is moved or resized. */
  raw?: XmlElement;
  /** Stacking order among pictures, charts and shapes. */
  z?: number;
}

export type ValidationType = 'none' | 'whole' | 'decimal' | 'list' | 'date' | 'time' | 'textLength' | 'custom';
export type ValidationOperator = 'between' | 'notBetween' | 'equal' | 'notEqual' | 'greaterThan' | 'lessThan' | 'greaterThanOrEqual' | 'lessThanOrEqual';

export interface DataValidation {
  sqref: Range[];
  type?: ValidationType | string;
  operator?: ValidationOperator | string;
  /** Formulas without '=' (a list is either a range or a quoted, comma separated text). */
  formula1?: string;
  formula2?: string;
  allowBlank?: boolean;
  /** Note: Excel's showDropDown="1" hides the in-cell drop-down. */
  hideDropDown?: boolean;
  showErrorMessage?: boolean;
  /** stop rejects the entry, warning asks, information only tells. */
  errorStyle?: 'stop' | 'warning' | 'information';
  errorTitle?: string;
  error?: string;
  showInputMessage?: boolean;
  promptTitle?: string;
  prompt?: string;
  /**
   * "Format style" action: instead of a message, invalid cells are formatted. Stored as a
   * conditional format (linked by `cfId`) so that Excel shows it too.
   */
  formatStyle?: ValidationFormatStyle;
  cfId?: string;
  /** The source element (keeps attributes the editor does not model). */
  raw?: XmlElement;
}

export interface ValidationFormatStyle {
  color?: string;
  background?: string;
  bold?: boolean;
}

export type FilterOperator =
  | 'contains' | 'notContains' | 'beginsWith' | 'endsWith' | 'equal' | 'notEqual'
  | 'greaterThan' | 'lessThan' | 'greaterThanOrEqual' | 'lessThanOrEqual';

export interface FilterCondition {
  op: FilterOperator;
  value: string;
}

export interface ColumnFilter {
  /** Displayed values to keep; undefined keeps every value. */
  values?: string[];
  /** Keep empty cells when filtering by values. */
  blanks?: boolean;
  /** Custom conditions (Excel stores up to two). */
  conditions?: FilterCondition[];
  /** Conditions joined by AND (default OR). */
  and?: boolean;
}

export interface AutoFilter {
  /** The filtered area, header row included. */
  ref: Range;
  /** Criteria by absolute column index. */
  columns: Map<number, ColumnFilter>;
  /** The source element (sort state and extensions are kept). */
  raw?: XmlElement;
}

export interface ConditionalFormat {
  sqref: Range[];
  raw: XmlElement;
  /** Set on rules created by the editor (e.g. for a validation's format style). */
  id?: string;
}

export interface FreezePanes {
  rows: number;
  cols: number;
}

export interface SheetSelection {
  active: { r: number; c: number };
  ranges: Range[];
}

/** Default column width in pixels (8.43 characters of Calibri 11). */
export const DEFAULT_COL_WIDTH_PX = 64;
/** Default row height in points. */
export const DEFAULT_ROW_HEIGHT_PT = 15;
/** Maximum digit width in pixels of the default font, used for column width units. */
export const MDW = 7;

export function widthCharsToPx(w: number): number {
  return Math.floor(((256 * w + Math.floor(128 / MDW)) / 256) * MDW);
}

/** Pixels to the stored column width, using Excel's own rounding (64px -> 9.140625). */
export function widthPxToChars(px: number): number {
  if (px <= 5) return Math.trunc((Math.max(0, px) / MDW) * 256) / 256;
  const chars = Math.trunc(((px - 5) / MDW) * 100 + 0.5) / 100;
  return Math.trunc(((chars * MDW + 5) / MDW) * 256) / 256;
}

/** New index of a row/column after inserting (delta > 0) or deleting (delta < 0) before `from`. */
export function shiftIndex(i: number, from: number, delta: number): number {
  if (i >= from) return i + delta;
  if (delta < 0 && i >= from + delta) return from + delta;
  return i;
}

/** True when index `i` lies in the lines removed by a deletion described like `shiftIndex`. */
export function isDeletedIndex(i: number, from: number, delta: number): boolean {
  return delta < 0 && i >= from + delta && i < from;
}

export function ptToPx(pt: number): number {
  return Math.round((pt * 96) / 72);
}

export function pxToPt(px: number): number {
  return Math.round(((px * 72) / 96) * 4) / 4;
}

let uidCounter = 1;

export class Sheet {
  readonly uid: number;
  name: string;
  /** sheetId in workbook.xml. */
  sheetId: number;
  state: 'visible' | 'hidden' | 'veryHidden' = 'visible';
  readonly cells = new Map<number, Map<number, Cell>>();
  readonly rows = new Map<number, RowInfo>();
  readonly cols = new Map<number, ColInfo>();
  /** Default column width in pixels. */
  defaultColWidthPx = DEFAULT_COL_WIDTH_PX;
  /** Default row height in points. */
  defaultRowHeightPt = DEFAULT_ROW_HEIGHT_PT;
  merges: Range[] = [];
  freeze: FreezePanes = { rows: 0, cols: 0 };
  zoom = 1;
  showGridLines = true;
  rightToLeft = false;
  tabColor?: ColorSpec;
  selection?: SheetSelection;
  hyperlinks: Hyperlink[] = [];
  comments: CommentThread[] = [];
  images: SheetImage[] = [];
  validations: DataValidation[] = [];
  condFormats: ConditionalFormat[] = [];
  autoFilter?: AutoFilter;
  /** Rows hidden by the auto filter (as opposed to hidden by the user). */
  filteredRows = new Set<number>();
  charts: SheetChart[] = [];
  shapes: SheetShape[] = [];
  /** Footer row shown under the grid (an editor feature, not stored in XLSX). */
  footer?: string[];
  /** Custom column header titles (an editor feature, not stored in XLSX). */
  colTitles?: Map<number, string>;
  /** Raw worksheet XML of the source file, used to keep unknown content on save. */
  sourceXml?: XmlElement;
  /** Path of the worksheet part in the source package. */
  sourcePath?: string;
  /** A chart sheet from the source file: kept as is, shown read-only. */
  chartsheet?: boolean;
  /** Drawing anchors that are not pictures (charts, shapes), kept unchanged. */
  otherAnchors: XmlElement[] = [];
  /** Incremented whenever row heights or column widths change (layout caches use it). */
  layoutVersion = 0;
  private hyperlinkIndex: Map<number, Hyperlink> | null = null;
  private mergeIndex: Map<number, Range> | null = null;

  constructor(name: string, sheetId = 1) {
    this.uid = uidCounter++;
    this.name = name;
    this.sheetId = sheetId;
  }

  // ---------------------------------------------------------------------------------------------
  // Cells

  getCell(r: number, c: number): Cell | undefined {
    return this.cells.get(r)?.get(c);
  }

  setCell(r: number, c: number, cell: Cell | undefined): void {
    let row = this.cells.get(r);
    if (!cell || (cell.v === null && !cell.f && !cell.s && !cell.rt && cell.spill === undefined)) {
      if (row) {
        row.delete(c);
        if (!row.size) this.cells.delete(r);
      }
      return;
    }
    if (!row) {
      row = new Map();
      this.cells.set(r, row);
    }
    row.set(c, cell);
  }

  value(r: number, c: number): Scalar {
    return this.getCell(r, c)?.v ?? null;
  }

  /** Style index of a cell, falling back to the row and column styles like Excel does. */
  styleAt(r: number, c: number): number {
    const cell = this.getCell(r, c);
    if (cell && cell.s !== undefined) return cell.s;
    const row = this.rows.get(r);
    if (row?.s !== undefined) return row.s;
    return this.cols.get(c)?.s ?? 0;
  }

  /** Visits existing cells inside a range, row by row. */
  forEachCellIn(range: Range, fn: (r: number, c: number, cell: Cell) => void): void {
    const rowCount = range.r1 - range.r0 + 1;
    if (rowCount <= this.cells.size) {
      for (let r = range.r0; r <= range.r1; r++) {
        const row = this.cells.get(r);
        if (row) this.visitRow(r, row, range, fn);
      }
      return;
    }
    const rows = [...this.cells.keys()].filter((r) => r >= range.r0 && r <= range.r1).sort((a, b) => a - b);
    for (const r of rows) this.visitRow(r, this.cells.get(r)!, range, fn);
  }

  private visitRow(r: number, row: Map<number, Cell>, range: Range, fn: (r: number, c: number, cell: Cell) => void) {
    const colCount = range.c1 - range.c0 + 1;
    if (colCount <= row.size) {
      for (let c = range.c0; c <= range.c1; c++) {
        const cell = row.get(c);
        if (cell) fn(r, c, cell);
      }
    } else {
      const cols = [...row.keys()].filter((c) => c >= range.c0 && c <= range.c1).sort((a, b) => a - b);
      for (const c of cols) fn(r, c, row.get(c)!);
    }
  }

  forEachCell(fn: (r: number, c: number, cell: Cell) => void): void {
    for (const [r, row] of this.cells) for (const [c, cell] of row) fn(r, c, cell);
  }

  /** The smallest range containing every non-empty cell (and merges), or null for an empty sheet. */
  usedRange(includeStyles = true): Range | null {
    let r0 = Infinity;
    let c0 = Infinity;
    let r1 = -1;
    let c1 = -1;
    for (const [r, row] of this.cells) {
      for (const [c, cell] of row) {
        if (!includeStyles && cell.v === null && !cell.f) continue;
        if (r < r0) r0 = r;
        if (r > r1) r1 = r;
        if (c < c0) c0 = c;
        if (c > c1) c1 = c;
      }
    }
    for (const m of this.merges) {
      r0 = Math.min(r0, m.r0);
      c0 = Math.min(c0, m.c0);
      r1 = Math.max(r1, m.r1);
      c1 = Math.max(c1, m.c1);
    }
    if (r1 < 0) return null;
    return { r0, c0, r1, c1 };
  }

  /** Last row and column that contain data (for whole-column references), -1 if none. */
  dataExtent(): { rows: number; cols: number } {
    let rows = -1;
    let cols = -1;
    for (const [r, row] of this.cells) {
      if (r > rows) rows = r;
      for (const c of row.keys()) if (c > cols) cols = c;
    }
    return { rows, cols };
  }

  // ---------------------------------------------------------------------------------------------
  // Rows and columns

  rowHeightPt(r: number): number {
    const info = this.rows.get(r);
    if (info?.hidden) return 0;
    return info?.h ?? this.defaultRowHeightPt;
  }

  rowHeightPx(r: number): number {
    return ptToPx(this.rowHeightPt(r));
  }

  colWidthPx(c: number): number {
    const info = this.cols.get(c);
    if (info?.hidden) return 0;
    if (info?.w !== undefined) return widthCharsToPx(info.w);
    return this.defaultColWidthPx;
  }

  isRowHidden(r: number): boolean {
    return !!this.rows.get(r)?.hidden;
  }

  isColHidden(c: number): boolean {
    return !!this.cols.get(c)?.hidden;
  }

  setRowInfo(r: number, info: RowInfo | undefined): void {
    this.layoutVersion++;
    if (!info || Object.keys(info).every((k) => (info as Record<string, unknown>)[k] === undefined)) this.rows.delete(r);
    else this.rows.set(r, info);
  }

  setColInfo(c: number, info: ColInfo | undefined): void {
    this.layoutVersion++;
    if (!info || Object.keys(info).every((k) => (info as Record<string, unknown>)[k] === undefined)) this.cols.delete(c);
    else this.cols.set(c, info);
  }

  // ---------------------------------------------------------------------------------------------
  // Merges

  mergeAt(r: number, c: number): Range | undefined {
    if (!this.merges.length) return undefined;
    if (!this.mergeIndex) {
      this.mergeIndex = new Map();
      for (const m of this.merges) {
        if ((m.r1 - m.r0 + 1) * (m.c1 - m.c0 + 1) > 100000) continue;
        for (let rr = m.r0; rr <= m.r1; rr++) for (let cc = m.c0; cc <= m.c1; cc++) this.mergeIndex.set(cellKey(rr, cc), m);
      }
    }
    const hit = this.mergeIndex.get(cellKey(r, c));
    if (hit) return hit;
    for (const m of this.merges) {
      if ((m.r1 - m.r0 + 1) * (m.c1 - m.c0 + 1) > 100000 && rangeContains(m, r, c)) return m;
    }
    return undefined;
  }

  mergesIntersecting(range: Range): Range[] {
    return this.merges.filter((m) => rangesIntersect(m, range));
  }

  setMerges(merges: Range[]): void {
    this.merges = merges;
    this.mergeIndex = null;
  }

  /** Expands a range so it fully contains every merge it touches. */
  expandToMerges(range: Range): Range {
    let r = { ...range };
    let changed = true;
    while (changed) {
      changed = false;
      for (const m of this.merges) {
        if (!rangesIntersect(m, r)) continue;
        if (m.r0 < r.r0 || m.c0 < r.c0 || m.r1 > r.r1 || m.c1 > r.c1) {
          r = { r0: Math.min(r.r0, m.r0), c0: Math.min(r.c0, m.c0), r1: Math.max(r.r1, m.r1), c1: Math.max(r.c1, m.c1) };
          changed = true;
        }
      }
    }
    return r;
  }

  // ---------------------------------------------------------------------------------------------
  // Hyperlinks and comments

  hyperlinkAt(r: number, c: number): Hyperlink | undefined {
    if (!this.hyperlinks.length) return undefined;
    if (!this.hyperlinkIndex) {
      this.hyperlinkIndex = new Map();
      for (const h of this.hyperlinks) {
        if ((h.ref.r1 - h.ref.r0 + 1) * (h.ref.c1 - h.ref.c0 + 1) > 10000) continue;
        for (let rr = h.ref.r0; rr <= h.ref.r1; rr++) for (let cc = h.ref.c0; cc <= h.ref.c1; cc++) this.hyperlinkIndex.set(cellKey(rr, cc), h);
      }
    }
    return this.hyperlinkIndex.get(cellKey(r, c)) ?? this.hyperlinks.find((h) => rangeContains(h.ref, r, c));
  }

  setHyperlinks(list: Hyperlink[]): void {
    this.hyperlinks = list;
    this.hyperlinkIndex = null;
  }

  commentAt(r: number, c: number): CommentThread | undefined {
    return this.comments.find((t) => t.r === r && t.c === c);
  }

  /** Marks derived indexes stale after structural changes. */
  invalidateIndexes(): void {
    this.hyperlinkIndex = null;
    this.mergeIndex = null;
  }

  // ---------------------------------------------------------------------------------------------
  // Structural changes (cells and per-sheet metadata only; formulas are adjusted by the engine)

  insertRows(at: number, count: number): void {
    this.shiftRows(at, count);
  }

  deleteRows(at: number, count: number): void {
    for (let r = at; r < at + count; r++) {
      this.cells.delete(r);
      this.rows.delete(r);
    }
    this.shiftRows(at + count, -count);
  }

  insertCols(at: number, count: number): void {
    this.shiftCols(at, count);
  }

  deleteCols(at: number, count: number): void {
    for (const row of this.cells.values()) for (let c = at; c < at + count; c++) row.delete(c);
    for (let c = at; c < at + count; c++) this.cols.delete(c);
    this.shiftCols(at + count, -count);
  }

  private shiftRows(from: number, delta: number): void {
    this.layoutVersion++;
    const moveMap = <T>(map: Map<number, T>) => {
      const entries = [...map.entries()].filter(([r]) => r >= from);
      for (const [r] of entries) map.delete(r);
      for (const [r, v] of entries) {
        const nr = r + delta;
        if (nr >= 0 && nr < MAX_ROWS) map.set(nr, v);
      }
    };
    moveMap(this.cells);
    moveMap(this.rows);
    for (const [, row] of this.cells) {
      for (const [c, cell] of row) {
        if (cell.spill !== undefined) {
          const sr = keyRow(cell.spill);
          if (sr >= from) row.set(c, { ...cell, spill: cellKey(sr + delta, keyCol(cell.spill)) });
        }
      }
    }
    const adjust = (r: Range): Range | null => shiftRange(r, 'row', from, delta);
    this.setMerges(this.merges.map(adjust).filter((m): m is Range => !!m && !(m.r0 === m.r1 && m.c0 === m.c1)));
    this.setHyperlinks(this.hyperlinks.flatMap((h) => { const ref = adjust(h.ref); return ref ? [{ ...h, ref }] : []; }));
    this.comments = this.comments.flatMap((t) => (isDeletedIndex(t.r, from, delta) ? [] : [{ ...t, r: shiftIndex(t.r, from, delta) }]));
    for (const img of this.images) {
      const nr = shiftIndex(img.from.r, from, delta);
      if (nr !== img.from.r) img.from = { ...img.from, r: nr, dy: isDeletedIndex(img.from.r, from, delta) ? 0 : img.from.dy };
      if (img.to) {
        const tr = shiftIndex(img.to.r, from, delta);
        if (tr !== img.to.r) img.to = { ...img.to, r: Math.max(tr, img.from.r), dy: isDeletedIndex(img.to.r, from, delta) ? 0 : img.to.dy };
      }
      img.raw = undefined;
    }
    for (const obj of [...this.charts, ...this.shapes]) shiftObjectAnchor(obj, 'row', from, delta);
    this.validations = this.validations.flatMap((v) => {
      const sq = v.sqref.map(adjust).filter((x): x is Range => !!x);
      return sq.length ? [{ ...v, sqref: sq }] : [];
    });
    this.condFormats = this.condFormats.flatMap((cf) => {
      const sq = cf.sqref.map(adjust).filter((x): x is Range => !!x);
      return sq.length ? [{ ...cf, sqref: sq }] : [];
    });
    if (this.autoFilter) {
      const ref = adjust(this.autoFilter.ref);
      this.autoFilter = ref ? { ...this.autoFilter, ref } : undefined;
    }
    this.filteredRows = new Set([...this.filteredRows].filter((r) => !isDeletedIndex(r, from, delta)).map((r) => shiftIndex(r, from, delta)));
    this.invalidateIndexes();
  }

  private shiftCols(from: number, delta: number): void {
    this.layoutVersion++;
    for (const [r, row] of this.cells) {
      const entries = [...row.entries()].filter(([c]) => c >= from);
      for (const [c] of entries) row.delete(c);
      for (const [c, v] of entries) {
        const nc = c + delta;
        if (nc >= 0 && nc < MAX_COLS) row.set(nc, v);
      }
      for (const [c, cell] of row) {
        if (cell.spill !== undefined) {
          const sc = keyCol(cell.spill);
          if (sc >= from) row.set(c, { ...cell, spill: cellKey(keyRow(cell.spill), sc + delta) });
        }
      }
      if (!row.size) this.cells.delete(r);
    }
    const colEntries = [...this.cols.entries()].filter(([c]) => c >= from);
    for (const [c] of colEntries) this.cols.delete(c);
    for (const [c, v] of colEntries) {
      const nc = c + delta;
      if (nc >= 0 && nc < MAX_COLS) this.cols.set(nc, v);
    }
    const adjust = (r: Range): Range | null => shiftRange(r, 'col', from, delta);
    this.setMerges(this.merges.map(adjust).filter((m): m is Range => !!m && !(m.r0 === m.r1 && m.c0 === m.c1)));
    this.setHyperlinks(this.hyperlinks.flatMap((h) => { const ref = adjust(h.ref); return ref ? [{ ...h, ref }] : []; }));
    this.comments = this.comments.flatMap((t) => (isDeletedIndex(t.c, from, delta) ? [] : [{ ...t, c: shiftIndex(t.c, from, delta) }]));
    for (const img of this.images) {
      const nc = shiftIndex(img.from.c, from, delta);
      if (nc !== img.from.c) img.from = { ...img.from, c: nc, dx: isDeletedIndex(img.from.c, from, delta) ? 0 : img.from.dx };
      if (img.to) {
        const tc = shiftIndex(img.to.c, from, delta);
        if (tc !== img.to.c) img.to = { ...img.to, c: Math.max(tc, img.from.c), dx: isDeletedIndex(img.to.c, from, delta) ? 0 : img.to.dx };
      }
      img.raw = undefined;
    }
    for (const obj of [...this.charts, ...this.shapes]) shiftObjectAnchor(obj, 'col', from, delta);
    this.validations = this.validations.flatMap((v) => {
      const sq = v.sqref.map(adjust).filter((x): x is Range => !!x);
      return sq.length ? [{ ...v, sqref: sq }] : [];
    });
    this.condFormats = this.condFormats.flatMap((cf) => {
      const sq = cf.sqref.map(adjust).filter((x): x is Range => !!x);
      return sq.length ? [{ ...cf, sqref: sq }] : [];
    });
    if (this.autoFilter) {
      const ref = adjust(this.autoFilter.ref);
      if (ref) {
        const columns = new Map<number, ColumnFilter>();
        for (const [c, f] of this.autoFilter.columns) if (!isDeletedIndex(c, from, delta)) columns.set(shiftIndex(c, from, delta), f);
        this.autoFilter = { ...this.autoFilter, ref, columns };
      } else this.autoFilter = undefined;
    }
    if (this.colTitles) {
      const titles = new Map<number, string>();
      for (const [c, t] of this.colTitles) if (!isDeletedIndex(c, from, delta)) titles.set(shiftIndex(c, from, delta), t);
      this.colTitles = titles;
    }
    if (this.footer) {
      const next = [...this.footer];
      if (delta > 0) next.splice(from, 0, ...new Array(delta).fill(''));
      else next.splice(from + delta, -delta);
      this.footer = next;
    }
    this.invalidateIndexes();
  }
}

/** Moves a chart or shape anchor for inserted (delta > 0) or deleted (delta < 0) rows or columns. */
function shiftObjectAnchor(obj: { from: AnchorPoint; to: AnchorPoint; source?: unknown }, axis: 'row' | 'col', from: number, delta: number): void {
  const key = axis === 'row' ? 'r' : 'c';
  const off = axis === 'row' ? 'dy' : 'dx';
  const a = shiftIndex(obj.from[key], from, delta);
  const b = shiftIndex(obj.to[key], from, delta);
  if (a === obj.from[key] && b === obj.to[key]) return;
  obj.from = { ...obj.from, [key]: a, [off]: isDeletedIndex(obj.from[key], from, delta) ? 0 : obj.from[off] };
  obj.to = { ...obj.to, [key]: Math.max(b, a), [off]: isDeletedIndex(obj.to[key], from, delta) ? 0 : obj.to[off] };
}

/**
 * Shifts a range for an insertion (delta > 0 at `from`) or a deletion (delta < 0 of the
 * lines `from + delta .. from - 1`). Returns null when the range is deleted entirely.
 */
export function shiftRange(range: Range, axis: 'row' | 'col', from: number, delta: number): Range | null {
  const lo = axis === 'row' ? range.r0 : range.c0;
  const hi = axis === 'row' ? range.r1 : range.c1;
  const max = (axis === 'row' ? MAX_ROWS : MAX_COLS) - 1;
  let nlo = lo;
  let nhi = hi;
  if (delta > 0) {
    if (lo >= from) nlo = lo + delta;
    if (hi >= from) nhi = hi + delta;
    if (lo === 0 && hi === max) {
      nlo = 0;
      nhi = max;
    }
  } else {
    const delStart = from + delta;
    const delEnd = from - 1;
    if (lo === 0 && hi === max) {
      nlo = 0;
      nhi = max;
    } else {
      if (lo >= delStart && hi <= delEnd) return null;
      nlo = lo > delEnd ? lo + delta : lo >= delStart ? delStart : lo;
      nhi = hi > delEnd ? hi + delta : hi >= delStart ? delStart - 1 : hi;
      if (nhi < nlo) return null;
    }
  }
  nhi = Math.min(nhi, max);
  return axis === 'row' ? { ...range, r0: nlo, r1: nhi } : { ...range, c0: nlo, c1: nhi };
}
