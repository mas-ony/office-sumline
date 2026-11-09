/**
 * The editing controller. Every change goes through a transaction (undo/redo), then the formula
 * engine recalculates and listeners (the grid, the toolbar, the compatibility API) are notified.
 */
import {
  MAX_COLS, MAX_ROWS, cellName, isFullColumns, isFullRows, makeRange, parseCellRef, parseRange, rangeContains, rangeName,
  rangeEquals, rangesIntersect, splitSheetRef, type Range,
} from './refs';
import { CellError, type Scalar } from './values';
import { Workbook, validateSheetName } from './model/workbook';
import { Sheet, ptToPx, pxToPt, widthPxToChars, type Cell, type ColumnFilter, type CommentThread, type DataValidation, type SheetImage, type SheetSelection } from './model/sheet';
import { newDrawingId, type SheetChart, type SheetShape } from './model/drawing';
import { columnValues, filterRegion, rowsToHide } from './filter';
import { listItems, satisfies, validationAt, validityFormula } from './validation';
import { el as xmlEl, txt as xmlText, type XmlElement } from './xml/xml';
import { formulaToFile } from './xlsx/formulaText';
import { applyPatch, type BorderLine, type CellStyle, type StylePatch } from './model/style';
import { CalcEngine, cellUid } from './formula/engine';
import { History, Transaction, restoreSheet, restoreWorkbook, type SelectionState } from './history';
import { parseInput } from './format/input';
import { editText, formatInfo, formatValue } from './format/display';
import { normalizeFormulaInput } from './format/formulaInput';
import {
  deleteSheetInFormula, moveRefsInFormula, renameSheetInFormula, shiftFormula, translateFormula,
} from './formula/adjust';
import { fillSeries } from './autofill';
import { parseDelimited, parseHtmlTable, toHtml, toTsv, type HtmlCellStyle, type HtmlOutCell } from './clipboard';
import { readXlsx } from './xlsx/import';
import { writeXlsx, newGuid } from './xlsx/export';
import { SheetLayout } from './layout';
import { resolveColor, rgbSpec } from './model/color';
import { pxToEmu } from './xlsx/util';
import { compareScalars } from './formula/runtime';

export type EditMode = 'editing' | 'viewOnly';

export interface Selection {
  active: { r: number; c: number };
  anchor: { r: number; c: number };
  ranges: Range[];
}

export type BorderKind = 'all' | 'outside' | 'inside' | 'horizontal' | 'vertical' | 'top' | 'bottom' | 'left' | 'right' | 'none';

export interface ChangeEvent {
  sheet?: Sheet;
  /** Structural change: rows, columns, sheets or merges changed. */
  structure: boolean;
  /** Row heights or column widths changed. */
  layout: boolean;
  /** The transaction (absent for recalculation-only updates). */
  tx?: Transaction;
  /** Cells changed by the user, as ranges per sheet. */
  ranges: { sheet: Sheet; range: Range }[];
}

export interface SearchOptions {
  matchCase?: boolean;
  wholeCell?: boolean;
  allSheets?: boolean;
  /** Search formulas instead of displayed values. */
  inFormulas?: boolean;
}

export interface SearchHit {
  sheet: Sheet;
  r: number;
  c: number;
  text: string;
}

export interface ClipPayload {
  text: string;
  html: string;
}

interface InternalClip {
  sheet: Sheet;
  range: Range;
  cells: (Cell | undefined)[][];
  merges: Range[];
  text: string;
  cut: boolean;
  colWidths: number[];
}

type Listener<T> = (e: T) => void;

interface EventMap {
  change: ChangeEvent;
  selection: { sheet: Sheet; selection: Selection };
  activeSheet: { index: number };
  sheets: { action: string; index?: number };
  mode: EditMode;
  history: { canUndo: boolean; canRedo: boolean };
  comments: { sheet?: Sheet };
  file: { name: string };
}

/** Measures text for auto-fit, provided by the UI (canvas based). */
export type TextMeasurer = (text: string, style: CellStyle) => { width: number; height: number };

/** Stand-in home sheet for chart references (they always name their sheet). */
const CHART_HOME = { name: '' } as unknown as Sheet;

function clone2d<T>(rows: number, cols: number, f: (r: number, c: number) => T): T[][] {
  return Array.from({ length: rows }, (_, r) => Array.from({ length: cols }, (_, c) => f(r, c)));
}

export class SpreadsheetDocument {
  wb: Workbook;
  engine: CalcEngine;
  readonly history = new History();
  mode: EditMode = 'editing';
  fileName: string;
  author = 'Guest';
  /** Unsaved changes. */
  modified = false;
  measure?: TextMeasurer;
  private selections = new WeakMap<Sheet, Selection>();
  private listeners = new Map<keyof EventMap, Set<Listener<unknown>>>();
  private clip?: InternalClip;
  private batch: Transaction | null = null;

  constructor(wb: Workbook, fileName = 'Untitled.xlsx') {
    this.wb = wb;
    this.fileName = fileName;
    this.engine = new CalcEngine(wb);
    this.engine.rebuild();
    this.engine.recalc();
    for (const s of wb.sheets) this.initSelection(s);
  }

  static blank(fileName = 'Untitled.xlsx'): SpreadsheetDocument {
    return new SpreadsheetDocument(Workbook.blank(), fileName);
  }

  static fromXlsx(bytes: Uint8Array, fileName = 'Workbook.xlsx'): SpreadsheetDocument {
    return new SpreadsheetDocument(readXlsx(bytes, { fileName }), fileName);
  }

  // ---------------------------------------------------------------------------------------------
  // Events

  on<K extends keyof EventMap>(event: K, fn: Listener<EventMap[K]>): () => void {
    let set = this.listeners.get(event);
    if (!set) this.listeners.set(event, (set = new Set()));
    set.add(fn as Listener<unknown>);
    return () => set!.delete(fn as Listener<unknown>);
  }

  emit<K extends keyof EventMap>(event: K, payload: EventMap[K]): void {
    for (const fn of this.listeners.get(event) ?? []) {
      try {
        (fn as Listener<EventMap[K]>)(payload);
      } catch (e) {
        console.error(e);
      }
    }
  }

  // ---------------------------------------------------------------------------------------------
  // Sheets and selection

  get sheet(): Sheet {
    return this.wb.activeSheet;
  }

  get activeIndex(): number {
    return this.wb.activeSheetIndex;
  }

  private initSelection(sheet: Sheet): Selection {
    const saved = sheet.selection;
    const active = saved?.active ?? { r: sheet.freeze.rows ? sheet.freeze.rows : 0, c: 0 };
    const sel: Selection = { active: { ...active }, anchor: { ...active }, ranges: saved?.ranges.length ? saved.ranges.map((r) => ({ ...r })) : [makeRange(active.r, active.c)] };
    if (!sel.ranges.some((r) => rangeContains(r, active.r, active.c))) sel.ranges = [makeRange(active.r, active.c)];
    this.selections.set(sheet, sel);
    return sel;
  }

  selectionOf(sheet: Sheet = this.sheet): Selection {
    return this.selections.get(sheet) ?? this.initSelection(sheet);
  }

  get selection(): Selection {
    return this.selectionOf(this.sheet);
  }

  /** The primary (last) selected range. */
  get range(): Range {
    const s = this.selection;
    return s.ranges[s.ranges.length - 1];
  }

  setActiveSheet(index: number): void {
    if (index < 0 || index >= this.wb.sheets.length || index === this.wb.activeSheetIndex) return;
    this.wb.activeSheetIndex = index;
    this.emit('activeSheet', { index });
    this.emitSelection();
  }

  setSelection(sel: Selection, sheet: Sheet = this.sheet): void {
    const clean: Selection = {
      active: { r: clampRow(sel.active.r), c: clampCol(sel.active.c) },
      anchor: { r: clampRow(sel.anchor.r), c: clampCol(sel.anchor.c) },
      ranges: sel.ranges.map((r) => makeRange(clampRow(r.r0), clampCol(r.c0), clampRow(r.r1), clampCol(r.c1))),
    };
    this.selections.set(sheet, clean);
    sheet.selection = { active: { ...clean.active }, ranges: clean.ranges.map((r) => ({ ...r })) };
    if (sheet === this.sheet) this.emitSelection();
  }

  private emitSelection(): void {
    this.emit('selection', { sheet: this.sheet, selection: this.selection });
  }

  /** Selects a range; the active cell defaults to its top-left corner. */
  select(range: Range, active?: { r: number; c: number }, add = false): void {
    const sheet = this.sheet;
    const r = sheet.expandToMerges(makeRange(range.r0, range.c0, range.r1, range.c1));
    const a = active ?? { r: r.r0, c: r.c0 };
    const ranges = add ? [...this.selection.ranges, r] : [r];
    this.setSelection({ active: a, anchor: { ...a }, ranges });
  }

  selectCell(r: number, c: number): void {
    const m = this.sheet.mergeAt(r, c);
    this.select(m ?? makeRange(r, c), m ? { r: m.r0, c: m.c0 } : { r, c });
  }

  /** Extends the primary range from the anchor to (r, c), keeping the active cell. */
  extendTo(r: number, c: number): void {
    const sel = this.selection;
    const range = this.sheet.expandToMerges(makeRange(sel.anchor.r, sel.anchor.c, r, c));
    const ranges = [...sel.ranges.slice(0, -1), range];
    this.setSelection({ active: sel.active, anchor: sel.anchor, ranges });
  }

  selectAll(): void {
    const used = this.sheet.usedRange(false);
    const sel = this.selection;
    const all = makeRange(0, 0, MAX_ROWS - 1, MAX_COLS - 1);
    // Ctrl+A first selects the current data region, then the whole sheet (like Excel).
    if (used && !rangeEquals(this.range, used) && rangeContains(used, sel.active.r, sel.active.c)) {
      this.setSelection({ active: sel.active, anchor: { r: used.r0, c: used.c0 }, ranges: [used] });
      return;
    }
    this.setSelection({ active: sel.active, anchor: { r: 0, c: 0 }, ranges: [all] });
  }

  selectRows(r0: number, r1: number, add = false): void {
    const range = makeRange(r0, 0, r1, MAX_COLS - 1);
    const active = { r: Math.min(r0, r1), c: 0 };
    this.setSelection({ active, anchor: { r: r0, c: 0 }, ranges: add ? [...this.selection.ranges, range] : [range] });
  }

  selectCols(c0: number, c1: number, add = false): void {
    const range = makeRange(0, c0, MAX_ROWS - 1, c1);
    const active = { r: 0, c: Math.min(c0, c1) };
    this.setSelection({ active, anchor: { r: 0, c: c0 }, ranges: add ? [...this.selection.ranges, range] : [range] });
  }

  /** Arrow-key navigation. `toEdge` jumps to the edge of the data region (Ctrl+arrow). */
  move(dr: number, dc: number, extend = false, toEdge = false): void {
    const sel = this.selection;
    const sheet = this.sheet;
    const from = extend ? this.extendCursor() : sel.active;
    let r = from.r;
    let c = from.c;
    if (toEdge) [r, c] = this.edgeFrom(r, c, dr, dc);
    else {
      // Step over merged cells and hidden lines.
      const m = sheet.mergeAt(r, c);
      if (m && !extend) {
        if (dr > 0) r = m.r1;
        if (dc > 0) c = m.c1;
        if (dr < 0) r = m.r0;
        if (dc < 0) c = m.c0;
      }
      do {
        r = clampRow(r + dr);
        c = clampCol(c + dc);
      } while ((dr && sheet.isRowHidden(r) && r > 0 && r < MAX_ROWS - 1) || (dc && sheet.isColHidden(c) && c > 0 && c < MAX_COLS - 1));
    }
    if (extend) {
      this.extendTo(r, c);
      this.extendCursorPos = { r, c };
    } else {
      this.extendCursorPos = null;
      this.selectCell(r, c);
    }
  }

  private extendCursorPos: { r: number; c: number } | null = null;

  private extendCursor(): { r: number; c: number } {
    const sel = this.selection;
    if (this.extendCursorPos) return this.extendCursorPos;
    const range = sel.ranges[sel.ranges.length - 1];
    return { r: range.r0 === sel.anchor.r ? range.r1 : range.r0, c: range.c0 === sel.anchor.c ? range.c1 : range.c0 };
  }

  private filled(r: number, c: number): boolean {
    const v = this.sheet.getCell(r, c);
    return !!v && (v.v !== null || v.f !== undefined);
  }

  private edgeFrom(r: number, c: number, dr: number, dc: number): [number, number] {
    const ext = this.sheet.dataExtent();
    const maxR = Math.max(ext.rows, r);
    const maxC = Math.max(ext.cols, c);
    const step = () => {
      r += dr;
      c += dc;
    };
    const inside = () => r >= 0 && c >= 0 && r <= maxR + 1 && c <= maxC + 1;
    const startFilled = this.filled(r, c);
    step();
    if (!inside()) return [clampRow(r - dr), clampCol(c - dc)];
    if (startFilled && this.filled(r, c)) {
      while (inside() && this.filled(r + dr, c + dc)) step();
      return [clampRow(r), clampCol(c)];
    }
    while (inside() && !this.filled(r, c)) {
      if ((dr > 0 && r > maxR) || (dc > 0 && c > maxC)) return [dr > 0 ? MAX_ROWS - 1 : r, dc > 0 ? MAX_COLS - 1 : c];
      step();
    }
    if (!inside() || r < 0 || c < 0) return [Math.max(0, clampRow(r)), Math.max(0, clampCol(c))];
    return [clampRow(r), clampCol(c)];
  }

  /** Moves the active cell inside the selection (Enter / Tab in a multi-cell selection). */
  cycleActive(dr: number, dc: number): void {
    const sel = this.selection;
    const range = this.range;
    if (range.r0 === range.r1 && range.c0 === range.c1 && sel.ranges.length === 1) {
      this.move(dr, dc);
      return;
    }
    let { r, c } = sel.active;
    if (dr) {
      r += dr;
      if (r > range.r1) {
        r = range.r0;
        c = c + 1 > range.c1 ? range.c0 : c + 1;
      } else if (r < range.r0) {
        r = range.r1;
        c = c - 1 < range.c0 ? range.c1 : c - 1;
      }
    } else {
      c += dc;
      if (c > range.c1) {
        c = range.c0;
        r = r + 1 > range.r1 ? range.r0 : r + 1;
      } else if (c < range.c0) {
        c = range.c1;
        r = r - 1 < range.r0 ? range.r1 : r - 1;
      }
    }
    this.setSelection({ active: { r, c }, anchor: sel.anchor, ranges: sel.ranges });
  }

  /** Navigates to an A1 reference or range, optionally on another sheet. */
  goTo(text: string): boolean {
    const { sheet: sheetName, ref } = splitSheetRef(text.trim());
    let range = parseRange(ref);
    if (!range) {
      const name = this.wb.nameLookup(ref, this.activeIndex);
      if (!name) return false;
      return this.goTo(name.formula);
    }
    if (sheetName !== undefined) {
      const target = this.wb.sheetByName(sheetName);
      if (!target) return false;
      this.setActiveSheet(this.wb.sheets.indexOf(target));
    }
    range = makeRange(range.r0, range.c0, range.r1, range.c1);
    this.select(range);
    return true;
  }

  // ---------------------------------------------------------------------------------------------
  // Reading

  cell(r: number, c: number, sheet: Sheet = this.sheet): Cell | undefined {
    return sheet.getCell(r, c);
  }

  styleAt(r: number, c: number, sheet: Sheet = this.sheet): CellStyle {
    return this.wb.styles.get(sheet.styleAt(r, c));
  }

  displayText(r: number, c: number, sheet: Sheet = this.sheet): string {
    const cell = sheet.getCell(r, c);
    if (!cell) return '';
    return formatValue(cell.v, this.styleAt(r, c, sheet).numFmt, this.wb.date1904).text;
  }

  /** Text shown in the formula bar and the cell editor. */
  editText(r: number, c: number, sheet: Sheet = this.sheet): string {
    const cell = sheet.getCell(r, c);
    if (!cell) return '';
    if (cell.f !== undefined) return `=${cell.f}`;
    const style = this.styleAt(r, c, sheet);
    const text = editText(cell.v, style.numFmt, this.wb.date1904);
    return style.quotePrefix ? `'${text}` : text;
  }

  /** Formula of the array/spill a cell belongs to, for the formula bar. */
  anchorOf(r: number, c: number, sheet: Sheet = this.sheet): { r: number; c: number } | null {
    const cell = sheet.getCell(r, c);
    if (cell?.spill === undefined) return null;
    return { r: Math.floor(cell.spill / MAX_COLS), c: cell.spill % MAX_COLS };
  }

  isReadOnly(): boolean {
    return this.mode === 'viewOnly';
  }

  // ---------------------------------------------------------------------------------------------
  // Transactions

  /** Runs `fn` as one undoable step. Nested calls join the outer transaction. */
  transact(label: string, fn: (tx: Transaction) => void, opts: { kind?: string; allowReadOnly?: boolean } = {}): Transaction | null {
    if (this.mode === 'viewOnly' && !opts.allowReadOnly) return null;
    if (this.batch) {
      fn(this.batch);
      if (opts.kind) this.batch.kind = opts.kind;
      return this.batch;
    }
    const tx = new Transaction(label);
    if (opts.kind) tx.kind = opts.kind;
    tx.selectionBefore = this.selectionState();
    this.batch = tx;
    try {
      fn(tx);
    } finally {
      this.batch = null;
    }
    if (tx.isEmpty()) return null;
    this.autoFitTouchedRows(tx);
    tx.selectionAfter = this.selectionState();
    tx.finish(this.wb);
    this.history.push(tx);
    this.afterChange(tx, false);
    return tx;
  }

  private selectionState(): SelectionState {
    return { sheetIndex: this.activeIndex, selection: this.sheet.selection ? { active: { ...this.sheet.selection.active }, ranges: this.sheet.selection.ranges.map((r) => ({ ...r })) } : undefined };
  }

  private afterChange(tx: Transaction, replay: boolean): void {
    const ranges: ChangeEvent['ranges'] = [];
    if (tx.structural) this.engine.invalidateAll();
    else {
      for (const m of tx.cells.values()) {
        for (const ch of m.values()) {
          this.engine.cellChanged(ch.sheet, ch.r, ch.c);
          ranges.push({ sheet: ch.sheet, range: makeRange(ch.r, ch.c) });
        }
      }
    }
    this.engine.recalc();
    for (const s of tx.sheets.keys()) SheetLayout.of(s).invalidate();
    const layout = tx.rowInfos.size > 0 || tx.colInfos.size > 0 || tx.sheets.size > 0;
    if (layout) for (const s of [...tx.rowInfos.keys(), ...tx.colInfos.keys()]) SheetLayout.of(s).invalidate();
    this.modified = true;
    const sheet = tx.sheets.size === 1 ? [...tx.sheets.keys()][0] : undefined;
    this.emit('change', { sheet, structure: tx.structural, layout, tx, ranges });
    this.emit('history', { canUndo: this.history.canUndo(), canRedo: this.history.canRedo() });
    if (tx.workbook) this.emit('sheets', { action: tx.label });
    void replay;
  }

  undo(): void {
    if (this.mode === 'viewOnly') return;
    const tx = this.history.popUndo();
    if (!tx) return;
    this.apply(tx, 'before');
  }

  redo(): void {
    if (this.mode === 'viewOnly') return;
    const tx = this.history.popRedo();
    if (!tx) return;
    this.apply(tx, 'after');
  }

  private apply(tx: Transaction, which: 'before' | 'after'): void {
    if (tx.workbook) restoreWorkbook(this.wb, tx.workbook[which]!);
    for (const [sheet, snap] of tx.sheets) restoreSheet(sheet, snap[which]!);
    for (const m of tx.cells.values()) for (const ch of m.values()) ch.sheet.setCell(ch.r, ch.c, ch[which]);
    for (const m of tx.rowInfos.values()) for (const ch of m.values()) ch.sheet.setRowInfo(ch.index, ch[which] ? { ...ch[which]! } : undefined);
    for (const m of tx.colInfos.values()) for (const ch of m.values()) ch.sheet.setColInfo(ch.index, ch[which] ? { ...ch[which]! } : undefined);
    const sel = which === 'before' ? tx.selectionBefore : tx.selectionAfter;
    if (sel) {
      this.wb.activeSheetIndex = Math.min(sel.sheetIndex, this.wb.sheets.length - 1);
      const sheet = this.sheet;
      if (sel.selection) {
        const s = sel.selection;
        this.selections.set(sheet, { active: { ...s.active }, anchor: { ...s.active }, ranges: s.ranges.map((r) => ({ ...r })) });
        sheet.selection = { active: { ...s.active }, ranges: s.ranges.map((r) => ({ ...r })) };
      }
    }
    this.afterChange(tx, true);
    this.emit('activeSheet', { index: this.activeIndex });
    this.emitSelection();
    this.emit('comments', {});
  }

  /** Sets a cell inside a transaction. */
  private put(tx: Transaction, sheet: Sheet, r: number, c: number, cell: Cell | undefined): void {
    tx.touch(sheet, r, c);
    sheet.setCell(r, c, cell);
  }

  // ---------------------------------------------------------------------------------------------
  // Cell input

  /** Applies text typed into a cell (formula, number, date, percent, text...). */
  setInput(r: number, c: number, text: string, sheet: Sheet = this.sheet): void {
    this.transact('Typing', (tx) => {
      const m = sheet.mergeAt(r, c);
      if (m) {
        r = m.r0;
        c = m.c0;
      }
      this.breakLegacyArray(tx, sheet, r, c);
      const existing = sheet.getCell(r, c);
      const styleId = existing?.s ?? sheet.styleAt(r, c);
      const style = this.wb.styles.get(styleId);
      const parsed = parseInput(text, { textFormat: style.numFmt === '@', date1904: this.wb.date1904 });
      let s: number | undefined = styleId || undefined;
      const cell: Cell = { v: parsed.value };
      if (parsed.formula !== undefined) {
        cell.f = normalizeFormulaInput(parsed.formula);
        cell.dyn = true;
        cell.v = null;
      }
      if (parsed.numFmt && (style.numFmt === 'General' || formatInfo(style.numFmt)?.type !== formatInfo(parsed.numFmt)?.type)) {
        s = this.wb.styles.patch(styleId, { numFmt: parsed.numFmt }) || undefined;
      }
      if (parsed.quotePrefix && !style.quotePrefix) s = this.wb.styles.patch(s, { quotePrefix: true }) || undefined;
      else if (!parsed.quotePrefix && style.quotePrefix && parsed.formula === undefined) s = this.wb.styles.patch(s, { quotePrefix: null }) || undefined;
      if (s) cell.s = s;
      // Typing a web address creates a link, like Excel's AutoCorrect.
      if (typeof cell.v === 'string' && /^(https?:\/\/|www\.)\S+$/i.test(cell.v) && !sheet.hyperlinkAt(r, c)) {
        this.addLinkInTx(tx, sheet, makeRange(r, c), cell.v.startsWith('www.') ? `http://${cell.v}` : cell.v);
        cell.s = this.linkStyle(cell.s);
      }
      this.put(tx, sheet, r, c, cell.v === null && cell.f === undefined && !cell.s ? undefined : cell);
    });
  }

  /** A cell inside a legacy (Ctrl+Shift+Enter) array formula cannot be edited alone. */
  private breakLegacyArray(tx: Transaction, sheet: Sheet, r: number, c: number): void {
    const cell = sheet.getCell(r, c);
    const anchorKey = cell?.spill;
    let ar = r;
    let ac = c;
    if (anchorKey !== undefined) {
      ar = Math.floor(anchorKey / MAX_COLS);
      ac = anchorKey % MAX_COLS;
    }
    const anchor = sheet.getCell(ar, ac);
    if (!anchor?.arr || anchor.dyn || (anchorKey === undefined && !(ar === r && ac === c))) return;
    const range = parseRange(anchor.arr);
    if (!range || (range.r0 === range.r1 && range.c0 === range.c1)) return;
    for (let rr = range.r0; rr <= range.r1; rr++) {
      for (let cc = range.c0; cc <= range.c1; cc++) {
        const x = sheet.getCell(rr, cc);
        if (!x) continue;
        const { spill: _s, arr: _a, f: _f, ...rest } = x;
        void _s;
        void _a;
        void _f;
        this.put(tx, sheet, rr, cc, rest);
      }
    }
  }

  /** Writes values (or formulas starting with '=') starting at a cell. */
  setValues(r0: number, c0: number, values: (Scalar | string)[][], sheet: Sheet = this.sheet): void {
    this.transact('Set values', (tx) => {
      values.forEach((row, i) =>
        row.forEach((v, j) => {
          const r = r0 + i;
          const c = c0 + j;
          const existing = sheet.getCell(r, c);
          if (typeof v === 'string' && v.startsWith('=') && v.length > 1) {
            this.put(tx, sheet, r, c, { s: existing?.s, v: null, f: normalizeFormulaInput(v.slice(1)), dyn: true });
          } else this.put(tx, sheet, r, c, { s: existing?.s, v: v as Scalar });
        }),
      );
    });
  }

  /** Effective ranges for iteration: whole rows/columns are limited to the used area. */
  private bounded(range: Range, sheet: Sheet): Range | null {
    if (!isFullColumns(range) && !isFullRows(range)) return range;
    const used = sheet.usedRange();
    if (!used) return null;
    const r = { r0: range.r0, c0: range.c0, r1: Math.min(range.r1, used.r1), c1: Math.min(range.c1, used.c1) };
    return r.r0 <= r.r1 && r.c0 <= r.c1 ? r : null;
  }

  clear(what: 'contents' | 'formats' | 'all' | 'comments' | 'links' = 'contents', ranges = this.selection.ranges, sheet = this.sheet): void {
    this.transact(what === 'contents' ? 'Clear contents' : `Clear ${what}`, (tx) => {
      for (const range of ranges) {
        const b = this.bounded(range, sheet);
        if (what === 'comments' || what === 'all') this.removeComments(tx, sheet, range);
        if (what === 'links' || what === 'all') this.removeLinksInTx(tx, sheet, range);
        if (!b) continue;
        const cells: [number, number, Cell][] = [];
        sheet.forEachCellIn(b, (r, c, cell) => cells.push([r, c, cell]));
        for (const [r, c, cell] of cells) {
          if (cell.spill !== undefined && cell.f === undefined && what === 'contents') continue;
          if (what === 'contents') this.put(tx, sheet, r, c, cell.s ? { v: null, s: cell.s } : undefined);
          else if (what === 'formats') this.put(tx, sheet, r, c, { ...cell, s: undefined, rt: undefined });
          else if (what === 'all') this.put(tx, sheet, r, c, undefined);
          else if (what === 'links' && cell.s) this.put(tx, sheet, r, c, { ...cell, s: this.wb.styles.patch(cell.s, { font: { color: null, underline: null } }) || undefined });
        }
        if (what === 'formats' || what === 'all') {
          if (isFullRows(range)) for (let r = range.r0; r <= range.r1; r++) this.setRowStyle(tx, sheet, r, undefined);
          if (isFullColumns(range)) for (let c = range.c0; c <= range.c1; c++) this.setColStyle(tx, sheet, c, undefined);
        }
      }
    });
  }

  // ---------------------------------------------------------------------------------------------
  // Formatting

  private setRowStyle(tx: Transaction, sheet: Sheet, r: number, s: number | undefined): void {
    tx.touchRow(sheet, r);
    const info = { ...(sheet.rows.get(r) ?? {}) };
    if (s) info.s = s;
    else delete info.s;
    sheet.setRowInfo(r, info);
  }

  private setColStyle(tx: Transaction, sheet: Sheet, c: number, s: number | undefined): void {
    tx.touchCol(sheet, c);
    const info = { ...(sheet.cols.get(c) ?? {}) };
    if (s) info.s = s;
    else delete info.s;
    sheet.setColInfo(c, info);
  }

  /** Applies a style change to every cell of the ranges (whole rows/columns use row/column styles). */
  private patchRanges(tx: Transaction, sheet: Sheet, ranges: Range[], patchFor: (r: number, c: number, current: CellStyle) => StylePatch | null): void {
    const styles = this.wb.styles;
    const cache = new Map<string, number>();
    const patched = (id: number, p: StylePatch) => {
      const key = `${id}|${JSON.stringify(p)}`;
      let v = cache.get(key);
      if (v === undefined) {
        v = styles.patch(id, p);
        cache.set(key, v);
      }
      return v;
    };
    for (const range of ranges) {
      const fullRows = isFullRows(range);
      const fullCols = isFullColumns(range);
      if (fullRows || fullCols) {
        if (fullCols && !fullRows) {
          for (let c = range.c0; c <= range.c1; c++) {
            const cur = sheet.cols.get(c)?.s ?? 0;
            const p = patchFor(-1, c, styles.get(cur));
            if (p) this.setColStyle(tx, sheet, c, patched(cur, p) || undefined);
          }
        } else {
          const rLimit = fullCols ? Math.max(sheet.dataExtent().rows, sheet.rows.size ? Math.max(...sheet.rows.keys()) : -1) : range.r1;
          for (let r = range.r0; r <= Math.min(range.r1, rLimit < 0 ? range.r0 + 999 : rLimit); r++) {
            const cur = sheet.rows.get(r)?.s ?? 0;
            const p = patchFor(r, -1, styles.get(cur));
            if (p) this.setRowStyle(tx, sheet, r, patched(cur, p) || undefined);
          }
        }
        // Existing cells in the area get the change too.
        const cells: [number, number, Cell][] = [];
        sheet.forEachCellIn(range, (r, c, cell) => cells.push([r, c, cell]));
        for (const [r, c, cell] of cells) {
          const cur = cell.s ?? 0;
          const p = patchFor(r, c, styles.get(cur));
          if (p) this.put(tx, sheet, r, c, { ...cell, s: patched(cur, p) || undefined });
        }
        continue;
      }
      for (let r = range.r0; r <= range.r1; r++) {
        for (let c = range.c0; c <= range.c1; c++) {
          const cell = sheet.getCell(r, c);
          const cur = cell?.s ?? sheet.styleAt(r, c);
          const p = patchFor(r, c, styles.get(cur));
          if (!p) continue;
          const s = patched(cur, p);
          if (cell?.s === s || (!cell && !s)) continue;
          this.put(tx, sheet, r, c, { ...(cell ?? { v: null }), s: s || undefined });
        }
      }
    }
  }

  format(patch: StylePatch, ranges = this.selection.ranges, sheet = this.sheet, label = 'Format'): void {
    this.transact(label, (tx) => {
      this.patchRanges(tx, sheet, ranges, () => patch);
      // Rich text keeps per-run fonts; a font change applied to the whole cell resets the runs.
      if (patch.font) {
        for (const range of ranges) {
          const b = this.bounded(range, sheet);
          if (!b) continue;
          const cells: [number, number, Cell][] = [];
          sheet.forEachCellIn(b, (r, c, cell) => {
            if (cell.rt) cells.push([r, c, cell]);
          });
          for (const [r, c, cell] of cells) {
            const keys = Object.keys(patch.font) as (keyof NonNullable<StylePatch['font']>)[];
            const runs = cell.rt!.map((run) => {
              if (!run.font) return run;
              const f = { ...run.font } as Record<string, unknown>;
              for (const k of keys) delete f[k];
              return { ...run, font: Object.keys(f).length ? f : undefined };
            });
            this.put(tx, sheet, r, c, { ...cell, rt: runs });
          }
        }
      }
    });
  }

  /** Toggles bold/italic/underline/strike based on the active cell (like a toolbar button). */
  toggleFont(prop: 'bold' | 'italic' | 'underline' | 'strike'): void {
    const sel = this.selection;
    const cur = this.styleAt(sel.active.r, sel.active.c).font;
    const on = prop === 'underline' ? !cur.underline : !cur[prop];
    this.format({ font: { [prop]: on ? (prop === 'underline' ? 'single' : true) : null } }, sel.ranges, this.sheet, `Toggle ${prop}`);
  }

  setNumberFormat(numFmt: string, ranges = this.selection.ranges): void {
    this.format({ numFmt }, ranges, this.sheet, 'Number format');
  }

  setBorder(kind: BorderKind, line: BorderLine | null = { style: 'thin' }, ranges = this.selection.ranges, sheet = this.sheet): void {
    this.transact('Borders', (tx) => {
      for (const range0 of ranges) {
        const range = this.bounded(range0, sheet) ?? range0;
        this.patchRanges(tx, sheet, [range], (r, c) => {
          if (r < 0 || c < 0) return null;
          const b: Partial<Record<'top' | 'bottom' | 'left' | 'right', BorderLine | null>> = {};
          const top = r === range.r0;
          const bottom = r === range.r1;
          const left = c === range.c0;
          const right = c === range.c1;
          const ln = line ?? null;
          switch (kind) {
            case 'none':
              return { border: null };
            case 'all':
              b.top = b.bottom = b.left = b.right = ln;
              break;
            case 'outside':
              if (top) b.top = ln;
              if (bottom) b.bottom = ln;
              if (left) b.left = ln;
              if (right) b.right = ln;
              break;
            case 'inside':
              if (!top) b.top = ln;
              if (!bottom) b.bottom = ln;
              if (!left) b.left = ln;
              if (!right) b.right = ln;
              break;
            case 'horizontal':
              if (!top) b.top = ln;
              if (!bottom) b.bottom = ln;
              break;
            case 'vertical':
              if (!left) b.left = ln;
              if (!right) b.right = ln;
              break;
            case 'top':
              if (top) b.top = ln;
              break;
            case 'bottom':
              if (bottom) b.bottom = ln;
              break;
            case 'left':
              if (left) b.left = ln;
              break;
            case 'right':
              if (right) b.right = ln;
              break;
          }
          return Object.keys(b).length ? { border: b } : null;
        });
      }
    });
  }

  /** Adds or removes one decimal place in the number format of the selection. */
  adjustDecimals(delta: 1 | -1): void {
    const { r, c } = this.selection.active;
    const cell = this.cell(r, c);
    let fmt = this.styleAt(r, c).numFmt;
    if (fmt === 'General') {
      if (typeof cell?.v !== 'number') fmt = '0';
      else {
        const s = formatValue(cell.v, 'General').text;
        const decimals = s.includes('.') ? s.split('.')[1].replace(/E.*/, '').length : 0;
        fmt = decimals ? `0.${'0'.repeat(decimals)}` : '0';
      }
    }
    const next = adjustDecimalFormat(fmt, delta);
    if (next !== null) this.setNumberFormat(next);
  }

  // ---------------------------------------------------------------------------------------------
  // Merges

  isMerged(range = this.range, sheet = this.sheet): boolean {
    return sheet.merges.some((m) => rangeEquals(m, range));
  }

  merge(ranges = this.selection.ranges, sheet = this.sheet, mode: 'all' | 'across' = 'all'): void {
    this.transact('Merge cells', (tx) => {
      tx.snapshotSheet(sheet);
      let merges = [...sheet.merges];
      for (const range0 of ranges) {
        let range = sheet.expandToMerges(range0);
        const pieces = mode === 'across' ? Array.from({ length: range.r1 - range.r0 + 1 }, (_, i) => makeRange(range.r0 + i, range.c0, range.r0 + i, range.c1)) : [range];
        for (const piece of pieces) {
          range = piece;
          if (range.r0 === range.r1 && range.c0 === range.c1) continue;
          merges = merges.filter((m) => !rangesIntersect(m, range));
          // Keep the top-left value only (like Excel).
          let keep: Cell | undefined;
          for (let r = range.r0; r <= range.r1 && !keep; r++) {
            for (let c = range.c0; c <= range.c1 && !keep; c++) {
              const cell = sheet.getCell(r, c);
              if (cell && (cell.v !== null || cell.f !== undefined)) keep = cell;
            }
          }
          const anchorStyle = sheet.getCell(range.r0, range.c0)?.s;
          for (let r = range.r0; r <= range.r1; r++) {
            for (let c = range.c0; c <= range.c1; c++) {
              if (r === range.r0 && c === range.c0) continue;
              const cell = sheet.getCell(r, c);
              if (cell && (cell.v !== null || cell.f !== undefined)) sheet.setCell(r, c, cell.s ? { v: null, s: cell.s } : undefined);
            }
          }
          if (keep) sheet.setCell(range.r0, range.c0, { ...keep, s: anchorStyle ?? keep.s });
          merges.push(range);
        }
      }
      sheet.setMerges(merges);
    }, { kind: 'MergeCells' });
    this.select(this.range);
  }

  unmerge(ranges = this.selection.ranges, sheet = this.sheet): void {
    this.transact('Unmerge cells', (tx) => {
      tx.snapshotSheet(sheet);
      sheet.setMerges(sheet.merges.filter((m) => !ranges.some((r) => rangesIntersect(m, r))));
    }, { kind: 'UnmergeCells' });
  }

  toggleMerge(): void {
    const sheet = this.sheet;
    const ranges = this.selection.ranges;
    const hasMerge = ranges.some((r) => sheet.merges.some((m) => rangesIntersect(m, r)));
    if (hasMerge && ranges.every((r) => sheet.merges.some((m) => rangeEquals(m, r)))) this.unmerge();
    else this.merge();
  }

  // ---------------------------------------------------------------------------------------------
  // Rows and columns

  private adjustAllFormulas(tx: Transaction, rewrite: (formula: string, sheet: Sheet) => string, skip?: Sheet): void {
    for (const s of this.wb.sheets) {
      if (s === skip) continue;
      const changes: [number, number, Cell][] = [];
      s.forEachCell((r, c, cell) => {
        if (cell.f === undefined) return;
        const f = rewrite(cell.f, s);
        if (f !== cell.f) changes.push([r, c, { ...cell, f }]);
      });
      for (const [r, c, cell] of changes) this.put(tx, s, r, c, cell);
    }
    // Charts made or edited here: their references follow too.
    for (const s of this.wb.sheets) {
      for (const chart of s.charts) {
        const next = rewriteChartRefs(chart, (f) => rewrite(f, CHART_HOME));
        if (next) {
          tx.snapshotSheet(s);
          const target = s.charts.find((x) => x.id === chart.id)!;
          Object.assign(target, next);
        }
      }
    }
    // Chart series references are always sheet-qualified, so no home sheet applies to them.
    if (this.wb.chartRefs.size) {
      let changed = false;
      const next = new Map<string, string[]>();
      for (const [path, refs] of this.wb.chartRefs) {
        const updated = refs.map((f) => rewrite(f, CHART_HOME));
        if (updated.some((f, i) => f !== refs[i])) changed = true;
        next.set(path, updated);
      }
      if (changed) {
        tx.snapshotWorkbook(this.wb);
        this.wb.chartRefs = next;
      }
    }
  }

  private shiftStructure(tx: Transaction, sheet: Sheet, axis: 'row' | 'col', from: number, delta: number): void {
    tx.snapshotWorkbook(this.wb);
    tx.snapshotSheet(sheet);
    if (axis === 'row') {
      if (delta > 0) sheet.insertRows(from, delta);
      else sheet.deleteRows(from + delta, -delta);
    } else if (delta > 0) sheet.insertCols(from, delta);
    else sheet.deleteCols(from + delta, -delta);
    // References in every sheet move with the cells.
    const name = sheet.name;
    for (const s of this.wb.sheets) {
      if (s === sheet) {
        const changes: [number, number, Cell][] = [];
        s.forEachCell((r, c, cell) => {
          if (cell.f === undefined) return;
          const f = shiftFormula(cell.f, s.name, name, axis, delta > 0 ? from : from, delta);
          if (f !== cell.f) changes.push([r, c, { ...cell, f }]);
        });
        for (const [r, c, cell] of changes) s.setCell(r, c, cell);
      }
    }
    this.adjustAllFormulas(tx, (f, s) => shiftFormula(f, s.name, name, axis, from, delta), sheet);
    this.wb.names = this.wb.names.map((n) => ({ ...n, formula: shiftFormula(n.formula, n.localSheet !== undefined ? this.wb.sheets[n.localSheet]?.name ?? '' : '', name, axis, from, delta) }));
  }

  insertRows(at: number, count: number, sheet = this.sheet): void {
    this.transact('Insert rows', (tx) => {
      this.shiftStructure(tx, sheet, 'row', at, count);
      // New rows take the formatting of the row above.
      if (at > 0) {
        const above = sheet.cells.get(at - 1);
        const info = sheet.rows.get(at - 1);
        for (let r = at; r < at + count; r++) {
          if (info?.s !== undefined || (info?.h !== undefined && info.customHeight)) sheet.setRowInfo(r, { s: info.s, h: info.customHeight ? info.h : undefined, customHeight: info.customHeight });
          if (above) for (const [c, cell] of above) if (cell.s) sheet.setCell(r, c, { v: null, s: cell.s });
        }
      }
    }, { kind: 'AddRowColumn' });
  }

  deleteRows(at: number, count: number, sheet = this.sheet): void {
    this.transact('Delete rows', (tx) => this.shiftStructure(tx, sheet, 'row', at + count, -count), { kind: 'DeleteRowColumn' });
    this.selectCell(Math.min(at, MAX_ROWS - 1), this.selection.active.c);
  }

  insertCols(at: number, count: number, sheet = this.sheet): void {
    this.transact('Insert columns', (tx) => {
      this.shiftStructure(tx, sheet, 'col', at, count);
      if (at > 0) {
        const info = sheet.cols.get(at - 1);
        for (let c = at; c < at + count; c++) {
          if (info) sheet.setColInfo(c, { ...info });
          for (const row of sheet.cells.values()) {
            const left = row.get(at - 1);
            if (left?.s && !row.has(c)) row.set(c, { v: null, s: left.s });
          }
        }
      }
    }, { kind: 'AddRowColumn' });
  }

  deleteCols(at: number, count: number, sheet = this.sheet): void {
    this.transact('Delete columns', (tx) => this.shiftStructure(tx, sheet, 'col', at + count, -count), { kind: 'DeleteRowColumn' });
    this.selectCell(this.selection.active.r, Math.min(at, MAX_COLS - 1));
  }

  /** Inserts as many rows as the selection spans, above or below it. */
  insertRowsAtSelection(where: 'above' | 'below'): void {
    const range = this.range;
    const count = Math.min(range.r1 - range.r0 + 1, 1000);
    this.insertRows(where === 'above' ? range.r0 : range.r1 + 1, isFullColumns(range) ? 1 : count);
  }

  insertColsAtSelection(where: 'left' | 'right'): void {
    const range = this.range;
    const count = Math.min(range.c1 - range.c0 + 1, 100);
    this.insertCols(where === 'left' ? range.c0 : range.c1 + 1, isFullRows(range) ? 1 : count);
  }

  deleteSelectedRows(): void {
    const range = this.range;
    if (isFullColumns(range)) return;
    this.deleteRows(range.r0, range.r1 - range.r0 + 1);
  }

  deleteSelectedCols(): void {
    const range = this.range;
    if (isFullRows(range)) return;
    this.deleteCols(range.c0, range.c1 - range.c0 + 1);
  }

  /** Row heights in points (null restores the default/automatic height). */
  setRowHeight(rows: number[], pt: number | null, sheet = this.sheet): void {
    this.transact('Row height', (tx) => {
      for (const r of rows) {
        tx.touchRow(sheet, r);
        const info = { ...(sheet.rows.get(r) ?? {}) };
        if (pt === null) {
          delete info.h;
          delete info.customHeight;
        } else {
          info.h = Math.max(0, Math.min(409, pt));
          info.customHeight = true;
          if (pt === 0) info.hidden = true;
        }
        sheet.setRowInfo(r, info);
      }
    });
  }

  /** Column widths in pixels (null restores the default width). */
  setColWidth(cols: number[], px: number | null, sheet = this.sheet): void {
    this.transact('Column width', (tx) => {
      for (const c of cols) {
        tx.touchCol(sheet, c);
        const info = { ...(sheet.cols.get(c) ?? {}) };
        if (px === null) {
          delete info.w;
          delete info.customWidth;
        } else {
          info.w = widthPxToChars(Math.max(0, Math.min(2000, px)));
          info.customWidth = true;
        }
        sheet.setColInfo(c, info);
      }
    });
  }

  setHidden(axis: 'row' | 'col', indexes: number[], hidden: boolean, sheet = this.sheet): void {
    this.transact(hidden ? 'Hide' : 'Unhide', (tx) => {
      for (const i of indexes) {
        if (axis === 'row') {
          tx.touchRow(sheet, i);
          sheet.setRowInfo(i, { ...(sheet.rows.get(i) ?? {}), hidden: hidden || undefined, h: !hidden && sheet.rows.get(i)?.h === 0 ? undefined : sheet.rows.get(i)?.h });
        } else {
          tx.touchCol(sheet, i);
          sheet.setColInfo(i, { ...(sheet.cols.get(i) ?? {}), hidden: hidden || undefined, w: !hidden && sheet.cols.get(i)?.w === 0 ? undefined : sheet.cols.get(i)?.w });
        }
      }
    });
  }

  /** Fits column widths to their content (double-click on the header border). */
  autoFitColumns(cols: number[], sheet = this.sheet): void {
    if (!this.measure) return;
    const measure = this.measure;
    const widths = cols.map((c) => {
      let max = 0;
      const range = makeRange(0, c, MAX_ROWS - 1, c);
      sheet.forEachCellIn(range, (r, cc, cell) => {
        if (cell.v === null || sheet.mergeAt(r, cc)) return;
        const style = this.wb.styles.get(sheet.styleAt(r, cc));
        const text = formatValue(cell.v, style.numFmt, this.wb.date1904).text;
        const lines = style.align?.wrap ? [text] : text.split('\n');
        for (const line of lines) max = Math.max(max, measure(line, style).width);
      });
      return max ? Math.ceil(max + 10) : null;
    });
    this.transact('Auto-fit columns', (tx) => {
      cols.forEach((c, i) => {
        const w = widths[i];
        if (w === null) return;
        tx.touchCol(sheet, c);
        sheet.setColInfo(c, { ...(sheet.cols.get(c) ?? {}), w: widthPxToChars(w), customWidth: true, bestFit: true });
      });
    });
  }

  /** Fits row heights to their content. */
  autoFitRows(rows: number[], sheet = this.sheet): void {
    this.transact('Auto-fit rows', (tx) => {
      for (const r of rows) this.fitRow(tx, sheet, r, true);
    });
  }

  private fitRow(tx: Transaction, sheet: Sheet, r: number, force: boolean): void {
    if (!this.measure) return;
    const info = sheet.rows.get(r);
    if (info?.customHeight && !force) return;
    let maxPx = 0;
    const row = sheet.cells.get(r);
    for (const [c, cell] of row ?? []) {
      const style = this.wb.styles.get(cell.s ?? sheet.styleAt(r, c));
      if (cell.v === null) {
        maxPx = Math.max(maxPx, this.measure('Xg', style).height);
        continue;
      }
      if (sheet.mergeAt(r, c)) continue;
      const text = formatValue(cell.v, style.numFmt, this.wb.date1904).text;
      const m = this.measure(text, style.align?.wrap ? { ...style, align: { ...style.align, wrap: true } } : style);
      if (style.align?.wrap) {
        const width = Math.max(1, sheet.colWidthPx(c) - 6);
        const lines = Math.max(1, Math.ceil(m.width / width) + (text.split('\n').length - 1));
        maxPx = Math.max(maxPx, lines * this.measure('Xg', style).height);
      } else maxPx = Math.max(maxPx, m.height * Math.max(1, text.split('\n').length));
    }
    const defaultPx = ptToPx(sheet.defaultRowHeightPt);
    const target = maxPx > defaultPx - 2 ? pxToPt(Math.ceil(maxPx + 4)) : undefined;
    const current = info?.h;
    if ((target ?? sheet.defaultRowHeightPt) === (current ?? sheet.defaultRowHeightPt) && !force) return;
    tx.touchRow(sheet, r);
    const next = { ...(info ?? {}) };
    if (target && target > sheet.defaultRowHeightPt) next.h = target;
    else delete next.h;
    if (force) delete next.customHeight;
    sheet.setRowInfo(r, next);
  }

  private autoFitTouchedRows(tx: Transaction): void {
    if (!this.measure) return;
    for (const [sheet, m] of tx.cells) {
      const rows = new Set<number>();
      for (const ch of m.values()) rows.add(ch.r);
      if (rows.size > 500) continue;
      for (const r of rows) this.fitRow(tx, sheet, r, false);
    }
  }

  setFreeze(rows: number, cols: number, sheet = this.sheet): void {
    this.transact('Freeze panes', (tx) => {
      tx.snapshotSheet(sheet);
      sheet.freeze = { rows: Math.max(0, rows), cols: Math.max(0, cols) };
    });
  }

  setZoom(zoom: number, sheet = this.sheet): void {
    sheet.zoom = Math.max(0.1, Math.min(4, zoom));
    this.emit('change', { sheet, structure: false, layout: true, ranges: [] });
  }

  setGridLines(show: boolean, sheet = this.sheet): void {
    this.transact('Gridlines', (tx) => {
      tx.snapshotSheet(sheet);
      sheet.showGridLines = show;
    });
  }

  // ---------------------------------------------------------------------------------------------
  // Sheets

  addSheet(name?: string, index = this.wb.sheets.length, baseName = 'Sheet'): Sheet | null {
    let sheet: Sheet | null = null;
    this.transact('Add sheet', (tx) => {
      tx.snapshotWorkbook(this.wb);
      const n = name && !validateSheetName(name, this.wb) ? name : this.wb.uniqueSheetName(baseName);
      sheet = this.wb.addSheet(n, index);
      sheet.defaultRowHeightPt = this.wb.sheets[0]?.defaultRowHeightPt ?? sheet.defaultRowHeightPt;
      this.remapLocalNames(index, 1);
      this.wb.activeSheetIndex = this.wb.sheets.indexOf(sheet);
    }, { kind: 'createSheet' });
    if (sheet) {
      this.emit('activeSheet', { index: this.activeIndex });
      this.emitSelection();
    }
    return sheet;
  }

  /** Shifts localSheetId of defined names after inserting (delta 1) or removing (delta -1) a sheet. */
  private remapLocalNames(index: number, delta: number): void {
    this.wb.names = this.wb.names
      .filter((n) => !(delta < 0 && n.localSheet === index))
      .map((n) => (n.localSheet !== undefined && n.localSheet >= index + (delta < 0 ? 1 : 0) ? { ...n, localSheet: n.localSheet + delta } : n));
  }

  removeSheet(index: number): boolean {
    const visible = this.wb.visibleSheets();
    const sheet = this.wb.sheets[index];
    if (!sheet || (sheet.state === 'visible' && visible.length <= 1)) return false;
    this.transact('Delete sheet', (tx) => {
      tx.snapshotWorkbook(this.wb);
      const name = sheet.name;
      this.wb.sheets.splice(index, 1);
      this.remapLocalNames(index, -1);
      this.adjustAllFormulas(tx, (f) => deleteSheetInFormula(f, name));
      this.wb.names = this.wb.names.map((n) => ({ ...n, formula: deleteSheetInFormula(n.formula, name) }));
      for (const s of this.wb.sheets) {
        if (s.hyperlinks.some((h) => h.location && splitSheetRef(h.location).sheet === name)) {
          tx.snapshotSheet(s);
          s.setHyperlinks(s.hyperlinks.map((h) => (h.location && splitSheetRef(h.location).sheet === name ? { ...h, location: '#REF!' } : h)));
        }
      }
      let next = Math.min(index, this.wb.sheets.length - 1);
      while (next > 0 && this.wb.sheets[next].state !== 'visible') next--;
      if (this.wb.sheets[next].state !== 'visible') next = this.wb.sheets.findIndex((s) => s.state === 'visible');
      this.wb.activeSheetIndex = next;
    }, { kind: 'removeSheet' });
    this.emit('activeSheet', { index: this.activeIndex });
    this.emitSelection();
    return true;
  }

  renameSheet(index: number, newName: string): string | null {
    const sheet = this.wb.sheets[index];
    if (!sheet) return 'Unknown sheet';
    const name = newName.trim();
    const error = validateSheetName(name, this.wb, sheet);
    if (error) return error;
    if (name === sheet.name) return null;
    this.transact('Rename sheet', (tx) => {
      tx.snapshotWorkbook(this.wb);
      const old = sheet.name;
      tx.snapshotSheet(sheet);
      sheet.name = name;
      this.adjustAllFormulas(tx, (f) => renameSheetInFormula(f, old, name));
      this.wb.names = this.wb.names.map((n) => ({ ...n, formula: renameSheetInFormula(n.formula, old, name) }));
      for (const s of this.wb.sheets) {
        if (s.hyperlinks.some((h) => h.location && splitSheetRef(h.location).sheet?.toLowerCase() === old.toLowerCase())) {
          tx.snapshotSheet(s);
          s.setHyperlinks(s.hyperlinks.map((h) => (h.location ? { ...h, location: renameSheetInFormula(h.location, old, name) } : h)));
        }
      }
    }, { kind: 'renameSheet' });
    return null;
  }

  moveSheet(from: number, to: number): void {
    if (from === to || from < 0 || to < 0 || from >= this.wb.sheets.length || to >= this.wb.sheets.length) return;
    this.transact('Move sheet', (tx) => {
      tx.snapshotWorkbook(this.wb);
      const active = this.sheet;
      const [s] = this.wb.sheets.splice(from, 1);
      this.wb.sheets.splice(to, 0, s);
      const map = (i: number) => {
        if (i === from) return to;
        if (from < to && i > from && i <= to) return i - 1;
        if (from > to && i >= to && i < from) return i + 1;
        return i;
      };
      this.wb.names = this.wb.names.map((n) => (n.localSheet !== undefined ? { ...n, localSheet: map(n.localSheet) } : n));
      this.wb.activeSheetIndex = this.wb.sheets.indexOf(active);
    });
    this.emit('activeSheet', { index: this.activeIndex });
  }

  setSheetHidden(index: number, hidden: boolean): boolean {
    const sheet = this.wb.sheets[index];
    if (!sheet) return false;
    if (hidden && this.wb.visibleSheets().length <= 1) return false;
    this.transact(hidden ? 'Hide sheet' : 'Unhide sheet', (tx) => {
      tx.snapshotWorkbook(this.wb);
      tx.snapshotSheet(sheet);
      sheet.state = hidden ? 'hidden' : 'visible';
      if (hidden && this.wb.activeSheetIndex === index) {
        const next = this.wb.sheets.findIndex((s) => s.state === 'visible');
        this.wb.activeSheetIndex = next;
      }
      if (!hidden) this.wb.activeSheetIndex = index;
    }, { kind: hidden ? 'hideSheet' : 'showSheet' });
    this.emit('activeSheet', { index: this.activeIndex });
    return true;
  }

  duplicateSheet(index: number): Sheet | null {
    const src = this.wb.sheets[index];
    if (!src) return null;
    let copy: Sheet | null = null;
    this.transact('Duplicate sheet', (tx) => {
      tx.snapshotWorkbook(this.wb);
      let n = 2;
      const base = src.name.replace(/ \(\d+\)$/, '').slice(0, 25);
      while (this.wb.sheetByName(`${base} (${n})`)) n++;
      copy = new Sheet(`${base} (${n})`, this.wb.nextSheetId());
      restoreSheet(copy, { ...snapshotOf(src), name: copy.name, state: 'visible' });
      copy.selection = src.selection;
      // Copied charts show the copied data; objects get their own identities.
      const copyName = copy.name;
      copy.charts = copy.charts.map((ch) => ({ ...ch, ...(rewriteChartRefs(ch, (f) => renameSheetInFormula(f, src.name, copyName)) ?? {}), id: newDrawingId('chart'), source: undefined, dirty: true }));
      copy.shapes = copy.shapes.map((sh) => ({ ...sh, id: newDrawingId('shape'), source: undefined, dirty: true }));
      copy.images = copy.images.map((img) => ({ ...img, id: newDrawingId('img'), raw: undefined }));
      this.wb.sheets.splice(index + 1, 0, copy);
      this.remapLocalNames(index + 1, 1);
      this.wb.activeSheetIndex = index + 1;
    }, { kind: 'createSheet' });
    this.emit('activeSheet', { index: this.activeIndex });
    return copy;
  }

  setTabColor(index: number, color: string | null): void {
    const sheet = this.wb.sheets[index];
    if (!sheet) return;
    this.transact('Tab color', (tx) => {
      tx.snapshotWorkbook(this.wb);
      tx.snapshotSheet(sheet);
      sheet.tabColor = color ? rgbSpec(color) : undefined;
    });
  }

  // ---------------------------------------------------------------------------------------------
  // Images

  /** Inserts an image at a pixel position (100% zoom) with a pixel size. */
  insertImage(data: Uint8Array, mime: string, rect: { x: number; y: number; width: number; height: number }, name?: string, sheet = this.sheet): SheetImage | null {
    let img: SheetImage | null = null;
    this.transact('Insert image', (tx) => {
      tx.snapshotSheet(sheet);
      const layout = SheetLayout.of(sheet);
      img = {
        id: `img-${Date.now().toString(36)}-${Math.floor(Math.random() * 1e6).toString(36)}`,
        name: name ?? `Picture ${sheet.images.length + 1}`,
        mime,
        data,
        anchorType: 'oneCell',
        from: layout.anchorAt(rect.x, rect.y),
        ext: { cx: pxToEmu(rect.width), cy: pxToEmu(rect.height) },
      };
      sheet.images.push(img);
    });
    return img;
  }

  /** Moves/resizes an image to a pixel rectangle (100% zoom). */
  setImageRect(id: string, rect: { x: number; y: number; width: number; height: number }, sheet = this.sheet): void {
    const img = sheet.images.find((i) => i.id === id);
    if (!img) return;
    this.transact('Move image', (tx) => {
      tx.snapshotSheet(sheet);
      const layout = SheetLayout.of(sheet);
      const target = sheet.images.find((i) => i.id === id)!;
      target.from = layout.anchorAt(rect.x, rect.y);
      if (target.anchorType === 'twoCell') target.to = layout.anchorAt(rect.x + rect.width, rect.y + rect.height);
      else target.ext = { cx: pxToEmu(rect.width), cy: pxToEmu(rect.height) };
      if (target.anchorType === 'absolute') {
        target.anchorType = 'oneCell';
        target.pos = undefined;
      }
      target.raw = undefined;
    });
  }

  deleteImage(id: string, sheet = this.sheet): void {
    this.transact('Delete image', (tx) => {
      tx.snapshotSheet(sheet);
      sheet.images = sheet.images.filter((i) => i.id !== id);
    });
  }

  // ---------------------------------------------------------------------------------------------
  // Comments

  allComments(): { sheet: Sheet; thread: CommentThread }[] {
    const out: { sheet: Sheet; thread: CommentThread }[] = [];
    for (const sheet of this.wb.sheets) for (const thread of sheet.comments) out.push({ sheet, thread });
    return out;
  }

  findThread(id: string): { sheet: Sheet; thread: CommentThread; index: number } | null {
    for (const sheet of this.wb.sheets) {
      for (const thread of sheet.comments) {
        const index = thread.comments.findIndex((c) => c.id === id);
        if (thread.id === id || index >= 0) return { sheet, thread, index: Math.max(0, index) };
      }
    }
    return null;
  }

  private commentTx(label: string, sheet: Sheet, fn: (sheet: Sheet) => void): void {
    this.transact(label, (tx) => {
      tx.snapshotSheet(sheet);
      fn(sheet);
    });
    this.emit('comments', { sheet });
  }

  /** Adds a comment to a cell; if the cell already has a thread the comment becomes a reply. */
  addComment(r: number, c: number, text: string, author = this.author, sheet = this.sheet): string {
    const existing = sheet.commentAt(r, c);
    if (existing) return this.reply(existing.id, text, author) ?? existing.id;
    const id = newGuid();
    this.commentTx('Add comment', sheet, (s) => {
      s.comments.push({ id, r, c, kind: 'threaded', resolved: false, comments: [{ id, author, text, created: new Date().toISOString() }] });
    });
    return id;
  }

  reply(threadOrCommentId: string, text: string, author = this.author): string | null {
    const hit = this.findThread(threadOrCommentId);
    if (!hit) return null;
    const id = newGuid();
    this.commentTx('Reply', hit.sheet, (s) => {
      const t = s.comments.find((x) => x.id === hit.thread.id)!;
      t.comments.push({ id, author, text, created: new Date().toISOString() });
      t.kind = 'threaded';
      t.vmlShape = undefined;
    });
    return id;
  }

  editComment(commentId: string, text: string): void {
    const hit = this.findThread(commentId);
    if (!hit) return;
    this.commentTx('Edit comment', hit.sheet, (s) => {
      const t = s.comments.find((x) => x.id === hit.thread.id)!;
      const cm = t.comments.find((x) => x.id === commentId) ?? t.comments[0];
      cm.text = text;
      // Editing a resolved thread reopens it.
      t.resolved = false;
    });
  }

  deleteComment(commentId: string): void {
    const hit = this.findThread(commentId);
    if (!hit) return;
    this.commentTx('Delete comment', hit.sheet, (s) => {
      const t = s.comments.find((x) => x.id === hit.thread.id)!;
      if (t.id === commentId || t.comments[0].id === commentId) s.comments = s.comments.filter((x) => x !== t);
      else t.comments = t.comments.filter((x) => x.id !== commentId);
    });
  }

  setResolved(threadOrCommentId: string, resolved: boolean): void {
    const hit = this.findThread(threadOrCommentId);
    if (!hit) return;
    this.commentTx(resolved ? 'Resolve' : 'Reopen', hit.sheet, (s) => {
      const t = s.comments.find((x) => x.id === hit.thread.id)!;
      t.resolved = resolved;
      t.kind = 'threaded';
    });
  }

  private removeComments(tx: Transaction, sheet: Sheet, range: Range): void {
    if (!sheet.comments.some((t) => rangeContains(range, t.r, t.c))) return;
    tx.snapshotSheet(sheet);
    sheet.comments = sheet.comments.filter((t) => !rangeContains(range, t.r, t.c));
  }

  // ---------------------------------------------------------------------------------------------
  // Hyperlinks

  private linkStyle(s: number | undefined): number | undefined {
    return this.wb.styles.patch(s, { font: { color: { theme: 10 }, underline: 'single' } }) || undefined;
  }

  private addLinkInTx(tx: Transaction, sheet: Sheet, range: Range, target: string): void {
    tx.snapshotSheet(sheet);
    const internal = /^#/.test(target) ? target.slice(1) : !/^[a-z][a-z0-9+.-]*:/i.test(target) && /!/.test(target) ? target : undefined;
    const links = sheet.hyperlinks.filter((h) => !rangesIntersect(h.ref, range));
    links.push(internal ? { ref: range, location: internal } : { ref: range, target });
    sheet.setHyperlinks(links);
  }

  private removeLinksInTx(tx: Transaction, sheet: Sheet, range: Range): void {
    if (!sheet.hyperlinks.some((h) => rangesIntersect(h.ref, range))) return;
    tx.snapshotSheet(sheet);
    sheet.setHyperlinks(sheet.hyperlinks.filter((h) => !rangesIntersect(h.ref, range)));
  }

  /** Adds a link to the active cell; `text` replaces the cell content when given. */
  setLink(target: string, text?: string, r = this.selection.active.r, c = this.selection.active.c, sheet = this.sheet): void {
    this.transact('Insert link', (tx) => {
      this.addLinkInTx(tx, sheet, makeRange(r, c), target);
      const cell = sheet.getCell(r, c);
      this.put(tx, sheet, r, c, { ...(cell ?? { v: null }), v: text ?? cell?.v ?? target, f: text !== undefined ? undefined : cell?.f, s: this.linkStyle(cell?.s ?? sheet.styleAt(r, c)) });
    });
  }

  removeLink(ranges = this.selection.ranges, sheet = this.sheet): void {
    this.clear('links', ranges, sheet);
  }

  // ---------------------------------------------------------------------------------------------
  // Clipboard

  /** Copies the primary selection; returns text and HTML for the system clipboard. */
  copy(cut = false): ClipPayload | null {
    const sheet = this.sheet;
    const range = this.bounded(this.range, sheet) ?? this.range;
    if ((range.r1 - range.r0 + 1) * (range.c1 - range.c0 + 1) > 2_000_000) return null;
    const cells = clone2d(range.r1 - range.r0 + 1, range.c1 - range.c0 + 1, (r, c) => sheet.getCell(range.r0 + r, range.c0 + c));
    const merges = sheet.merges.filter((m) => rangesIntersect(m, range));
    const textRows: string[][] = [];
    const htmlRows: HtmlOutCell[][] = [];
    for (let r = range.r0; r <= range.r1; r++) {
      const tr: string[] = [];
      const hr: HtmlOutCell[] = [];
      for (let c = range.c0; c <= range.c1; c++) {
        const text = this.displayText(r, c, sheet);
        tr.push(text);
        const m = sheet.mergeAt(r, c);
        const hidden = !!m && (m.r0 !== r || m.c0 !== c);
        hr.push({ text, style: this.htmlStyle(this.styleAt(r, c, sheet)), hidden, colspan: m && !hidden ? Math.min(m.c1, range.c1) - c + 1 : undefined, rowspan: m && !hidden ? Math.min(m.r1, range.r1) - r + 1 : undefined });
      }
      textRows.push(tr);
      htmlRows.push(hr);
    }
    const text = toTsv(textRows);
    this.clip = {
      sheet,
      range,
      cells,
      merges,
      text,
      cut,
      colWidths: Array.from({ length: range.c1 - range.c0 + 1 }, (_, i) => sheet.colWidthPx(range.c0 + i)),
    };
    return { text, html: toHtml(htmlRows) };
  }

  cut(): ClipPayload | null {
    if (this.isReadOnly()) return this.copy(false);
    return this.copy(true);
  }

  hasClip(): boolean {
    return !!this.clip;
  }

  clearClip(): void {
    this.clip = undefined;
  }

  private htmlStyle(s: CellStyle): HtmlCellStyle {
    const out: HtmlCellStyle = {};
    if (s.font.bold) out.bold = true;
    if (s.font.italic) out.italic = true;
    if (s.font.underline) out.underline = true;
    if (s.font.strike) out.strike = true;
    if (s.font.color) out.color = `#${resolveColor(s.font.color, this.wb.theme, '000000')}`;
    if (s.fill && 'pattern' in s.fill && s.fill.pattern === 'solid') out.background = `#${resolveColor(s.fill.fg, this.wb.theme, 'FFFFFF')}`;
    const h = s.align?.horizontal;
    if (h === 'left' || h === 'center' || h === 'right') out.align = h;
    out.fontSize = s.font.size;
    out.fontFamily = s.font.name;
    return out;
  }

  /**
   * Pastes at the selection. Internal data (copied in this editor) keeps formulas and formatting;
   * external data is parsed from HTML tables or tab separated text.
   */
  paste(data: { text?: string; html?: string }, mode: 'all' | 'values' | 'formats' | 'formulas' = 'all'): void {
    const internal = this.clip && (data.text === undefined || normalizeNewlines(data.text) === normalizeNewlines(this.clip.text)) ? this.clip : undefined;
    if (internal) {
      this.pasteInternal(internal, mode);
      return;
    }
    let rows: { text: string; style?: HtmlCellStyle }[][] | null = null;
    if (data.html) {
      const table = parseHtmlTable(data.html);
      if (table) rows = table.rows.map((row) => row.map((c) => ({ text: c.text, style: c.style })));
    }
    if (!rows && data.text !== undefined) {
      const text = data.text.replace(/\n$/, '');
      rows = parseDelimited(text, '\t').map((row) => row.map((t) => ({ text: t })));
    }
    if (!rows || !rows.length) return;
    const sheet = this.sheet;
    const { r: r0, c: c0 } = { r: this.range.r0, c: this.range.c0 };
    this.transact('Paste', (tx) => {
      rows!.forEach((row, i) => {
        row.forEach((cellData, j) => {
          const r = r0 + i;
          const c = c0 + j;
          if (r >= MAX_ROWS || c >= MAX_COLS) return;
          const existing = sheet.getCell(r, c);
          const styleId = existing?.s ?? sheet.styleAt(r, c);
          const parsed = parseInput(cellData.text, { date1904: this.wb.date1904 });
          let s = styleId;
          if (parsed.numFmt && this.wb.styles.get(s).numFmt === 'General') s = this.wb.styles.patch(s, { numFmt: parsed.numFmt });
          if (cellData.style && mode !== 'values') s = this.wb.styles.patch(s, htmlToPatch(cellData.style));
          const cell: Cell = parsed.formula !== undefined ? { v: null, f: normalizeFormulaInput(parsed.formula), dyn: true } : { v: parsed.value };
          if (s) cell.s = s;
          this.put(tx, sheet, r, c, cell.v === null && cell.f === undefined && !cell.s ? undefined : cell);
        });
      });
    });
    this.select(makeRange(r0, c0, r0 + rows.length - 1, c0 + Math.max(...rows.map((r) => r.length)) - 1), { r: r0, c: c0 });
  }

  private pasteInternal(clip: InternalClip, mode: 'all' | 'values' | 'formats' | 'formulas'): void {
    const sheet = this.sheet;
    const target = this.range;
    const h = clip.range.r1 - clip.range.r0 + 1;
    const w = clip.range.c1 - clip.range.c0 + 1;
    // A target that is a multiple of the source size is filled by tiling (like Excel).
    const th = target.r1 - target.r0 + 1;
    const tw = target.c1 - target.c0 + 1;
    const tilesR = !clip.cut && th % h === 0 && th > h ? th / h : 1;
    const tilesC = !clip.cut && tw % w === 0 && tw > w ? tw / w : 1;
    const r0 = target.r0;
    const c0 = target.c0;
    const dr = r0 - clip.range.r0;
    const dc = c0 - clip.range.c0;
    this.transact(clip.cut ? 'Move' : 'Paste', (tx) => {
      const merges = clip.merges.length || sheet.merges.some((m) => rangesIntersect(m, makeRange(r0, c0, r0 + h * tilesR - 1, c0 + w * tilesC - 1)));
      if (merges || clip.cut) tx.snapshotSheet(sheet);
      if (clip.cut) {
        // Clear the source first (the target may overlap it).
        for (let r = clip.range.r0; r <= clip.range.r1; r++) {
          for (let c = clip.range.c0; c <= clip.range.c1; c++) {
            if (clip.sheet === sheet) sheet.setCell(r, c, undefined);
            else this.put(tx, clip.sheet, r, c, undefined);
          }
        }
        if (clip.sheet !== sheet) tx.snapshotSheet(clip.sheet);
        clip.sheet.setMerges(clip.sheet.merges.filter((m) => !clip.merges.some((x) => rangeEquals(x, m))));
      }
      for (let tr = 0; tr < tilesR; tr++) {
        for (let tc = 0; tc < tilesC; tc++) {
          const br = r0 + tr * h;
          const bc = c0 + tc * w;
          for (let i = 0; i < h; i++) {
            for (let j = 0; j < w; j++) {
              const r = br + i;
              const c = bc + j;
              if (r >= MAX_ROWS || c >= MAX_COLS) continue;
              const src = clip.cells[i][j];
              const existing = sheet.getCell(r, c);
              let cell: Cell | undefined;
              if (mode === 'formats') cell = src?.s || existing ? { ...(existing ?? { v: null }), s: src?.s } : undefined;
              else if (mode === 'values') cell = src && src.v !== null ? { v: src.v, s: existing?.s } : existing?.s ? { v: null, s: existing.s } : undefined;
              else if (src) {
                cell = { ...src };
                delete cell.spill;
                if (src.spill !== undefined && src.f === undefined) cell = src.v !== null ? { v: src.v, s: src.s } : src.s ? { v: null, s: src.s } : undefined;
                if (cell?.f !== undefined && !clip.cut) cell.f = translateFormula(cell.f, r - (clip.range.r0 + i), c - (clip.range.c0 + j));
                if (cell?.f !== undefined && clip.cut && clip.sheet !== sheet) cell.f = cell.f;
                if (cell && mode === 'formulas') cell.s = existing?.s;
                if (cell?.arr && !cell.dyn) cell.arr = undefined;
              }
              if (sheet === clip.sheet && clip.cut) sheet.setCell(r, c, cell);
              else this.put(tx, sheet, r, c, cell);
            }
          }
          if (mode !== 'values') {
            const moved = clip.merges.map((m) => makeRange(m.r0 - clip.range.r0 + br, m.c0 - clip.range.c0 + bc, m.r1 - clip.range.r0 + br, m.c1 - clip.range.c0 + bc));
            if (moved.length || merges) {
              const area = makeRange(br, bc, br + h - 1, bc + w - 1);
              sheet.setMerges([...sheet.merges.filter((m) => !rangesIntersect(m, area)), ...moved]);
            }
          }
        }
      }
      if (clip.cut) {
        // Formulas pointing at the moved block follow it; comments and links move too.
        const srcName = clip.sheet.name;
        this.adjustAllFormulas(tx, (f, s) => moveRefsInFormula(f, s.name, srcName, clip.range, sheet.name, dr, dc));
        const moveComments = clip.sheet.comments.filter((t) => rangeContains(clip.range, t.r, t.c));
        if (moveComments.length) {
          tx.snapshotSheet(clip.sheet);
          clip.sheet.comments = clip.sheet.comments.filter((t) => !moveComments.includes(t));
          tx.snapshotSheet(sheet);
          for (const t of moveComments) sheet.comments.push({ ...t, r: t.r + dr, c: t.c + dc });
        }
        this.clip = undefined;
      }
    }, { kind: clip.cut ? 'CutPaste' : 'DataChange' });
    this.select(makeRange(r0, c0, r0 + h * tilesR - 1, c0 + w * tilesC - 1), { r: r0, c: c0 });
  }

  // ---------------------------------------------------------------------------------------------
  // Fill

  /** Extends the selection to `target` with the fill handle. */
  fill(target: Range, copy = false): void {
    const sheet = this.sheet;
    const src = this.range;
    let axis: 'down' | 'up' | 'right' | 'left' | null = null;
    if (target.r1 > src.r1 && target.c0 === src.c0 && target.c1 === src.c1) axis = 'down';
    else if (target.r0 < src.r0 && target.c0 === src.c0 && target.c1 === src.c1) axis = 'up';
    else if (target.c1 > src.c1 && target.r0 === src.r0 && target.r1 === src.r1) axis = 'right';
    else if (target.c0 < src.c0 && target.r0 === src.r0 && target.r1 === src.r1) axis = 'left';
    if (!axis) return;
    const vertical = axis === 'down' || axis === 'up';
    const reverse = axis === 'up' || axis === 'left';
    const count = vertical ? (reverse ? src.r0 - target.r0 : target.r1 - src.r1) : reverse ? src.c0 - target.c0 : target.c1 - src.c1;
    this.transact('Fill', (tx) => {
      const lines = vertical ? src.c1 - src.c0 + 1 : src.r1 - src.r0 + 1;
      for (let l = 0; l < lines; l++) {
        const source: (Cell | undefined)[] = [];
        const len = vertical ? src.r1 - src.r0 + 1 : src.c1 - src.c0 + 1;
        for (let k = 0; k < len; k++) source.push(vertical ? sheet.getCell(src.r0 + k, src.c0 + l) : sheet.getCell(src.r0 + l, src.c0 + k));
        const out = fillSeries(source, count, {
          copy,
          reverse,
          isDate: (cell) => !!formatInfo(this.wb.styles.get(cell.s).numFmt)?.isDate,
          translate: (f, d) => (vertical ? translateFormula(f, d, 0) : translateFormula(f, 0, d)),
        });
        out.forEach((cell, k) => {
          const step = k + 1;
          const r = vertical ? (reverse ? src.r0 - step : src.r1 + step) : src.r0 + l;
          const c = vertical ? src.c0 + l : reverse ? src.c0 - step : src.c1 + step;
          if (r < 0 || c < 0 || r >= MAX_ROWS || c >= MAX_COLS) return;
          this.put(tx, sheet, r, c, cell);
        });
      }
    });
    this.select(makeRange(Math.min(src.r0, target.r0), Math.min(src.c0, target.c0), Math.max(src.r1, target.r1), Math.max(src.c1, target.c1)), this.selection.active);
  }

  // ---------------------------------------------------------------------------------------------
  // Find and replace

  find(query: string, opts: SearchOptions = {}): SearchHit[] {
    if (!query) return [];
    const hits: SearchHit[] = [];
    const q = opts.matchCase ? query : query.toLowerCase();
    const sheets = opts.allSheets === false ? [this.sheet] : this.wb.sheets.filter((s) => s.state === 'visible');
    for (const sheet of sheets) {
      const found: SearchHit[] = [];
      sheet.forEachCell((r, c, cell) => {
        const text = opts.inFormulas && cell.f !== undefined ? `=${cell.f}` : this.displayText(r, c, sheet);
        if (!text) return;
        const t = opts.matchCase ? text : text.toLowerCase();
        if (opts.wholeCell ? t === q : t.includes(q)) found.push({ sheet, r, c, text });
      });
      found.sort((a, b) => a.r - b.r || a.c - b.c);
      hits.push(...found);
    }
    return hits;
  }

  /** Replaces the query inside one cell (constants only). Returns true if it changed. */
  replace(hit: SearchHit, query: string, replacement: string, opts: SearchOptions = {}): boolean {
    const cell = hit.sheet.getCell(hit.r, hit.c);
    if (!cell || (cell.f !== undefined && !opts.inFormulas)) return false;
    const source = cell.f !== undefined ? `=${cell.f}` : this.editText(hit.r, hit.c, hit.sheet);
    const re = new RegExp(opts.wholeCell ? `^${escapeRe(query)}$` : escapeRe(query), opts.matchCase ? 'g' : 'gi');
    const next = source.replace(re, replacement.replace(/\$/g, '$$$$'));
    if (next === source) return false;
    this.setInput(hit.r, hit.c, next, hit.sheet);
    return true;
  }

  replaceAll(query: string, replacement: string, opts: SearchOptions = {}): number {
    let n = 0;
    this.transact('Replace all', () => {
      for (const hit of this.find(query, opts)) if (this.replace(hit, query, replacement, opts)) n++;
    });
    return n;
  }

  // ---------------------------------------------------------------------------------------------
  // Sorting

  sortRange(range: Range, byCol: number, ascending = true, hasHeader = false, sheet = this.sheet): void {
    const b = this.bounded(range, sheet) ?? range;
    const start = hasHeader ? b.r0 + 1 : b.r0;
    if (start >= b.r1) return;
    this.transact('Sort', (tx) => {
      const rows: { cells: (Cell | undefined)[]; key: Scalar; index: number }[] = [];
      for (let r = start; r <= b.r1; r++) {
        const cells: (Cell | undefined)[] = [];
        for (let c = b.c0; c <= b.c1; c++) cells.push(sheet.getCell(r, c));
        rows.push({ cells, key: sheet.getCell(r, byCol)?.v ?? null, index: r });
      }
      rows.sort((x, y) => {
        if (x.key === null && y.key !== null) return 1;
        if (y.key === null && x.key !== null) return -1;
        if (x.key instanceof CellError || y.key instanceof CellError) return x.key instanceof CellError ? 1 : -1;
        const cmp = compareScalars(x.key, y.key);
        return (ascending ? cmp : -cmp) || x.index - y.index;
      });
      rows.forEach((row, i) => {
        const r = start + i;
        row.cells.forEach((cell, j) => {
          const c = b.c0 + j;
          const moved = cell?.f !== undefined ? { ...cell, f: translateFormula(cell.f, r - row.index, 0) } : cell;
          this.put(tx, sheet, r, c, moved);
        });
      });
    });
  }

  // ---------------------------------------------------------------------------------------------
  // Auto filter

  /** Turns the auto filter on (around the selection) or off for the active sheet. */
  toggleAutoFilter(sheet = this.sheet): boolean {
    if (sheet.autoFilter) {
      this.transact('Remove filter', (tx) => {
        tx.snapshotSheet(sheet);
        this.unhideFiltered(sheet);
        sheet.autoFilter = undefined;
      });
      return false;
    }
    const sel = this.selectionOf(sheet);
    const region = filterRegion(sheet, sel.ranges[sel.ranges.length - 1], sel.active);
    if (!region) return false;
    this.transact('Filter', (tx) => {
      tx.snapshotSheet(sheet);
      sheet.autoFilter = { ref: region, columns: new Map() };
    });
    return true;
  }

  private unhideFiltered(sheet: Sheet): void {
    for (const r of sheet.filteredRows) {
      const info = sheet.rows.get(r);
      if (info?.hidden) sheet.setRowInfo(r, { ...info, hidden: undefined });
    }
    sheet.filteredRows = new Set();
  }

  /** Re-evaluates the filter criteria and hides the rows that do not match. */
  private applyFilterInTx(sheet: Sheet): void {
    const af = sheet.autoFilter;
    this.unhideFiltered(sheet);
    if (!af || !af.columns.size) return;
    const hide = rowsToHide(af, (r, c) => this.displayText(r, c, sheet), (r, c) => sheet.getCell(r, c)?.v ?? null);
    for (const r of hide) {
      const info = sheet.rows.get(r);
      if (info?.hidden) continue;
      sheet.setRowInfo(r, { ...(info ?? {}), hidden: true });
      sheet.filteredRows.add(r);
    }
    SheetLayout.of(sheet).invalidate();
  }

  /** Sets (or clears, with null) the criteria of one column of the auto filter. */
  setColumnFilter(col: number, filter: ColumnFilter | null, sheet = this.sheet): void {
    const af = sheet.autoFilter;
    if (!af) return;
    this.transact('Filter', (tx) => {
      tx.snapshotSheet(sheet);
      const columns = new Map(af.columns);
      if (filter && (filter.values || filter.conditions?.length)) columns.set(col, filter);
      else columns.delete(col);
      sheet.autoFilter = { ...af, columns };
      this.applyFilterInTx(sheet);
    });
  }

  clearFilters(sheet = this.sheet): void {
    const af = sheet.autoFilter;
    if (!af || !af.columns.size) return;
    this.transact('Clear filters', (tx) => {
      tx.snapshotSheet(sheet);
      sheet.autoFilter = { ...af, columns: new Map() };
      this.unhideFiltered(sheet);
    });
  }

  /** Sorts the rows of the filtered area (below the header) by one column. */
  sortFilterColumn(col: number, ascending: boolean, sheet = this.sheet): void {
    const af = sheet.autoFilter;
    if (!af) return;
    this.transact('Sort', (tx) => {
      this.sortRange(af.ref, col, ascending, true, sheet);
      tx.snapshotSheet(sheet);
      this.applyFilterInTx(sheet);
    });
  }

  /** Distinct displayed values offered in a column's filter list. */
  filterValues(col: number, sheet = this.sheet): { text: string; count: number; numeric: boolean }[] {
    const af = sheet.autoFilter;
    if (!af) return [];
    return columnValues(af, col, (r, c) => this.displayText(r, c, sheet), (r, c) => sheet.getCell(r, c)?.v ?? null);
  }

  // ---------------------------------------------------------------------------------------------
  // Data validation

  addValidation(rule: DataValidation, sheet = this.sheet): void {
    this.transact('Add validation', (tx) => {
      tx.snapshotSheet(sheet);
      sheet.validations = [...sheet.validations, this.syncFormatRule(sheet, { ...rule })];
    });
  }

  updateValidation(index: number, rule: DataValidation, sheet = this.sheet): void {
    const old = sheet.validations[index];
    if (!old) return;
    this.transact('Edit validation', (tx) => {
      tx.snapshotSheet(sheet);
      const next = this.syncFormatRule(sheet, { ...rule, raw: old.raw, cfId: old.cfId });
      sheet.validations = sheet.validations.map((v, i) => (i === index ? next : v));
    });
  }

  removeValidation(index: number, sheet = this.sheet): void {
    const old = sheet.validations[index];
    if (!old) return;
    this.transact('Remove validation', (tx) => {
      tx.snapshotSheet(sheet);
      if (old.cfId) sheet.condFormats = sheet.condFormats.filter((cf) => cf.id !== old.cfId);
      sheet.validations = sheet.validations.filter((_, i) => i !== index);
    });
  }

  /**
   * Keeps the conditional format of a "Format style" validation in step with the rule: invalid
   * (non-blank) cells get the style. Returns the rule with its `cfId`.
   */
  private syncFormatRule(sheet: Sheet, rule: DataValidation): DataValidation {
    if (rule.cfId) sheet.condFormats = sheet.condFormats.filter((cf) => cf.id !== rule.cfId);
    const style = rule.formatStyle;
    const first = rule.sqref[0];
    const valid = first ? validityFormula(rule, cellName(first.r0, first.c0)) : null;
    if (!style || !valid || !first) return { ...rule, cfId: undefined };
    const fontKids: XmlElement[] = [];
    if (style.bold) fontKids.push(xmlEl('b'));
    if (style.color) fontKids.push(xmlEl('color', { rgb: `FF${style.color.replace('#', '').toUpperCase()}` }));
    const dxfKids: XmlElement[] = [];
    if (fontKids.length) dxfKids.push(xmlEl('font', {}, fontKids));
    if (style.background) dxfKids.push(xmlEl('fill', {}, [xmlEl('patternFill', {}, [xmlEl('bgColor', { rgb: `FF${style.background.replace('#', '').toUpperCase()}` })])]));
    const dxfId = this.wb.dxfs.length;
    this.wb.dxfs = [...this.wb.dxfs, xmlEl('dxf', {}, dxfKids)];
    const cell = cellName(first.r0, first.c0);
    const formula = rule.allowBlank === false ? `NOT(${valid})` : `AND(LEN(${cell})>0,NOT(${valid}))`;
    const priority = 1 + Math.max(0, ...sheet.condFormats.flatMap((cf) => cf.raw.children.filter((k): k is XmlElement => k.type === 'element').map((k) => Number(k.attrs.priority) || 0)));
    const id = newDrawingId('cf');
    const raw = xmlEl('conditionalFormatting', { sqref: rule.sqref.map((r) => rangeName(r)).join(' ') }, [
      xmlEl('cfRule', { type: 'expression', dxfId, priority }, [xmlEl('formula', {}, [xmlText(formulaToFile(formula))])]),
    ]);
    sheet.condFormats = [...sheet.condFormats, { sqref: rule.sqref.map((r) => ({ ...r })), raw, id }];
    return { ...rule, cfId: id };
  }

  /** The value that typing `text` into a cell would store (for validation before committing). */
  parsedInput(r: number, c: number, text: string, sheet = this.sheet): Scalar {
    const style = this.styleAt(r, c, sheet);
    const parsed = parseInput(text, { textFormat: style.numFmt === '@', date1904: this.wb.date1904 });
    if (parsed.formula !== undefined) {
      const v = this.engine.evaluateText(normalizeFormulaInput(parsed.formula), sheet, r, c);
      return v === null || typeof v !== 'object' || v instanceof CellError ? (v as Scalar) : null;
    }
    return parsed.value;
  }

  /** The validation rule a typed entry breaks, if any. */
  invalidInput(r: number, c: number, text: string, sheet = this.sheet): DataValidation | null {
    const rule = validationAt(sheet, r, c);
    if (!rule || !rule.type || rule.type === 'none' || text === '') return null;
    const value = this.parsedInput(r, c, text, sheet);
    const display = (v: Scalar, s: Sheet, rr: number, cc: number) => formatValue(v, this.styleAt(rr, cc, s).numFmt, this.wb.date1904).text;
    return satisfies(this.engine, sheet, rule, r, c, value, display) ? null : rule;
  }

  /** Items of the list validation at a cell (null when the cell has no list). */
  validationListAt(r: number, c: number, sheet = this.sheet): string[] | null {
    const rule = validationAt(sheet, r, c);
    if (!rule || rule.type !== 'list' || rule.hideDropDown) return null;
    const display = (v: Scalar, s: Sheet, rr: number, cc: number) => formatValue(v, this.styleAt(rr, cc, s).numFmt, this.wb.date1904).text;
    return listItems(this.engine, sheet, rule, r, c, display);
  }

  // ---------------------------------------------------------------------------------------------
  // Defined names

  /** Adds or updates a workbook-level name; returns an error message or null. */
  setDefinedName(name: string, formula: string, previous?: string): string | null {
    const n = name.trim();
    if (!/^[A-Za-z_\\][A-Za-z0-9_.\\]*$/.test(n) || /^[A-Za-z]{1,3}\d+$/.test(n) || /^[rRcC]$/.test(n) || /^R\d*C\d*$/i.test(n)) return 'The name is not valid.';
    const f = formula.trim().replace(/^=/, '');
    if (!f) return 'Enter the cell range the name refers to.';
    const clash = this.wb.names.find((x) => x.localSheet === undefined && x.name.toUpperCase() === n.toUpperCase());
    if (clash && clash.name.toUpperCase() !== (previous ?? '').toUpperCase()) return 'This name already exists.';
    this.transact('Define name', (tx) => {
      tx.snapshotWorkbook(this.wb);
      const prevName = (previous ?? n).toUpperCase();
      const existing = this.wb.names.find((x) => x.localSheet === undefined && x.name.toUpperCase() === prevName);
      if (existing) this.wb.names = this.wb.names.map((x) => (x === existing ? { ...x, name: n, formula: f } : x));
      else this.wb.names = [...this.wb.names, { name: n, formula: f }];
      this.engine.invalidateAll();
    });
    return null;
  }

  deleteDefinedName(name: string): void {
    this.transact('Delete name', (tx) => {
      tx.snapshotWorkbook(this.wb);
      this.wb.names = this.wb.names.filter((x) => !(x.localSheet === undefined && x.name.toUpperCase() === name.toUpperCase()));
      this.engine.invalidateAll();
    });
  }

  // ---------------------------------------------------------------------------------------------
  // Notes (legacy comments)

  /** Sets the note of a cell; an empty text removes it. */
  setNote(r: number, c: number, text: string, sheet = this.sheet): void {
    this.commentTx(text ? 'Edit note' : 'Clear note', sheet, (s) => {
      const existing = s.commentAt(r, c);
      if (!text) {
        s.comments = s.comments.filter((t) => t !== existing);
        return;
      }
      if (existing) {
        const first = { ...existing.comments[0], text, author: existing.comments[0]?.author ?? this.author };
        s.comments = s.comments.map((t) => (t === existing ? { ...t, comments: [first, ...t.comments.slice(1)] } : t));
      } else {
        const id = newGuid();
        s.comments = [...s.comments, { id, r, c, kind: 'note', resolved: false, comments: [{ id, author: this.author, text, created: new Date().toISOString() }] }];
      }
    });
  }

  // ---------------------------------------------------------------------------------------------
  // Footer row and column titles (view features)

  setFooter(values: string[] | null, sheet = this.sheet): void {
    this.transact(values ? 'Footer' : 'Remove footer', (tx) => {
      tx.snapshotSheet(sheet);
      sheet.footer = values ?? undefined;
    });
  }

  setColumnTitle(c: number, title: string | null, sheet = this.sheet): void {
    this.transact('Rename column', (tx) => {
      tx.snapshotSheet(sheet);
      const titles = new Map(sheet.colTitles ?? []);
      if (title) titles.set(c, title);
      else titles.delete(c);
      sheet.colTitles = titles.size ? titles : undefined;
    });
  }

  // ---------------------------------------------------------------------------------------------
  // Charts and shapes

  private nextZ(sheet: Sheet): number {
    return Math.max(0, ...sheet.charts.map((x) => x.z), ...sheet.shapes.map((x) => x.z), ...sheet.images.map((x) => x.z ?? 0)) + 1;
  }

  /** Adds a chart; `rect` is in pixels at 100% zoom. */
  addChart(chart: Omit<SheetChart, 'id' | 'z' | 'from' | 'to'>, rect: { x: number; y: number; width: number; height: number }, sheet = this.sheet): SheetChart | null {
    let out: SheetChart | null = null;
    this.transact('Insert chart', (tx) => {
      tx.snapshotSheet(sheet);
      const layout = SheetLayout.of(sheet);
      out = { ...chart, id: newDrawingId('chart'), z: this.nextZ(sheet), from: layout.anchorAt(rect.x, rect.y), to: layout.anchorAt(rect.x + rect.width, rect.y + rect.height), dirty: true };
      sheet.charts = [...sheet.charts, out];
    });
    return out;
  }

  updateChart(id: string, patch: Partial<SheetChart>, sheet = this.sheet): void {
    if (!sheet.charts.some((c) => c.id === id)) return;
    this.transact('Edit chart', (tx) => {
      tx.snapshotSheet(sheet);
      sheet.charts = sheet.charts.map((c) => (c.id === id ? { ...c, ...patch, dirty: true } : c));
    });
  }

  addShape(shape: Omit<SheetShape, 'id' | 'z' | 'from' | 'to'>, rect: { x: number; y: number; width: number; height: number }, sheet = this.sheet): SheetShape | null {
    let out: SheetShape | null = null;
    this.transact('Insert shape', (tx) => {
      tx.snapshotSheet(sheet);
      const layout = SheetLayout.of(sheet);
      out = { ...shape, id: newDrawingId('shape'), z: this.nextZ(sheet), from: layout.anchorAt(rect.x, rect.y), to: layout.anchorAt(rect.x + rect.width, rect.y + rect.height), dirty: true };
      sheet.shapes = [...sheet.shapes, out];
    });
    return out;
  }

  updateShape(id: string, patch: Partial<SheetShape>, sheet = this.sheet): void {
    if (!sheet.shapes.some((s) => s.id === id)) return;
    this.transact('Edit shape', (tx) => {
      tx.snapshotSheet(sheet);
      sheet.shapes = sheet.shapes.map((s) => (s.id === id ? { ...s, ...patch, dirty: true } : s));
    });
  }

  /** The chart, shape or picture with this id. */
  objectById(id: string, sheet = this.sheet): { kind: 'chart'; obj: SheetChart } | { kind: 'shape'; obj: SheetShape } | { kind: 'image'; obj: SheetImage } | null {
    const chart = sheet.charts.find((c) => c.id === id);
    if (chart) return { kind: 'chart', obj: chart };
    const shape = sheet.shapes.find((s) => s.id === id);
    if (shape) return { kind: 'shape', obj: shape };
    const img = sheet.images.find((i) => i.id === id);
    return img ? { kind: 'image', obj: img } : null;
  }

  /** Moves or resizes a chart, shape or picture to a pixel rectangle (100% zoom). */
  setObjectRect(id: string, rect: { x: number; y: number; width: number; height: number }, sheet = this.sheet): void {
    const hit = this.objectById(id, sheet);
    if (!hit) return;
    if (hit.kind === 'image') {
      this.setImageRect(id, rect, sheet);
      return;
    }
    this.transact('Move object', (tx) => {
      tx.snapshotSheet(sheet);
      const layout = SheetLayout.of(sheet);
      const from = layout.anchorAt(rect.x, rect.y);
      const to = layout.anchorAt(rect.x + rect.width, rect.y + rect.height);
      if (hit.kind === 'chart') sheet.charts = sheet.charts.map((c) => (c.id === id ? { ...c, from, to } : c));
      else sheet.shapes = sheet.shapes.map((s) => (s.id === id ? { ...s, from, to } : s));
    });
  }

  deleteObject(id: string, sheet = this.sheet): void {
    const hit = this.objectById(id, sheet);
    if (!hit) return;
    this.transact(`Delete ${hit.kind}`, (tx) => {
      tx.snapshotSheet(sheet);
      sheet.charts = sheet.charts.filter((c) => c.id !== id);
      sheet.shapes = sheet.shapes.filter((s) => s.id !== id);
      sheet.images = sheet.images.filter((i) => i.id !== id);
    });
  }

  /** Changes the stacking order: to the very front/back, or one step forward/backward. */
  reorderObject(id: string, how: 'front' | 'back' | 'forward' | 'backward', sheet = this.sheet): void {
    const all = [
      ...sheet.charts.map((o) => ({ id: o.id, z: o.z })),
      ...sheet.shapes.map((o) => ({ id: o.id, z: o.z })),
      ...sheet.images.map((o, i) => ({ id: o.id, z: o.z ?? i })),
    ].sort((a, b) => a.z - b.z);
    const i = all.findIndex((o) => o.id === id);
    if (i < 0) return;
    const [item] = all.splice(i, 1);
    const at = how === 'front' ? all.length : how === 'back' ? 0 : how === 'forward' ? Math.min(all.length, i + 1) : Math.max(0, i - 1);
    all.splice(at, 0, item);
    const order = new Map(all.map((o, k) => [o.id, k + 1]));
    this.transact('Arrange', (tx) => {
      tx.snapshotSheet(sheet);
      sheet.charts = sheet.charts.map((c) => ({ ...c, z: order.get(c.id)! }));
      sheet.shapes = sheet.shapes.map((s) => ({ ...s, z: order.get(s.id)! }));
      sheet.images = sheet.images.map((img) => ({ ...img, z: order.get(img.id)! }));
    });
  }

  // ---------------------------------------------------------------------------------------------
  // Modes and files

  setMode(mode: EditMode): void {
    if (mode === this.mode) return;
    this.mode = mode;
    this.emit('mode', mode);
  }

  setFileName(name: string): void {
    this.fileName = name;
    this.emit('file', { name });
  }

  /** Writes the workbook as XLSX; `markSaved` clears the unsaved-changes flag (not for autosave). */
  save(markSaved = true, opts: { macros?: boolean } = {}): Uint8Array {
    for (const s of this.wb.sheets) {
      const sel = this.selections.get(s);
      if (sel) s.selection = { active: { ...sel.active }, ranges: sel.ranges.map((r) => ({ ...r })) } satisfies SheetSelection;
    }
    const bytes = writeXlsx(this.wb, opts);
    if (markSaved) {
      this.modified = false;
      this.emit('file', { name: this.fileName });
    }
    return bytes;
  }

  /** Recalculates volatile functions (NOW, TODAY, RAND...). */
  recalcVolatile(): void {
    const changed = this.engine.recalc();
    if (changed.size) this.emit('change', { structure: false, layout: false, ranges: [] });
  }
}

/** Applies `fn` to every reference of a chart; returns the changed fields, or null if none changed. */
export function rewriteChartRefs(chart: SheetChart, fn: (f: string) => string): Partial<SheetChart> | null {
  let changed = false;
  const map = <T extends string | undefined>(f: T): T => {
    if (f === undefined || f === '') return f;
    const g = fn(f);
    if (g !== f) changed = true;
    return g as T;
  };
  const series = chart.series.map((s) => ({ ...s, values: map(s.values), nameRef: map(s.nameRef), xValues: map(s.xValues), sizes: map(s.sizes) }));
  const range = map(chart.range);
  const categories = map(chart.categories);
  return changed ? { series, range, categories } : null;
}

function snapshotOf(sheet: Sheet) {
  // Local import to keep the top of the file readable.
  return snapshotSheetCopy(sheet);
}

import { snapshotSheet as snapshotSheetCopy } from './history';

function clampRow(r: number): number {
  return Math.max(0, Math.min(MAX_ROWS - 1, r));
}

function clampCol(c: number): number {
  return Math.max(0, Math.min(MAX_COLS - 1, c));
}

function normalizeNewlines(s: string): string {
  return s.replace(/\r\n?/g, '\n').replace(/\n$/, '');
}

function escapeRe(s: string): string {
  return s.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
}

function htmlToPatch(st: HtmlCellStyle): StylePatch {
  const patch: StylePatch = { font: {} };
  if (st.bold) patch.font!.bold = true;
  if (st.italic) patch.font!.italic = true;
  if (st.underline) patch.font!.underline = 'single';
  if (st.strike) patch.font!.strike = true;
  if (st.color) {
    const c = rgbSpec(st.color);
    if (c && !/^#?0{6}$/.test(st.color)) patch.font!.color = c;
  }
  if (st.fontSize) patch.font!.size = st.fontSize;
  if (st.background) {
    const c = rgbSpec(st.background);
    if (c) patch.fill = { pattern: 'solid', fg: c };
  }
  if (st.align) patch.align = { horizontal: st.align };
  if (!Object.keys(patch.font!).length) delete patch.font;
  return patch;
}

/** Adds or removes a decimal place in a number format; null when it cannot be adjusted. */
export function adjustDecimalFormat(fmt: string, delta: 1 | -1): string | null {
  if (/^General$/i.test(fmt) || /[dmyhs]/i.test(fmt.replace(/"[^"]*"/g, '').replace(/\[[^\]]*\]/g, '')) || fmt === '@') return null;
  const sections = fmt.split(';');
  const adjusted = sections.map((sec) => {
    // Find the last numeric placeholder group.
    const m = /([#0,]*0)(\.0*)?/.exec(sec);
    if (!m) return sec;
    const [whole, intPart, dec = ''] = m;
    let nextDec: string;
    if (delta > 0) nextDec = dec ? `${dec}0` : '.0';
    else nextDec = dec.length > 2 ? dec.slice(0, -1) : '';
    return sec.slice(0, m.index) + intPart + nextDec + sec.slice(m.index + whole.length);
  });
  const out = adjusted.join(';');
  return out === fmt ? null : out;
}

export { cellName, parseCellRef, cellUid, applyPatch };
