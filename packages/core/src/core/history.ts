/**
 * Undo/redo. A transaction records the previous state of what a command touches: single cells
 * for ordinary edits, a whole-sheet snapshot for structural changes (rows, columns, merges...) and
 * a workbook snapshot for sheet-level changes (add, remove, rename, reorder).
 */
import type { AutoFilter, Cell, ColInfo, CommentThread, ConditionalFormat, DataValidation, FreezePanes, Hyperlink, RowInfo, Sheet, SheetImage, SheetSelection } from './model/sheet';
import type { DefinedName, Person, Workbook } from './model/workbook';
import type { ColorSpec } from './model/color';
import { cloneChart, cloneShape, type SheetChart, type SheetShape } from './model/drawing';
import type { Range } from './refs';
import type { XmlElement } from './xml/xml';

function cloneFilter(f: AutoFilter | undefined): AutoFilter | undefined {
  if (!f) return undefined;
  const columns = new Map([...f.columns].map(([k, v]) => [k, { ...v, values: v.values ? [...v.values] : undefined, conditions: v.conditions?.map((c) => ({ ...c })) }]));
  return { ...f, ref: { ...f.ref }, columns };
}

export interface SheetSnapshot {
  cells: Map<number, Map<number, Cell>>;
  rows: Map<number, RowInfo>;
  cols: Map<number, ColInfo>;
  merges: Range[];
  hyperlinks: Hyperlink[];
  comments: CommentThread[];
  images: SheetImage[];
  charts: SheetChart[];
  shapes: SheetShape[];
  validations: DataValidation[];
  condFormats: ConditionalFormat[];
  autoFilter?: AutoFilter;
  filteredRows: Set<number>;
  footer?: string[];
  colTitles?: Map<number, string>;
  freeze: FreezePanes;
  name: string;
  state: Sheet['state'];
  tabColor?: ColorSpec;
  zoom: number;
  showGridLines: boolean;
  defaultColWidthPx: number;
  defaultRowHeightPt: number;
  otherAnchors: XmlElement[];
}

export function snapshotSheet(s: Sheet): SheetSnapshot {
  const cells = new Map<number, Map<number, Cell>>();
  for (const [r, row] of s.cells) cells.set(r, new Map(row));
  return {
    cells,
    rows: new Map([...s.rows].map(([k, v]) => [k, { ...v }])),
    cols: new Map([...s.cols].map(([k, v]) => [k, { ...v }])),
    merges: s.merges.map((m) => ({ ...m })),
    hyperlinks: s.hyperlinks.map((h) => ({ ...h, ref: { ...h.ref } })),
    comments: s.comments.map((t) => ({ ...t, comments: t.comments.map((c) => ({ ...c })) })),
    images: s.images.map((i) => ({ ...i, from: { ...i.from }, to: i.to ? { ...i.to } : undefined, ext: i.ext ? { ...i.ext } : undefined })),
    charts: s.charts.map(cloneChart),
    shapes: s.shapes.map(cloneShape),
    validations: s.validations.map((v) => ({ ...v, sqref: v.sqref.map((r) => ({ ...r })) })),
    condFormats: s.condFormats.map((c) => ({ ...c, sqref: c.sqref.map((r) => ({ ...r })) })),
    autoFilter: cloneFilter(s.autoFilter),
    filteredRows: new Set(s.filteredRows),
    footer: s.footer ? [...s.footer] : undefined,
    colTitles: s.colTitles ? new Map(s.colTitles) : undefined,
    freeze: { ...s.freeze },
    name: s.name,
    state: s.state,
    tabColor: s.tabColor,
    zoom: s.zoom,
    showGridLines: s.showGridLines,
    defaultColWidthPx: s.defaultColWidthPx,
    defaultRowHeightPt: s.defaultRowHeightPt,
    otherAnchors: [...s.otherAnchors],
  };
}

export function restoreSheet(s: Sheet, snap: SheetSnapshot): void {
  s.cells.clear();
  for (const [r, row] of snap.cells) s.cells.set(r, new Map(row));
  s.rows.clear();
  for (const [k, v] of snap.rows) s.rows.set(k, { ...v });
  s.cols.clear();
  for (const [k, v] of snap.cols) s.cols.set(k, { ...v });
  s.setMerges(snap.merges.map((m) => ({ ...m })));
  s.setHyperlinks(snap.hyperlinks.map((h) => ({ ...h, ref: { ...h.ref } })));
  s.comments = snap.comments.map((t) => ({ ...t, comments: t.comments.map((c) => ({ ...c })) }));
  s.images = snap.images.map((i) => ({ ...i, from: { ...i.from }, to: i.to ? { ...i.to } : undefined, ext: i.ext ? { ...i.ext } : undefined }));
  s.charts = snap.charts.map(cloneChart);
  s.shapes = snap.shapes.map(cloneShape);
  s.validations = snap.validations.map((v) => ({ ...v, sqref: v.sqref.map((r) => ({ ...r })) }));
  s.condFormats = snap.condFormats.map((c) => ({ ...c, sqref: c.sqref.map((r) => ({ ...r })) }));
  s.autoFilter = cloneFilter(snap.autoFilter);
  s.filteredRows = new Set(snap.filteredRows);
  s.footer = snap.footer ? [...snap.footer] : undefined;
  s.colTitles = snap.colTitles ? new Map(snap.colTitles) : undefined;
  s.freeze = { ...snap.freeze };
  s.name = snap.name;
  s.state = snap.state;
  s.tabColor = snap.tabColor;
  s.zoom = snap.zoom;
  s.showGridLines = snap.showGridLines;
  s.defaultColWidthPx = snap.defaultColWidthPx;
  s.defaultRowHeightPt = snap.defaultRowHeightPt;
  s.otherAnchors = [...snap.otherAnchors];
  s.layoutVersion++;
  s.invalidateIndexes();
}

export interface WorkbookSnapshot {
  sheets: Sheet[];
  names: DefinedName[];
  active: number;
  persons: Person[];
  date1904: boolean;
  chartRefs: Map<string, string[]>;
}

export function snapshotWorkbook(wb: Workbook): WorkbookSnapshot {
  return {
    sheets: [...wb.sheets],
    names: wb.names.map((n) => ({ ...n })),
    active: wb.activeSheetIndex,
    persons: wb.persons.map((p) => ({ ...p })),
    date1904: wb.date1904,
    chartRefs: new Map([...wb.chartRefs].map(([k, v]) => [k, [...v]])),
  };
}

export function restoreWorkbook(wb: Workbook, snap: WorkbookSnapshot): void {
  wb.sheets = [...snap.sheets];
  wb.names = snap.names.map((n) => ({ ...n }));
  wb.activeSheetIndex = Math.min(snap.active, wb.sheets.length - 1);
  wb.persons = snap.persons.map((p) => ({ ...p }));
  wb.date1904 = snap.date1904;
  wb.chartRefs = new Map([...snap.chartRefs].map(([k, v]) => [k, [...v]]));
}

export interface CellChange {
  sheet: Sheet;
  r: number;
  c: number;
  before: Cell | undefined;
  after: Cell | undefined;
}

export interface SelectionState {
  sheetIndex: number;
  selection?: SheetSelection;
}

export interface LineChange<T> {
  sheet: Sheet;
  index: number;
  before: T | undefined;
  after: T | undefined;
}

export class Transaction {
  readonly cells = new Map<Sheet, Map<number, CellChange>>();
  readonly rowInfos = new Map<Sheet, Map<number, LineChange<RowInfo>>>();
  readonly colInfos = new Map<Sheet, Map<number, LineChange<ColInfo>>>();
  readonly sheets = new Map<Sheet, { before: SheetSnapshot; after?: SheetSnapshot }>();
  workbook?: { before: WorkbookSnapshot; after?: WorkbookSnapshot };
  selectionBefore?: SelectionState;
  selectionAfter?: SelectionState;
  /** Kind of change for listeners (e.g. DataChangeType of the compatibility API). */
  kind = 'DataChange';

  constructor(readonly label: string) {}

  get structural(): boolean {
    return this.sheets.size > 0 || this.workbook !== undefined;
  }

  isEmpty(): boolean {
    if (this.workbook || this.sheets.size) return false;
    for (const m of this.cells.values()) if (m.size) return false;
    for (const m of this.rowInfos.values()) if (m.size) return false;
    for (const m of this.colInfos.values()) if (m.size) return false;
    return true;
  }

  touchRow(sheet: Sheet, r: number): void {
    if (this.sheets.has(sheet)) return;
    let m = this.rowInfos.get(sheet);
    if (!m) this.rowInfos.set(sheet, (m = new Map()));
    if (!m.has(r)) {
      const info = sheet.rows.get(r);
      m.set(r, { sheet, index: r, before: info ? { ...info } : undefined, after: undefined });
    }
  }

  touchCol(sheet: Sheet, c: number): void {
    if (this.sheets.has(sheet)) return;
    let m = this.colInfos.get(sheet);
    if (!m) this.colInfos.set(sheet, (m = new Map()));
    if (!m.has(c)) {
      const info = sheet.cols.get(c);
      m.set(c, { sheet, index: c, before: info ? { ...info } : undefined, after: undefined });
    }
  }

  /** Records the state of a cell before its first change in this transaction. */
  touch(sheet: Sheet, r: number, c: number): void {
    if (this.sheets.has(sheet)) return;
    let m = this.cells.get(sheet);
    if (!m) this.cells.set(sheet, (m = new Map()));
    const key = r * 16384 + c;
    if (!m.has(key)) m.set(key, { sheet, r, c, before: sheet.getCell(r, c), after: undefined });
  }

  snapshotSheet(sheet: Sheet): void {
    if (!this.sheets.has(sheet)) {
      // Cell-level records for this sheet are superseded by the snapshot, but their "before"
      // values must win: restore them into the snapshot.
      const snap = snapshotSheet(sheet);
      const cells = this.cells.get(sheet);
      if (cells) {
        for (const ch of cells.values()) {
          let row = snap.cells.get(ch.r);
          if (!row) snap.cells.set(ch.r, (row = new Map()));
          if (ch.before) row.set(ch.c, ch.before);
          else row.delete(ch.c);
          if (!row.size) snap.cells.delete(ch.r);
        }
        this.cells.delete(sheet);
      }
      for (const ch of this.rowInfos.get(sheet)?.values() ?? []) {
        if (ch.before) snap.rows.set(ch.index, { ...ch.before });
        else snap.rows.delete(ch.index);
      }
      for (const ch of this.colInfos.get(sheet)?.values() ?? []) {
        if (ch.before) snap.cols.set(ch.index, { ...ch.before });
        else snap.cols.delete(ch.index);
      }
      this.rowInfos.delete(sheet);
      this.colInfos.delete(sheet);
      this.sheets.set(sheet, { before: snap });
    }
  }

  snapshotWorkbook(wb: Workbook): void {
    if (!this.workbook) this.workbook = { before: snapshotWorkbook(wb) };
  }

  /** Captures the "after" states. */
  finish(wb: Workbook): void {
    for (const m of this.cells.values()) for (const ch of m.values()) ch.after = ch.sheet.getCell(ch.r, ch.c);
    for (const m of this.rowInfos.values()) for (const ch of m.values()) ch.after = ch.sheet.rows.get(ch.index) ? { ...ch.sheet.rows.get(ch.index)! } : undefined;
    for (const m of this.colInfos.values()) for (const ch of m.values()) ch.after = ch.sheet.cols.get(ch.index) ? { ...ch.sheet.cols.get(ch.index)! } : undefined;
    for (const [sheet, s] of this.sheets) s.after = snapshotSheet(sheet);
    if (this.workbook) this.workbook.after = snapshotWorkbook(wb);
  }
}

export class History {
  private undoStack: Transaction[] = [];
  private redoStack: Transaction[] = [];
  limit = 200;

  push(t: Transaction): void {
    this.undoStack.push(t);
    if (this.undoStack.length > this.limit) this.undoStack.shift();
    this.redoStack = [];
  }

  canUndo(): boolean {
    return this.undoStack.length > 0;
  }

  canRedo(): boolean {
    return this.redoStack.length > 0;
  }

  popUndo(): Transaction | undefined {
    const t = this.undoStack.pop();
    if (t) this.redoStack.push(t);
    return t;
  }

  popRedo(): Transaction | undefined {
    const t = this.redoStack.pop();
    if (t) this.undoStack.push(t);
    return t;
  }

  clear(): void {
    this.undoStack = [];
    this.redoStack = [];
  }
}
