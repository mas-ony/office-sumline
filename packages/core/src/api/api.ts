/**
 * The programmatic interface of an editor (getValue/setValue, getData/setData, insertRow,
 * orderBy...), given to `onReady` and put on `window.spreadsheetEditor` by `installPageApi`. It follows the workbook the editor shows.
 * Useful for automation and tests.
 */
import type { SpreadsheetDocument } from '../core/document';
import type { GridView } from '../grid/GridView';
import type { DownloadFormat } from '../ui/context';
import type { SumlineUI } from '../ui/controller';
import { version } from '../../package.json';
import { isFileHandle } from './source';
import { parseRange, rangeName, splitSheetRef, type Range } from '../core/refs';
import { rgbSpec } from '../core/model/color';
import type { StylePatch } from '../core/model/style';
import type { Scalar } from '../core/values';

/** Events of the page API, with what their listeners receive. */
export interface ApiEvents {
  /** Cells or the workbook's structure changed. */
  change: unknown;
  selection: unknown;
  activeSheet: unknown;
  sheets: unknown;
  /** The workbook was renamed or saved. */
  file: { name: string };
  /** Another workbook was opened. */
  open: { fileName: string };
}

export interface SpreadsheetApi {
  /** The editor's version: that of `@sumline/core`, read from its package.json. */
  version: string;
  /** The open workbook (the native document model). */
  readonly doc: SpreadsheetDocument;
  readonly grid: GridView | null;
  getValue(ref: string): Scalar;
  getText(ref: string): string;
  setValue(ref: string, value: string | number | boolean | null): void;
  getData(sheet?: string): Scalar[][];
  setData(data: (string | number | boolean | null)[][], at?: string): void;
  insertRow(count: number, row: number, before?: boolean): void;
  insertColumn(count: number, col: number, before?: boolean): void;
  deleteRow(row: number, count?: number): void;
  deleteColumn(col: number, count?: number): void;
  setStyle(ref: string, property: string, value: string): void;
  setFormat(ref: string, code: string): void;
  orderBy(col: number, descending?: boolean): void;
  setFreezeRows(n: number): void;
  setFreezeColumns(n: number): void;
  getSelected(): string;
  select(ref: string): void;
  undo(): void;
  redo(): void;
  createWorksheet(name?: string): string | null;
  renameWorksheet(index: number, name: string): string | null;
  deleteWorksheet(index: number): boolean;
  worksheets(): string[];
  /**
   * Opens a workbook (XLSX, XLSM, XLTX, XLS, XLSB or CSV) in place of the current one; the API then
   * serves the new workbook. Given a file handle (File System Access API), Save writes back to that
   * file where the browser allows.
   */
  open(file: Blob | ArrayBuffer | Uint8Array | FileSystemFileHandle, name?: string): Promise<void>;
  /** Downloads a copy (xlsx, csv, tsv) or opens the print preview for pdf. */
  download(format: DownloadFormat): void;
  save(): void;
  /** Listen to an event; returns the function that removes the listener. */
  on<K extends keyof ApiEvents>(event: K, fn: (e: ApiEvents[K]) => void): () => void;
}

function resolve(doc: SpreadsheetDocument, ref: string): { sheetIndex: number; range: Range } {
  const { sheet, ref: local } = splitSheetRef(ref);
  const index = sheet ? doc.wb.sheets.findIndex((s) => s.name.toLowerCase() === sheet.toLowerCase()) : doc.activeIndex;
  const range = parseRange(local.toUpperCase().replace(/\$/g, ''));
  if (index < 0 || !range) throw new Error(`Invalid reference: ${ref}`);
  return { sheetIndex: index, range };
}

/** The CSS-like style properties that `setStyle` takes, mapped to cell style changes. */
function cssToPatch(property: string, value: string): StylePatch | null {
  const v = value.trim();
  switch (property.toLowerCase()) {
    case 'font-weight':
      return { font: { bold: v === 'bold' || Number(v) >= 600 ? true : null } };
    case 'font-style':
      return { font: { italic: v === 'italic' ? true : null } };
    case 'text-decoration':
      return { font: { underline: v.includes('underline') ? 'single' : null, strike: v.includes('line-through') ? true : null } };
    case 'color':
      return { font: { color: rgbSpec(v) ?? null } };
    case 'background-color':
    case 'background':
      return { fill: rgbSpec(v) ? { pattern: 'solid', fg: rgbSpec(v)! } : null };
    case 'font-size': {
      const m = /^([\d.]+)(px|pt)?$/.exec(v);
      if (!m) return null;
      return { font: { size: m[2] === 'px' ? Math.round(Number(m[1]) * 0.75 * 2) / 2 : Number(m[1]) } };
    }
    case 'font-family':
      return { font: { name: v.replace(/["']/g, '').split(',')[0].trim() } };
    case 'text-align':
      return { align: { horizontal: v === 'left' || v === 'center' || v === 'right' || v === 'justify' ? v : null } };
    case 'vertical-align':
      return { align: { vertical: v === 'top' ? 'top' : v === 'middle' ? 'center' : v === 'bottom' ? 'bottom' : null } };
    case 'white-space':
      return { align: { wrap: v === 'normal' || v === 'pre-wrap' ? true : null } };
    default:
      return null;
  }
}

function noWorkbook(): never {
  throw new Error('No workbook is open');
}

/** The page API of an editor. */
export function createApi(ui: SumlineUI): SpreadsheetApi {
  const files = ui.files;
  // The editor's start screen has no workbook; the API is handed out once one is open.
  const current = () => ui.doc ?? noWorkbook();
  const api: SpreadsheetApi = {
    version,
    get doc() {
      return current();
    },
    get grid() {
      return ui.grid;
    },
    getValue(ref) {
      const doc = current();
      const { sheetIndex, range } = resolve(doc, ref);
      return doc.wb.sheets[sheetIndex].getCell(range.r0, range.c0)?.v ?? null;
    },
    getText(ref) {
      const doc = current();
      const { sheetIndex, range } = resolve(doc, ref);
      return doc.displayText(range.r0, range.c0, doc.wb.sheets[sheetIndex]);
    },
    setValue(ref, value) {
      const doc = current();
      const { sheetIndex, range } = resolve(doc, ref);
      doc.setInput(range.r0, range.c0, value === null ? '' : typeof value === 'boolean' ? (value ? 'TRUE' : 'FALSE') : String(value), doc.wb.sheets[sheetIndex]);
    },
    getData(sheetName) {
      const doc = current();
      const sheet = sheetName ? doc.wb.sheetByName(sheetName) ?? doc.sheet : doc.sheet;
      const used = sheet.usedRange(false);
      if (!used) return [];
      const out: Scalar[][] = [];
      for (let r = 0; r <= used.r1; r++) {
        const row: Scalar[] = [];
        for (let c = 0; c <= used.c1; c++) row.push(sheet.getCell(r, c)?.v ?? null);
        out.push(row);
      }
      return out;
    },
    setData(data, at = 'A1') {
      const doc = current();
      const { sheetIndex, range } = resolve(doc, at);
      doc.setValues(range.r0, range.c0, data, doc.wb.sheets[sheetIndex]);
    },
    insertRow(count, row, before = false) {
      const doc = current();
      doc.insertRows(before ? row : row + 1, count);
    },
    insertColumn(count, col, before = false) {
      const doc = current();
      doc.insertCols(before ? col : col + 1, count);
    },
    deleteRow(row, count = 1) {
      const doc = current();
      doc.deleteRows(row, count);
    },
    deleteColumn(col, count = 1) {
      const doc = current();
      doc.deleteCols(col, count);
    },
    setStyle(ref, property, value) {
      const doc = current();
      const patch = cssToPatch(property, value);
      if (!patch) throw new Error(`Unsupported style property: ${property}`);
      const { sheetIndex, range } = resolve(doc, ref);
      doc.format(patch, [range], doc.wb.sheets[sheetIndex], 'Style');
    },
    setFormat(ref, code) {
      const doc = current();
      const { sheetIndex, range } = resolve(doc, ref);
      doc.format({ numFmt: code }, [range], doc.wb.sheets[sheetIndex], 'Number format');
    },
    orderBy(col, descending = false) {
      const doc = current();
      const used = doc.sheet.usedRange(false);
      if (used) doc.sortRange(used, col, !descending, false);
    },
    setFreezeRows(n) {
      const doc = current();
      doc.setFreeze(n, doc.sheet.freeze.cols);
    },
    setFreezeColumns(n) {
      const doc = current();
      doc.setFreeze(doc.sheet.freeze.rows, n);
    },
    getSelected() {
      const doc = current();
      return rangeName(doc.range);
    },
    select(ref) {
      const doc = current();
      if (!doc.goTo(ref)) throw new Error(`Invalid reference: ${ref}`);
    },
    undo: () => current().undo(),
    redo: () => current().redo(),
    createWorksheet(name) {
      const doc = current();
      return doc.addSheet(name)?.name ?? null;
    },
    renameWorksheet: (index, name) => current().renameSheet(index, name),
    deleteWorksheet: (index) => current().removeSheet(index),
    worksheets: () => current().wb.sheets.map((s) => s.name),
    async open(file, name) {
      if (isFileHandle(file)) {
        if (!(await ui.openFile(await file.getFile(), file))) throw new Error(`Could not open ${file.name}`);
        await ui.whenGrid();
        return;
      }
      const bytes = file instanceof Uint8Array ? file : file instanceof ArrayBuffer ? new Uint8Array(file) : new Uint8Array(await file.arrayBuffer());
      const fileName = name ?? (typeof File !== 'undefined' && file instanceof File ? file.name : 'Workbook.xlsx');
      if (!(await ui.openBytes(bytes, fileName))) throw new Error(`Could not open ${fileName}`);
      // Resolve once the new workbook is shown.
      await ui.whenGrid();
    },
    download: (format) => files.download(format),
    save: () => files.save(),
    on: (event, fn) => ui.on(event, fn),
  };
  return api;
}

/** Puts an editor's API on `window.spreadsheetEditor`; returns the function that removes it. */
export function installPageApi(api: SpreadsheetApi): () => void {
  const w = window as unknown as { spreadsheetEditor?: SpreadsheetApi };
  w.spreadsheetEditor = api;
  return () => {
    if (w.spreadsheetEditor === api) delete w.spreadsheetEditor;
  };
}
