/**
 * The editor without its views: the state the React and Vue interfaces render, the actions their
 * controls call, and what needs no view (the grid and its callbacks, file actions, autosave and
 * recovery, drag and drop, the phone layout). A view creates one, renders `getState()` and `ctx`,
 * calls `attach()` once its root element exists, `attachGrid()` with the grid's element (again for
 * each workbook opened: `docKey` changes) and `detach()` when it goes away. While the start screen
 * shows there is no workbook: `doc` is null and so is the context.
 */
import type { SpreadsheetDocument } from '../core/document';
import { sheetToCsv } from '../core/csv';
import { GridView, type EditState } from '../grid/GridView';
import type { ViewTransform } from '../grid/viewport';
import { clearMeasureCache } from '../grid/fonts';
import { clearTextCache } from '../grid/render';
import { ACCEPT, newDocument, openDocument } from '../api/open';
import { clearAutosave, loadAutosave, saveAutosave, type AutosaveRecord } from '../api/autosave';
import { createApi, type ApiEvents, type SpreadsheetApi } from '../api/api';
import { documentOf, downloadBlob, isFileHandle, nameOfUrl, type DocumentInput, type DocumentSource } from '../api/source';
import type { ConfirmOptions, DialogSpec, DownloadFormat, FileActions, Prefs, UIContextValue } from './context';
import { computeDocState, DOC_STATE_EVENTS, type DocState } from './docState';
import { isTouchDevice, matchesMedia, PHONE_QUERY, watchMedia } from './dom';
import { FormulaAssistState } from './formulaAssist';
import type { MenuAnchor, MenuItem } from './menuItems';
import { cellMenu, columnMenu, objectMenu, rowMenu, topMenus } from './menus';
import * as cmd from './commands';

/** Options of the SumlineEditor components (React and Vue) and of mount(). */
export interface SumlineOptions {
  /** Opened at start, and again whenever it changes. */
  document?: DocumentInput;
  /**
   * Start on the start screen (a drop zone, "Create new spreadsheet" and "Upload file") instead of
   * a blank workbook; it also shows while `document` loads. Read at start.
   */
  startScreen?: boolean;
  /** Workbooks open view-only. */
  readOnly?: boolean;
  /** Shown on comments and notes. */
  author?: string;
  /** Keeps the recent colours and print options in localStorage under `<storageKey>-…`. */
  storageKey?: string;
  /** Keeps unsaved changes in an IndexedDB database of this name, to recover them after a crash. */
  autosave?: string;
  /** With `autosave` and `startScreen`: the start screen offers to recover unsaved changes (default true). */
  offerRecovery?: boolean;
  /** Replace some of the built-in file actions (New, Open, Save, Download, Print…). */
  files?: Partial<FileActions>;
  /** Called once a workbook is open (the initial document, or the first one opened from the start screen). */
  onReady?: (api: SpreadsheetApi) => void;
}

/** A question shown as a dialog (confirm or a text prompt). */
export type PendingQuestion =
  | { kind: 'confirm'; message: string; opts?: ConfirmOptions; resolve: (v: boolean) => void }
  | { kind: 'prompt'; message: string; value: string; opts?: ConfirmOptions; resolve: (v: string | null) => void };

/** A menu or list opened by the menu bar or the grid. */
export interface OpenMenu {
  items: MenuItem[];
  at: MenuAnchor;
  title?: string;
  testId?: string;
}

/** State of the interface around the workbook. */
export interface UIState {
  /** The open workbook; null while the start screen shows. */
  doc: SpreadsheetDocument | null;
  /** Changes with each workbook opened: views draw the editor for it afresh. */
  docKey: number;
  /** Counts the document's events; the document state is computed for it. */
  docVersion: number;
  grid: GridView | null;
  /** The cell edit in progress (mirrored by the formula bar). */
  edit: EditState;
  phone: boolean;
  dialog: DialogSpec | null;
  menu: OpenMenu | null;
  /** The menu bar's menu that is open (its index). */
  topMenu: number | null;
  toast: string | null;
  pending: PendingQuestion | null;
  comments: { r: number; c: number; rect: DOMRect } | null;
  note: { r: number; c: number; rect: DOMRect } | null;
  filter: { col: number; at: { x: number; y: number } } | null;
  selectedObject: string | null;
  /** Unsaved work found by autosave, offered for recovery on the start screen. */
  restore: { fileName: string; savedAt: number } | null;
  /** A message that stays until dismissed (a file that could not be opened). */
  notice: string | null;
  /** A file is being read. */
  busy: string | null;
  /** Files are dragged over the editor. */
  dragging: boolean;
  /** A workbook is open and the page API is available. */
  ready: boolean;
}

export const DISCARD_PROMPT = 'Discard unsaved changes to the current workbook?';

const NO_EDIT: EditState = { active: false, text: '', mode: 'enter', origin: 'cell', caret: 0, r: 0, c: 0 };
const TOAST_MS = 5000;
const AUTOSAVE_DELAY = 2500;
const AUTOSAVE_MAX_CELLS = 250_000;
const RECOVERY_MS = 1000 * 60 * 60 * 24 * 14;
const XLSX_MIME = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet';
const XLSM_MIME = 'application/vnd.ms-excel.sheet.macroEnabled.12';

interface FilePickerWindow {
  showOpenFilePicker?: (opts: unknown) => Promise<FileSystemFileHandle[]>;
}

/** The state of the interface that belongs to one workbook (reset when another opens). */
const PER_DOCUMENT = {
  grid: null,
  edit: NO_EDIT,
  dialog: null,
  menu: null,
  topMenu: null,
  toast: null,
  pending: null,
  comments: null,
  note: null,
  filter: null,
  selectedObject: null,
} satisfies Partial<UIState>;

function cellCount(doc: SpreadsheetDocument): number {
  let n = 0;
  for (const sheet of doc.wb.sheets) for (const row of sheet.cells.values()) n += row.size;
  return n;
}

const baseName = (d: SpreadsheetDocument) => d.fileName.replace(/\.(xlsx|xlsm|xls|xlsb|csv|tsv|txt)$/i, '') || 'Untitled spreadsheet';

const isMacro = (d: SpreadsheetDocument) => /\.xlsm$/i.test(d.fileName);

const isSheetFile = (f: File) => /\.(xlsx|xlsm|xltx|xlsb|xls|csv|tsv)$/i.test(f.name) || f.type === XLSX_MIME || f.type === 'text/csv';

function xlsxBlob(d: SpreadsheetDocument, macros: boolean, mark = true): Blob {
  return new Blob([d.save(mark, { macros }) as BlobPart], { type: macros ? XLSM_MIME : XLSX_MIME });
}

function sameDocument(a: DocumentInput | undefined, b: DocumentInput | undefined): boolean {
  if (a === b) return true;
  if (!a || !b) return false;
  const x = documentOf(a);
  const y = documentOf(b);
  return x.source === y.source && x.name === y.name;
}

let fontsWatched = false;

/** Text is measured on a canvas: once the bundled fonts finish loading, measurements change. */
function watchFonts(): void {
  if (fontsWatched || typeof document === 'undefined') return;
  fontsWatched = true;
  void document.fonts?.ready.then(() => {
    clearMeasureCache();
    clearTextCache();
    window.dispatchEvent(new Event('resize'));
  });
}

export class SumlineUI {
  options: SumlineOptions;
  /** The page API of this editor (it follows the open workbook). */
  readonly api: SpreadsheetApi;
  readonly files: FileActions;
  readonly prefs: Prefs;
  /** Container on the page's body for menus, dialogs and messages, so that they overlay everything. Created by attach(). */
  readonly layer: HTMLDivElement = document.createElement('div');
  /** The suggestion list and argument hint of formula editing. */
  assist: FormulaAssistState;
  private state: UIState;
  private ctxCache: UIContextValue | null = null;
  private docStateCache: { doc: SpreadsheetDocument; version: number; value: DocState } | null = null;
  private listeners = new Set<() => void>();
  private paintListeners = new Set<(vt: ViewTransform) => void>();
  private apiListeners = new Map<keyof ApiEvents, Set<(e: unknown) => void>>();
  private offs: (() => void)[] = [];
  private docOffs: (() => void)[] = [];
  private root: HTMLElement | null = null;
  private gridHost: HTMLElement | null = null;
  private input: HTMLInputElement | null = null;
  private handle: FileSystemFileHandle | null = null;
  private recovery: AutosaveRecord | null = null;
  private memory = new Map<string, unknown>();
  private toastTimer: ReturnType<typeof setTimeout> | null = null;
  private autosaveTimer: ReturnType<typeof setTimeout> | null = null;
  private autosaveOff = false;
  private gridWaiters: (() => void)[] = [];
  private started = false;

  constructor(options: SumlineOptions = {}) {
    this.options = { ...options };
    this.prefs = { get: (name, initial) => this.getPref(name, initial), set: (name, value) => this.setPref(name, value) };
    const doc = options.startScreen ? null : this.prepare(newDocument());
    this.state = {
      doc,
      docKey: 1,
      docVersion: 0,
      ...PER_DOCUMENT,
      phone: matchesMedia(PHONE_QUERY),
      restore: null,
      notice: null,
      busy: null,
      dragging: false,
      ready: false,
    };
    this.assist = this.makeAssist(doc);
    this.files = this.makeFiles();
    this.api = createApi(this);
    if (doc) this.wireDocument(doc);
  }

  // ---------------------------------------------------------------------------------------------
  // State

  subscribe = (fn: () => void): (() => void) => {
    this.listeners.add(fn);
    return () => this.listeners.delete(fn);
  };

  getState = (): UIState => this.state;

  /** The context the controls work with (null while the start screen shows); a new object after every change. */
  getCtx = (): UIContextValue | null => {
    const doc = this.state.doc;
    if (!doc) return null;
    return (this.ctxCache ??= this.makeCtx(doc));
  };

  /** The context, for actions that need a workbook. */
  private get ctx(): UIContextValue {
    const ctx = this.getCtx();
    if (!ctx) throw new Error('No workbook is open');
    return ctx;
  }

  get doc(): SpreadsheetDocument | null {
    return this.state.doc;
  }

  /** The interface's root element while attached. */
  get rootElement(): HTMLElement | null {
    return this.root;
  }

  get grid(): GridView | null {
    return this.state.grid;
  }

  set(patch: Partial<UIState>): void {
    const keys = Object.keys(patch) as (keyof UIState)[];
    if (keys.every((k) => Object.is(this.state[k], patch[k]))) return;
    this.state = { ...this.state, ...patch };
    this.ctxCache = null;
    this.notify();
  }

  private notify(): void {
    for (const fn of [...this.listeners]) fn();
  }

  /** Whether a node belongs to this interface (its root or its layer). */
  contains(node: Node | null): boolean {
    return !!node && (!!this.root?.contains(node) || this.layer.contains(node));
  }

  /** The document state for the current version. */
  docState(): DocState {
    const { doc, docVersion } = this.state;
    if (!doc) throw new Error('No workbook is open');
    const c = this.docStateCache;
    if (c && c.doc === doc && c.version === docVersion) return c.value;
    const value = computeDocState(doc, docVersion);
    this.docStateCache = { doc, version: docVersion, value };
    return value;
  }

  private makeCtx(doc: SpreadsheetDocument): UIContextValue {
    const s = this.state;
    return {
      ui: this,
      layer: this.layer,
      prefs: this.prefs,
      doc,
      grid: s.grid,
      state: this.docState(),
      edit: s.edit,
      phone: s.phone,
      files: this.files,
      dialog: s.dialog,
      openDialog: this.openDialog,
      closeDialog: this.closeDialog,
      notify: this.notify_,
      confirm: this.confirm,
      prompt: this.prompt,
      openComments: this.openComments,
      openNote: this.openNote,
      openFilter: this.openFilter,
      selectedObject: s.selectedObject,
      selectObject: this.selectObject,
    };
  }

  /** Listen to the grid's repaints (the footer row follows the columns); returns the function that stops. */
  onGridPaint(fn: (vt: ViewTransform) => void): () => void {
    this.paintListeners.add(fn);
    return () => this.paintListeners.delete(fn);
  }

  // ---------------------------------------------------------------------------------------------
  // Remembered values

  private key(name: string): string | undefined {
    return this.options.storageKey && `${this.options.storageKey}-${name}`;
  }

  private getPref<T>(name: string, initial: T): T {
    if (this.memory.has(name)) return this.memory.get(name) as T;
    const key = this.key(name);
    if (!key) return initial;
    try {
      const raw = localStorage.getItem(key);
      return raw === null ? initial : (JSON.parse(raw) as T);
    } catch {
      return initial;
    }
  }

  private setPref(name: string, value: unknown): void {
    this.memory.set(name, value);
    const key = this.key(name);
    if (!key) return;
    try {
      localStorage.setItem(key, JSON.stringify(value));
    } catch {
      // Optional.
    }
  }

  private author(): string {
    if (this.options.author) return this.options.author;
    const key = this.key('author');
    try {
      return (key && localStorage.getItem(key)) || 'Guest';
    } catch {
      return 'Guest';
    }
  }

  // ---------------------------------------------------------------------------------------------
  // Actions

  /** A message at the bottom that goes away by itself. */
  notify_ = (message: string): void => {
    this.set({ toast: message });
    if (this.toastTimer) clearTimeout(this.toastTimer);
    this.toastTimer = setTimeout(() => this.set({ toast: null }), TOAST_MS);
  };

  dismissToast = (): void => this.set({ toast: null });

  /** Show a message that stays until dismissed (null removes it). */
  notice = (message: string | null): void => this.set({ notice: message });

  confirm = (message: string, opts?: ConfirmOptions): Promise<boolean> => new Promise((resolve) => this.set({ pending: { kind: 'confirm', message, opts, resolve } }));

  prompt = (message: string, value = '', opts?: ConfirmOptions): Promise<string | null> => new Promise((resolve) => this.set({ pending: { kind: 'prompt', message, value, opts, resolve } }));

  /** Answer the question shown (a boolean for a confirm, the text or null for a prompt). */
  answer = (value: boolean | string | null): void => {
    const p = this.state.pending;
    if (!p) return;
    this.set({ pending: null });
    (p.resolve as (v: typeof value) => void)(value);
  };

  /** Client rectangle of a cell (merged cells: the whole merge). */
  cellRect(r: number, c: number): DOMRect {
    const { grid: g, doc } = this.state;
    if (!g || !doc) return new DOMRect(100, 100, 80, 24);
    const vt = g.transform();
    const m = doc.sheet.mergeAt(r, c);
    const rect = m ? vt.rangeRect(m) : vt.cellRect(r, c);
    const root = g.root.getBoundingClientRect();
    return new DOMRect(root.left + rect.x, root.top + rect.y, rect.w, rect.h);
  }

  private focusGrid(): void {
    this.state.grid?.focusIdle();
  }

  openDialog = (d: DialogSpec): void => {
    const grid = this.state.grid;
    if (grid?.isEditing()) {
      // The function list adds to the formula being typed, at its caret, so that edit stays open
      // behind it (without the focus, so typing cannot change it unseen). Other dialogs end the edit.
      const formula = d.kind === 'functions' && grid.editState.text.startsWith('=');
      if (formula) (document.activeElement as HTMLElement | null)?.blur?.();
      else if (d.kind !== 'search') grid.commitEdit();
    }
    this.set({ dialog: d });
  };

  closeDialog = (): void => {
    this.set({ dialog: null });
    const grid = this.state.grid;
    // An edit the dialog left open gets the focus back (and, on a phone, the keyboard).
    if (grid?.isEditing()) {
      if (grid.editState.origin === 'bar') this.root?.querySelector<HTMLElement>('[data-testid="formula-input"]')?.focus({ preventScroll: true });
      else grid.focusEdit();
    } else this.focusGrid();
  };

  openComments = (r: number, c: number): void => this.set({ comments: { r, c, rect: this.cellRect(r, c) } });

  closeComments = (): void => {
    this.set({ comments: null });
    this.focusGrid();
  };

  openNote = (r: number, c: number): void => this.set({ note: { r, c, rect: this.cellRect(r, c) } });

  closeNote = (): void => {
    this.set({ note: null });
    this.focusGrid();
  };

  openFilter = (col: number, at: { x: number; y: number }): void => this.set({ filter: { col, at } });

  closeFilter = (): void => {
    this.set({ filter: null });
    this.focusGrid();
  };

  openMenu = (menu: OpenMenu): void => this.set({ menu });

  /** Close the open menu (also the menu bar's). */
  closeMenu = (): void => {
    this.set({ menu: null, topMenu: null });
    this.focusGrid();
  };

  private showTopMenu(index: number, at: DOMRect, topMenu: number | null): void {
    const m = topMenus(this.ctx)[index];
    if (!m) return;
    this.set({ topMenu, menu: { items: m.items(), at, title: m.title, testId: `menu-${m.title.toLowerCase()}-list` } });
  }

  /** A menu bar title was clicked (`at`: its rectangle). */
  clickTopMenu = (index: number, at: DOMRect): void => {
    this.commitEdit();
    this.showTopMenu(index, at, this.state.topMenu === index ? null : index);
  };

  /** The pointer moved onto a menu bar title: another menu replaces the open one (not on phones). */
  hoverTopMenu = (index: number, at: DOMRect): void => {
    const { topMenu, phone } = this.state;
    if (topMenu === null || topMenu === index || phone) return;
    this.showTopMenu(index, at, index);
  };

  selectObject = (id: string | null): void => {
    this.state.grid?.selectImage(id);
    this.set({ selectedObject: id });
  };

  // ---------------------------------------------------------------------------------------------
  // Documents

  /** Settings of a workbook about to be shown: view-only and the author. */
  private prepare(doc: SpreadsheetDocument): SpreadsheetDocument {
    if (this.options.readOnly) doc.mode = 'viewOnly';
    doc.author = this.author();
    return doc;
  }

  private makeAssist(doc: SpreadsheetDocument | null): FormulaAssistState {
    const names = doc ? doc.wb.names.filter((n) => !n.hidden).map((n) => n.name) : [];
    return new FormulaAssistState(names, (text, caret) => this.state.grid?.replaceEditText(text, caret), () => {
      this.ctxCache = null;
      this.notify();
    });
  }

  /** Show a workbook in place of the current one (or of the start screen). */
  private load(next: SpreadsheetDocument, handle: FileSystemFileHandle | null = null): void {
    this.prepare(next);
    this.handle = handle;
    this.recovery = null;
    this.unwireDocument();
    // The old grid goes with its workbook; the view attaches a new one for this workbook.
    const prev = this.state.grid;
    if (prev) {
      prev.destroy();
      this.gridHost = null;
    }
    this.assist = this.makeAssist(next);
    this.set({ ...PER_DOCUMENT, doc: next, docKey: this.state.docKey + 1, docVersion: 0, restore: null });
    this.wireDocument(next);
    this.emit('open', { fileName: next.fileName });
  }

  private wireDocument(doc: SpreadsheetDocument): void {
    const bump = () => this.set({ docVersion: this.state.docVersion + 1 });
    this.docOffs.push(...DOC_STATE_EVENTS.map((e) => doc.on(e as 'change', bump)));
    for (const name of ['change', 'selection', 'activeSheet', 'sheets', 'file'] as const) {
      this.docOffs.push(doc.on(name as 'change', (e: unknown) => this.emit(name, e)));
    }
    this.autosaveOff = false;
    this.docOffs.push(doc.on('change', () => this.scheduleAutosave(doc)));
  }

  private unwireDocument(): void {
    for (const off of this.docOffs.splice(0)) off();
    if (this.autosaveTimer) clearTimeout(this.autosaveTimer);
    this.autosaveTimer = null;
  }

  private scheduleAutosave(doc: SpreadsheetDocument): void {
    const db = this.options.autosave;
    if (!doc.modified || !db || this.autosaveOff) return;
    if (this.autosaveTimer) clearTimeout(this.autosaveTimer);
    this.autosaveTimer = setTimeout(() => {
      if (cellCount(doc) > AUTOSAVE_MAX_CELLS) {
        this.autosaveOff = true;
        return;
      }
      try {
        const started = performance.now();
        const bytes = doc.save(false);
        if (performance.now() - started > 1500) this.autosaveOff = true;
        void saveAutosave(db, { fileName: doc.fileName, bytes, savedAt: Date.now() });
      } catch (e) {
        console.warn('Autosave failed', e);
      }
    }, AUTOSAVE_DELAY);
  }

  private clearAutosave(): void {
    if (this.options.autosave) void clearAutosave(this.options.autosave);
  }

  /** Opens a workbook from bytes (XLSX, XLSM, XLS, XLSB, CSV, TSV); resolves to whether it could be read. */
  openBytes(bytes: Uint8Array, name: string, handle: FileSystemFileHandle | null = null): Promise<boolean> {
    return new Promise<boolean>((resolve) => {
      this.set({ busy: `Opening “${name}”…` });
      // Importing is synchronous: let the message paint first.
      requestAnimationFrame(() =>
        setTimeout(() => {
          try {
            this.load(openDocument(bytes, name), handle);
            this.clearAutosave();
            resolve(true);
          } catch (e) {
            console.error(e);
            this.notice(`Could not open “${name}”: ${(e as Error)?.message ?? e}`);
            resolve(false);
          } finally {
            this.set({ busy: null });
          }
        }, 0),
      );
    });
  }

  openFile(file: File, handle: FileSystemFileHandle | null = null): Promise<boolean> {
    return file.arrayBuffer().then((buf) => this.openBytes(new Uint8Array(buf), file.name, handle));
  }

  async openUrl(url: string, name: string): Promise<boolean> {
    try {
      const r = await fetch(url);
      if (!r.ok) throw new Error(`HTTP ${r.status}`);
      return await this.openBytes(new Uint8Array(await r.arrayBuffer()), name);
    } catch (e) {
      this.notice(`Could not load ${url}: ${(e as Error).message}`);
      return false;
    }
  }

  /** Open a workbook: a URL, a file handle, a File or Blob, or the bytes; the file picker when no source is given. */
  open = async (source?: DocumentSource, name?: string): Promise<boolean> => {
    if (source === undefined) {
      this.files.open();
      return false;
    }
    if (typeof source === 'string') return this.openUrl(source, name ?? nameOfUrl(source));
    let bytes: Uint8Array;
    try {
      if (isFileHandle(source)) return await this.openFile(await source.getFile(), source);
      bytes = source instanceof Uint8Array ? source : source instanceof ArrayBuffer ? new Uint8Array(source) : new Uint8Array(await source.arrayBuffer());
    } catch (e) {
      this.notice(`Could not open the file: ${(e as Error).message}`);
      return false;
    }
    return this.openBytes(bytes, name ?? (typeof File !== 'undefined' && source instanceof File ? source.name : 'Workbook.xlsx'));
  };

  private async openDocument(input: DocumentInput): Promise<boolean> {
    const { source, name } = documentOf(input);
    return source === undefined ? false : this.open(source, name);
  }

  /** The start screen's "Create new spreadsheet". */
  startNew = (): void => {
    if (this.options.files?.newWorkbook) this.options.files.newWorkbook();
    else this.load(newDocument());
  };

  /** Recover the autosaved workbook offered. */
  recover = (): void => {
    const r = this.recovery;
    if (!r) return;
    try {
      const recovered = openDocument(r.bytes, r.fileName);
      recovered.modified = true;
      this.load(recovered);
    } catch (e) {
      this.notice(`Could not recover “${r.fileName}”: ${(e as Error).message}`);
    }
  };

  discardRecovery = (): void => {
    this.recovery = null;
    this.set({ restore: null });
    this.clearAutosave();
  };

  private async offerRecovery(): Promise<void> {
    const db = this.options.autosave;
    if (!db || this.options.offerRecovery === false) return;
    const rec = await loadAutosave(db);
    if (this.state.doc || !rec || Date.now() - rec.savedAt >= RECOVERY_MS) return;
    this.recovery = rec;
    this.set({ restore: { fileName: rec.fileName, savedAt: rec.savedAt } });
  }

  /** Ask before replacing a workbook with unsaved changes. */
  confirmDiscard = (): boolean => !this.state.doc?.modified || window.confirm(DISCARD_PROMPT);

  // ---------------------------------------------------------------------------------------------
  // Files

  private makeFiles(): FileActions {
    const pick = <K extends Exclude<keyof FileActions, 'share'>>(k: K): FileActions[K] => this.options.files?.[k] ?? this.builtinFiles[k];
    const files: FileActions = {
      newWorkbook: () => pick('newWorkbook')(),
      open: () => pick('open')(),
      save: () => pick('save')(),
      download: (format) => pick('download')(format),
      print: () => pick('print')(),
      rename: (name) => pick('rename')(name),
    };
    // Only where the browser can share files (or the share action was replaced).
    Object.defineProperty(files, 'share', { enumerable: true, get: () => this.options.files?.share ?? this.builtinFiles.share });
    return files;
  }

  private commitEdit(): void {
    const grid = this.state.grid;
    if (grid?.isEditing()) grid.commitEdit();
  }

  private readonly builtinFiles: FileActions = (() => {
    const canShare = typeof navigator !== 'undefined' && !!navigator.canShare && (() => {
      try {
        return navigator.canShare({ files: [new File([new Uint8Array(1)], 'x.xlsx', { type: XLSX_MIME })] });
      } catch {
        return false;
      }
    })();
    const files: FileActions = {
      newWorkbook: () => {
        if (!this.confirmDiscard()) return;
        this.load(newDocument());
        this.clearAutosave();
      },
      open: async () => {
        if (!this.confirmDiscard()) return;
        const w = window as unknown as FilePickerWindow;
        if (w.showOpenFilePicker && window.isSecureContext) {
          try {
            const [handle] = await w.showOpenFilePicker({
              types: [{ description: 'Spreadsheets', accept: { [XLSX_MIME]: ['.xlsx', '.xlsm', '.xlsb', '.xls'], 'text/csv': ['.csv', '.tsv', '.txt'] } }],
              multiple: false,
            });
            void this.openFile(await handle.getFile(), handle);
          } catch (e) {
            if ((e as DOMException)?.name !== 'AbortError') this.notice(`Could not open the file: ${(e as Error).message}`);
          }
          return;
        }
        this.input?.click();
      },
      save: async () => {
        const d = this.state.doc;
        if (!d) return;
        this.commitEdit();
        const handle = this.handle;
        if (handle && /\.(xlsx|xlsm)$/i.test(handle.name)) {
          try {
            const writable = await (handle as unknown as { createWritable(): Promise<FileSystemWritableFileStream> }).createWritable();
            await writable.write(xlsxBlob(d, /\.xlsm$/i.test(handle.name)));
            await writable.close();
            this.clearAutosave();
            return;
          } catch (e) {
            if ((e as DOMException)?.name === 'AbortError') return;
            console.warn('Saving to the file failed, downloading instead', e);
          }
        }
        this.files.download('xlsx');
      },
      download: (format: DownloadFormat) => {
        const d = this.state.doc;
        if (!d) return;
        this.commitEdit();
        if (format === 'pdf') {
          this.set({ dialog: { kind: 'print' } });
          return;
        }
        if (format === 'csv' || format === 'tsv') {
          const delimiter = format === 'csv' ? ',' : '\t';
          downloadBlob(new Blob([`﻿${sheetToCsv(d, d.sheet, delimiter)}`], { type: format === 'csv' ? 'text/csv' : 'text/tab-separated-values' }), `${baseName(d)}.${format}`);
          return;
        }
        const macros = isMacro(d);
        downloadBlob(xlsxBlob(d, macros), `${baseName(d)}.${macros ? 'xlsm' : 'xlsx'}`);
        this.clearAutosave();
      },
      print: () => {
        if (!this.state.doc) return;
        this.commitEdit();
        this.set({ dialog: { kind: 'print' } });
      },
      rename: (name: string) => {
        const d = this.state.doc;
        if (!d) return;
        const ext = /\.(xlsx|xlsm|csv|tsv)$/i.exec(d.fileName)?.[0] ?? '.xlsx';
        d.setFileName(`${name.replace(/\.(xlsx|xlsm|xls|xlsb|csv|tsv)$/i, '')}${ext}`);
        this.handle = null;
      },
      share: canShare && typeof window !== 'undefined' && window.isSecureContext
        ? async () => {
            const d = this.state.doc;
            if (!d) return;
            this.commitEdit();
            try {
              const macros = isMacro(d);
              const name = `${baseName(d)}.${macros ? 'xlsm' : 'xlsx'}`;
              await navigator.share({ files: [new File([xlsxBlob(d, macros, false)], name, { type: macros ? XLSM_MIME : XLSX_MIME })], title: name });
            } catch (e) {
              if ((e as DOMException)?.name !== 'AbortError') this.notice(`Sharing failed: ${(e as Error).message}`);
            }
          }
        : undefined,
    };
    return files;
  })();

  // ---------------------------------------------------------------------------------------------
  // Page API events

  /** Listen to events of the open workbook (they follow the workbooks opened later). */
  on<K extends keyof ApiEvents>(event: K, fn: (e: ApiEvents[K]) => void): () => void {
    let set = this.apiListeners.get(event);
    if (!set) this.apiListeners.set(event, (set = new Set()));
    const f = fn as (e: unknown) => void;
    set.add(f);
    return () => set.delete(f);
  }

  private emit(event: keyof ApiEvents, e: unknown): void {
    for (const fn of [...(this.apiListeners.get(event) ?? [])]) fn(e);
  }

  /** Resolves once the grid of the open workbook is attached. */
  whenGrid(): Promise<void> {
    if (this.state.grid) return Promise.resolve();
    return new Promise((resolve) => this.gridWaiters.push(resolve));
  }

  // ---------------------------------------------------------------------------------------------
  // Options

  /** New options from the component's props; the document and the mode follow them. */
  setOptions(next: SumlineOptions): void {
    const prev = this.options;
    this.options = { ...next };
    const doc = this.state.doc;
    if (doc && next.author !== prev.author) doc.author = this.author();
    if (doc && !!next.readOnly !== !!prev.readOnly) doc.setMode(next.readOnly ? 'viewOnly' : 'editing');
    if (this.started && next.document && !sameDocument(next.document, prev.document)) void this.openDocument(next.document);
  }

  // ---------------------------------------------------------------------------------------------
  // Lifetime

  /** Start the interface in its root element: the layer, listeners, and the initial document. */
  attach(root: HTMLElement): void {
    if (this.root) this.detach();
    this.root = root;
    watchFonts();
    const layer = this.layer;
    layer.className = 'sl-layer';
    document.body.appendChild(layer);
    const input = document.createElement('input');
    input.type = 'file';
    input.accept = ACCEPT;
    input.hidden = true;
    input.dataset.testid = 'open-input';
    input.addEventListener('change', () => {
      const file = input.files?.[0];
      input.value = '';
      if (file) void this.openFile(file);
    });
    layer.appendChild(input);
    this.input = input;
    this.ctxCache = null;

    // Spreadsheets dropped on the editor open; images dropped on a workbook are inserted.
    const hasFiles = (e: DragEvent) => Array.from(e.dataTransfer?.types ?? []).includes('Files');
    let depth = 0;
    const enter = (e: DragEvent) => {
      if (!hasFiles(e) || !this.state.doc) return;
      depth++;
      this.set({ dragging: true });
    };
    const leave = () => {
      depth = Math.max(0, depth - 1);
      if (!depth) this.set({ dragging: false });
    };
    const over = (e: DragEvent) => {
      if (hasFiles(e)) e.preventDefault();
    };
    const drop = (e: DragEvent) => {
      depth = 0;
      this.set({ dragging: false });
      // The start screen's drop zone opens what is dropped on it itself (React's handlers run after this one).
      if ((e.target as Element | null)?.closest?.('.sl-dropzone')) return;
      const all = Array.from(e.dataTransfer?.files ?? []);
      const file = all.find(isSheetFile);
      if (!file) {
        const img = all.find((f) => f.type.startsWith('image/'));
        const doc = this.state.doc;
        if (img && doc && !doc.isReadOnly()) {
          e.preventDefault();
          void this.state.grid?.insertImageFile(img, img.name);
        }
        return;
      }
      e.preventDefault();
      if (!this.confirmDiscard()) return;
      void this.openFile(file);
    };
    for (const el of [root, layer]) {
      el.addEventListener('dragenter', enter);
      el.addEventListener('dragleave', leave);
      el.addEventListener('dragover', over);
      el.addEventListener('drop', drop);
      this.offs.push(() => {
        el.removeEventListener('dragenter', enter);
        el.removeEventListener('dragleave', leave);
        el.removeEventListener('dragover', over);
        el.removeEventListener('drop', drop);
      });
    }

    // Volatile functions (NOW, TODAY, RAND) refresh every minute while the page is visible.
    const volatile = setInterval(() => {
      if (document.visibilityState === 'visible' && !this.state.grid?.isEditing()) this.state.doc?.recalcVolatile();
    }, 60_000);
    this.offs.push(() => clearInterval(volatile), watchMedia(PHONE_QUERY, (phone) => this.set({ phone })));
    this.set({ phone: matchesMedia(PHONE_QUERY) });
    this.notify();
    if (!this.started) {
      this.started = true;
      void this.start();
    }
  }

  private async start(): Promise<void> {
    // Let a view that is mounted twice in a row (React's StrictMode) settle first.
    await Promise.resolve();
    const o = this.options;
    if (o.document) await this.openDocument(o.document);
    else if (!this.state.doc) void this.offerRecovery();
    await this.whenGrid();
    this.set({ ready: true });
    this.options.onReady?.(this.api);
  }

  /** The grid of the open workbook in its element (inside the root). */
  attachGrid(host: HTMLElement): void {
    if (this.gridHost === host && this.state.grid) return;
    const doc = this.state.doc;
    if (!doc) return;
    this.state.grid?.destroy();
    const g = new GridView(host, doc, {
      onEdit: (e) => {
        this.assist.update(e);
        this.set({ edit: e });
      },
      editorKey: (e) => this.assist.handleKey(e),
      onContextMenu: (x, y, target) => {
        const ctx = this.ctx;
        if (target === 'image') {
          const id = ctx.grid?.selectedImageId;
          if (id) this.openMenu({ items: objectMenu(ctx, id), at: { x, y }, testId: 'object-menu' });
          return;
        }
        const items = target === 'col' ? columnMenu(ctx) : target === 'row' || target === 'corner' ? rowMenu(ctx) : cellMenu(ctx);
        this.openMenu({ items, at: { x, y }, title: ctx.state.rangeRef, testId: `${target === 'col' ? 'column' : target === 'row' ? 'row' : 'cell'}-menu` });
      },
      onCommand: (c) => {
        const ctx = this.ctx;
        switch (c) {
          case 'find':
          case 'replace':
            this.set({ dialog: { kind: 'search', replace: true } });
            break;
          case 'link':
            if (!doc.isReadOnly()) this.set({ dialog: { kind: 'link' } });
            break;
          case 'comment':
            cmd.addComment(ctx);
            break;
          case 'save':
            ctx.files.save();
            break;
          case 'open':
            ctx.files.open();
            break;
          case 'print':
            ctx.files.print();
            break;
          case 'clear-format':
            if (!doc.isReadOnly()) cmd.removeFormat(ctx);
            break;
          case 'next-sheet':
          case 'prev-sheet': {
            const visible = doc.wb.sheets.map((s, i) => (s.state === 'visible' ? i : -1)).filter((i) => i >= 0);
            const pos = visible.indexOf(doc.activeIndex);
            doc.setActiveSheet(visible[(pos + (c === 'next-sheet' ? 1 : -1) + visible.length) % visible.length]);
            break;
          }
          default:
            break;
        }
      },
      onImageSelect: (id) => this.set({ selectedObject: id }),
      onObjectEdit: (id) => {
        const hit = doc.objectById(id);
        if (hit?.kind === 'chart') this.set({ dialog: { kind: 'chart-settings', id } });
        else if (hit?.kind === 'shape') this.set({ dialog: { kind: 'shape-settings', id } });
      },
      onLinkTap: (link) => {
        if (link.location) {
          doc.goTo(link.location.replace(/^#/, ''));
          return;
        }
        if (link.url) {
          const w = window.open(link.url, '_blank', 'noopener,noreferrer');
          if (!w) this.notify_(`Open ${link.url}`);
        }
      },
      onValidationList: (options, rect, pick) => {
        this.openMenu({ items: options.map((o) => ({ title: o, onClick: () => pick(o) })), at: rect, testId: 'validation-list' });
      },
      validateEdit: (r, c, text, sheet, decide) => {
        const rule = doc.invalidInput(r, c, text, sheet);
        if (!rule || rule.formatStyle || rule.showErrorMessage === false) return false;
        const message = rule.error || 'The content of the cell is not valid.';
        if (rule.errorStyle === 'stop') {
          void this.confirm(message, { title: rule.errorTitle || 'Invalid entry', ok: 'Retry', cancel: 'Cancel' }).then((retry) => {
            if (retry) decide(false);
            else g.cancelEdit();
          });
        } else {
          void this.confirm(`${message}\n\nKeep this value?`, { title: rule.errorTitle || (rule.errorStyle === 'warning' ? 'Warning' : 'Information'), ok: 'Yes', cancel: 'No' }).then((keep) => decide(keep));
        }
        return true;
      },
      onFilterButton: (col, x, y) => this.set({ filter: { col, at: { x, y } } }),
      onPaint: (vt) => {
        for (const fn of [...this.paintListeners]) fn(vt);
      },
    }, { mobile: isTouchDevice() });
    this.gridHost = host;
    this.set({ grid: g });
    if (!isTouchDevice()) g.focusIdle();
    for (const resolve of this.gridWaiters.splice(0)) resolve();
  }

  /** The grid's element goes away. */
  detachGrid(host: HTMLElement): void {
    if (this.gridHost !== host) return;
    this.state.grid?.destroy();
    this.gridHost = null;
    this.set({ grid: null });
  }

  /** Remove the interface and stop listening; attach() may be called again. */
  detach(): void {
    for (const off of this.offs.splice(0)) off();
    if (this.toastTimer) clearTimeout(this.toastTimer);
    this.toastTimer = null;
    this.input?.remove();
    this.input = null;
    this.layer.remove();
    this.root = null;
  }

  /** Remove the interface for good (the component unmounted): also the grid and the document's listeners. */
  destroy(): void {
    this.detach();
    if (this.gridHost) this.detachGrid(this.gridHost);
    this.unwireDocument();
  }
}
