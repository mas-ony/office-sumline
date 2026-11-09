/** What the interface's controls work with: the document, the grid, dialogs, notices and file actions. */
import type { SpreadsheetDocument } from '../core/document';
import type { EditState, GridView } from '../grid/GridView';
import type { DocState } from './docState';
import type { SumlineUI } from './controller';

export type DownloadFormat = 'xlsx' | 'csv' | 'tsv' | 'pdf';

export interface FileActions {
  newWorkbook(): void;
  open(): void;
  /** Ctrl+S / toolbar save: writes the workbook (to the opened file where possible). */
  save(): void;
  download(format: DownloadFormat): void;
  print(): void;
  rename(name: string): void;
  share?: () => void;
}

export type DialogSpec =
  | { kind: 'format'; tab?: 'mask' | 'editor' }
  | { kind: 'search'; replace?: boolean }
  | { kind: 'chart-insert' }
  | { kind: 'chart-settings'; id: string }
  | { kind: 'validations' }
  | { kind: 'shapes' }
  | { kind: 'shape-settings'; id: string }
  | { kind: 'print' }
  | { kind: 'name-edit'; name?: string }
  | { kind: 'link' }
  | { kind: 'about' }
  | { kind: 'shortcuts' }
  | { kind: 'functions' }
  | { kind: 'column-title'; col: number };

export interface ConfirmOptions {
  title?: string;
  ok?: string;
  cancel?: string;
  danger?: boolean;
}

/** Values the interface remembers (recent colours, print options): in localStorage with the storageKey option. */
export interface Prefs {
  get<T>(name: string, initial: T): T;
  set(name: string, value: unknown): void;
}

export interface UIContextValue {
  /** The controller of this editor. */
  ui: SumlineUI;
  /** Container on the page's body for menus, dialogs and pop-ups. */
  layer: HTMLElement;
  prefs: Prefs;
  doc: SpreadsheetDocument;
  grid: GridView | null;
  state: DocState;
  /** The cell edit in progress (mirrored by the formula bar). */
  edit: EditState;
  phone: boolean;
  files: FileActions;
  dialog: DialogSpec | null;
  openDialog(d: DialogSpec): void;
  closeDialog(): void;
  notify(message: string): void;
  confirm(message: string, opts?: ConfirmOptions): Promise<boolean>;
  prompt(message: string, value?: string, opts?: ConfirmOptions): Promise<string | null>;
  /** Comment thread box at a cell (threaded comments). */
  openComments(r: number, c: number): void;
  /** Note editor at a cell (legacy notes). */
  openNote(r: number, c: number): void;
  /** Filter pop-up of a column of the auto filter. */
  openFilter(col: number, at: { x: number; y: number }): void;
  /** Chart, shape or picture currently selected. */
  selectedObject: string | null;
  selectObject(id: string | null): void;
}
