/**
 * The interactive grid: a canvas drawn over a natively scrolling layer (momentum scrolling on
 * touch devices), with mouse, touch, pen and keyboard interaction, an in-cell editor (a real
 * textarea, so mobile keyboards, autocorrect and IME work) and image objects.
 */
import { MAX_COLS, MAX_ROWS, cellName, makeRange, rangeName, rangesIntersect, type Range } from '../core/refs';
import type { Sheet, SheetImage } from '../core/model/sheet';
import type { SpreadsheetDocument } from '../core/document';
import { SheetLayout } from '../core/layout';
import { tokenize } from '../core/formula/tokenizer';
import { quoteSheetName } from '../core/refs';
import { formatValue } from '../core/format/display';
import { cssFont, lineHeightPx, textWidth } from './fonts';
import { ViewTransform, type Rect, type ViewState } from './viewport';
import { clearTextCache, fillColor, filterButtonRect, headerHeightFor, headerWidthFor, render, type Highlight, type RenderState } from './render';
import { bumpConditionalFormats } from './condFormat';
import { resolveColor } from '../core/model/color';
import { chartPlaceholders } from './charts';
import { chartSvg } from './chartRender';
import { shapeSvg } from './shapes';
import { readChartData } from '../core/chartData';

export const REF_COLORS = ['#1a73e8', '#d93025', '#8430ce', '#188038', '#e37400', '#c5221f', '#129eaf', '#b31412'];

export interface EditState {
  active: boolean;
  text: string;
  /** 'enter': started by typing (arrow keys commit); 'edit': F2 / double click (arrows move the caret). */
  mode: 'enter' | 'edit';
  origin: 'cell' | 'bar';
  caret: number;
  r: number;
  c: number;
  sheet?: Sheet;
}

export type GridCommand =
  | 'find'
  | 'replace'
  | 'link'
  | 'comment'
  | 'save'
  | 'open'
  | 'print'
  | 'next-sheet'
  | 'prev-sheet'
  | 'context-menu'
  | 'clear-format';

export interface GridCallbacks {
  /** Context menu at client coordinates. */
  onContextMenu?(x: number, y: number, target: 'cell' | 'row' | 'col' | 'corner' | 'image'): void;
  onEdit?(state: EditState): void;
  onCommand?(cmd: GridCommand): void;
  onZoom?(zoom: number): void;
  /** A picture, chart or shape was selected (null: none). */
  onImageSelect?(id: string | null): void;
  /** Double click / double tap on a chart or shape. */
  onObjectEdit?(id: string): void;
  onLinkTap?(link: { url?: string; location?: string; x: number; y: number }): void;
  /** Lets the UI consume keys while the formula autocomplete is open. */
  editorKey?(e: KeyboardEvent): boolean;
  onValidationList?(options: string[], rect: DOMRect, pick: (v: string) => void): void;
  /**
   * Called before an edit is written. Return true to take over: the edit stays open until
   * `decide(true)` writes it or `decide(false)` returns to editing.
   */
  validateEdit?(r: number, c: number, text: string, sheet: Sheet, decide: (accept: boolean) => void): boolean;
  /** The drop-down button of an auto filter column was pressed. */
  onFilterButton?(col: number, x: number, y: number): void;
  /** After every repaint (the footer row follows the columns). */
  onPaint?(vt: ViewTransform): void;
}

type Drag =
  | { kind: 'select'; startR: number; startC: number; pointerId: number; additive: boolean }
  | { kind: 'header'; axis: 'row' | 'col'; start: number; pointerId: number; additive: boolean }
  | { kind: 'resize'; axis: 'row' | 'col'; index: number; startPos: number; startSize: number; pointerId: number; indexes: number[] }
  | { kind: 'fill'; pointerId: number }
  | { kind: 'handle'; corner: 'start' | 'end'; pointerId: number }
  | { kind: 'image'; id: string; objKind: 'image' | 'chart' | 'shape'; mode: 'move' | string; startX: number; startY: number; rect: Rect; pointerId: number; moved: boolean }
  | { kind: 'point'; pointerId: number; startR: number; startC: number }
  | { kind: 'touch'; pointerId: number; x: number; y: number; time: number; r: number; c: number; longPress: number | null; moved: boolean; selecting: boolean };

const isTouchPointer = (e: PointerEvent) => e.pointerType === 'touch' || e.pointerType === 'pen';

/** Two taps on the same cell within this time are a double tap (edit). */
const DOUBLE_TAP_MS = 320;

export class GridView {
  readonly root: HTMLElement;
  private scroller: HTMLDivElement;
  private spacer: HTMLDivElement;
  private canvas: HTMLCanvasElement;
  private ctx: CanvasRenderingContext2D;
  private layer: HTMLDivElement;
  private mainImages: HTMLDivElement;
  private frozenImages: HTMLDivElement;
  readonly editor: HTMLTextAreaElement;
  private handles: HTMLDivElement[] = [];
  private dvButton: HTMLButtonElement;
  private tooltip: HTMLDivElement;
  private frame = 0;
  private disposers: (() => void)[] = [];
  private drag: Drag | null = null;
  private extentRows = 200;
  private extentCols = 40;
  private width = 0;
  private height = 0;
  private dpr = 1;
  private edit: EditState = { active: false, text: '', mode: 'enter', origin: 'cell', caret: 0, r: 0, c: 0 };
  /** Range inserted by the last click in formula point mode (replaced by the next click). */
  private pointRef: { start: number; end: number } | null = null;
  private fillPreview: Range | undefined;
  private resizeGuide: RenderState['resizeGuide'];
  private selectedImage: string | null = null;
  private imageEls = new Map<string, HTMLDivElement>();
  /** Last markup drawn into each chart or shape element (redrawn only when it changes). */
  private objectMarkup = new Map<string, string>();
  private lastObjectTap = { id: '', time: 0 };
  /** Start of a run of Tab entries (Enter returns to its column). */
  private tabStart: { r: number; c: number } | null = null;
  private pointers = new Map<number, { x: number; y: number }>();
  private pinch: { dist: number; zoom: number; cx: number; cy: number } | null = null;
  private lastTap = { time: 0, r: -1, c: -1 };
  private swallowTouchEnd = false;
  /** Menu or link popup waiting for the double-tap window to pass. */
  private pendingTap = 0;
  private autoScroll: { dx: number; dy: number; x: number; y: number } | null = null;
  private autoScrollFrame = 0;
  private scrollMemo = new WeakMap<Sheet, { x: number; y: number }>();
  private currentSheet: Sheet;
  searchHit?: { r: number; c: number };
  extraHighlights: Highlight[] = [];
  /** True on phones and tablets (no hover, coarse pointer). */
  mobile: boolean;

  constructor(root: HTMLElement, readonly doc: SpreadsheetDocument, readonly cb: GridCallbacks = {}, opts: { mobile?: boolean } = {}) {
    this.root = root;
    this.mobile = !!opts.mobile;
    root.classList.add('xgrid');
    root.innerHTML = '';
    this.scroller = document.createElement('div');
    this.scroller.className = 'xgrid-scroller';
    this.scroller.tabIndex = -1;
    this.spacer = document.createElement('div');
    this.spacer.className = 'xgrid-spacer';
    this.scroller.appendChild(this.spacer);
    this.canvas = document.createElement('canvas');
    this.canvas.className = 'xgrid-canvas';
    this.canvas.setAttribute('aria-hidden', 'true');
    this.ctx = this.canvas.getContext('2d', { alpha: false })!;
    this.layer = document.createElement('div');
    this.layer.className = 'xgrid-layer';
    this.frozenImages = document.createElement('div');
    this.frozenImages.className = 'xgrid-images';
    this.mainImages = document.createElement('div');
    this.mainImages.className = 'xgrid-images';
    this.editor = document.createElement('textarea');
    this.editor.className = 'xgrid-editor';
    this.editor.setAttribute('aria-label', 'Cell editor');
    this.editor.setAttribute('autocapitalize', 'off');
    this.editor.setAttribute('autocomplete', 'off');
    this.editor.setAttribute('spellcheck', 'false');
    this.editor.setAttribute('data-testid', 'cell-editor');
    this.editor.rows = 1;
    this.dvButton = document.createElement('button');
    this.dvButton.type = 'button';
    this.dvButton.className = 'xgrid-dv';
    this.dvButton.setAttribute('aria-label', 'Show list');
    this.dvButton.innerHTML = '<svg viewBox="0 0 10 10" width="10" height="10"><path d="M1 3l4 4 4-4" fill="none" stroke="currentColor" stroke-width="1.5"/></svg>';
    this.tooltip = document.createElement('div');
    this.tooltip.className = 'xgrid-tooltip';
    for (const corner of ['start', 'end'] as const) {
      const h = document.createElement('div');
      h.className = `xgrid-handle xgrid-handle-${corner}`;
      h.dataset.corner = corner;
      this.handles.push(h);
    }
    this.layer.append(this.frozenImages, this.mainImages, ...this.handles, this.dvButton, this.editor, this.tooltip);
    root.append(this.scroller, this.canvas, this.layer);
    this.currentSheet = doc.sheet;
    // Text measurement for auto-fit (row heights grow with wrapped or larger text).
    doc.measure ??= (text, style) => ({ width: textWidth(text, cssFont(style.font)), height: lineHeightPx(style.font.size) });
    this.bind();
    this.resize();
    this.syncImages();
    this.syncEditorIdle();
    this.requestRender();
  }

  // ---------------------------------------------------------------------------------------------
  // Setup

  private bind(): void {
    const on = <K extends keyof HTMLElementEventMap>(el: HTMLElement | Window | Document, type: K, fn: (e: HTMLElementEventMap[K]) => void, opts?: AddEventListenerOptions) => {
      el.addEventListener(type, fn as EventListener, opts);
      this.disposers.push(() => el.removeEventListener(type, fn as EventListener, opts));
    };
    on(this.scroller, 'scroll', () => this.onScroll());
    on(this.scroller, 'pointerdown', (e) => this.onPointerDown(e));
    on(this.scroller, 'pointermove', (e) => this.onPointerMove(e));
    on(this.scroller, 'pointerup', (e) => this.onPointerUp(e));
    on(this.scroller, 'pointercancel', (e) => this.onPointerCancel(e));
    on(this.scroller, 'dblclick', (e) => this.onDoubleClick(e));
    on(this.scroller, 'contextmenu', (e) => this.onContextMenu(e));
    on(this.scroller, 'wheel', (e) => this.onWheel(e), { passive: false });
    on(this.scroller, 'touchmove', (e) => this.onTouchMove(e), { passive: false });
    on(this.scroller, 'touchend', (e) => {
      if (this.swallowTouchEnd && e.cancelable) e.preventDefault();
      this.swallowTouchEnd = false;
    }, { passive: false });
    on(this.scroller, 'mouseleave', () => this.hideTooltip());
    for (const h of this.handles) {
      on(h, 'pointerdown', (e) => this.onHandleDown(e, h.dataset.corner as 'start' | 'end'));
      on(h, 'pointermove', (e) => this.onPointerMove(e));
      on(h, 'pointerup', (e) => this.onPointerUp(e));
      on(h, 'pointercancel', (e) => this.onPointerCancel(e));
    }
    on(this.editor, 'keydown', (e) => this.onKeyDown(e));
    on(this.editor, 'input', () => this.onEditorInput());
    on(this.editor, 'compositionstart', () => this.onCompositionStart());
    on(this.editor, 'select', () => this.emitEdit());
    on(this.editor, 'keyup', () => this.edit.active && this.updateCaret());
    on(this.editor, 'click', () => this.edit.active && this.updateCaret());
    on(this.editor, 'copy', (e) => this.onCopy(e, false));
    on(this.editor, 'cut', (e) => this.onCopy(e, true));
    on(this.editor, 'paste', (e) => this.onPaste(e));
    on(this.dvButton, 'pointerdown', (e) => {
      e.preventDefault();
      e.stopPropagation();
      this.openValidationList();
    });
    const ro = new ResizeObserver(() => this.resize());
    ro.observe(this.root);
    this.disposers.push(() => ro.disconnect());
    const offs = [
      this.doc.on('change', (ev) => {
        clearTextCache();
        bumpConditionalFormats();
        if (ev.structure || ev.layout) this.updateExtent();
        this.syncImages();
        this.requestRender();
      }),
      this.doc.on('selection', () => {
        if (!this.drag) this.scrollIntoView(this.doc.selection.active.r, this.doc.selection.active.c);
        this.syncEditorIdle();
        this.requestRender();
      }),
      this.doc.on('activeSheet', () => this.onSheetChange()),
      this.doc.on('mode', () => {
        if (this.doc.isReadOnly() && this.edit.active) this.cancelEdit();
        this.requestRender();
      }),
    ];
    this.disposers.push(...offs);
    const mq = window.matchMedia?.('(prefers-color-scheme: dark)');
    void mq;
  }

  destroy(): void {
    this.cancelPendingTap();
    cancelAnimationFrame(this.frame);
    cancelAnimationFrame(this.autoScrollFrame);
    for (const d of this.disposers) d();
    this.root.innerHTML = '';
  }

  private get sheet(): Sheet {
    return this.doc.sheet;
  }

  private onSheetChange(): void {
    if (this.edit.active && !this.isFormulaEdit()) {
      // Switching sheets finishes a plain edit in place (formula edits continue for references).
      const { r, c, text, sheet } = this.edit;
      this.edit = { ...this.edit, active: false };
      this.pointRef = null;
      this.editor.classList.remove('editing');
      this.editor.value = '';
      if (sheet && text !== this.doc.editText(r, c, sheet)) this.doc.setInput(r, c, text, sheet);
      this.emitEdit();
    }
    this.scrollMemo.set(this.currentSheet, { x: this.scroller.scrollLeft, y: this.scroller.scrollTop });
    this.currentSheet = this.sheet;
    this.selectedImage = null;
    this.updateExtent();
    const memo = this.scrollMemo.get(this.sheet);
    this.scroller.scrollLeft = memo?.x ?? 0;
    this.scroller.scrollTop = memo?.y ?? 0;
    this.syncImages();
    this.syncEditorIdle();
    this.cb.onZoom?.(this.sheet.zoom);
    this.requestRender();
  }

  resize(): void {
    this.dpr = Math.min(3, window.devicePixelRatio || 1);
    this.width = this.scroller.clientWidth || this.root.clientWidth;
    this.height = this.scroller.clientHeight || this.root.clientHeight;
    this.canvas.style.width = `${this.width}px`;
    this.canvas.style.height = `${this.height}px`;
    this.canvas.width = Math.max(1, Math.floor(this.width * this.dpr));
    this.canvas.height = Math.max(1, Math.floor(this.height * this.dpr));
    this.updateExtent();
    // The on-screen keyboard shrinks the viewport: keep the cell being edited in view.
    if (this.edit.active && this.edit.sheet === this.sheet && this.width && this.height) this.scrollIntoView(this.edit.r, this.edit.c);
    this.requestRender();
  }

  get zoom(): number {
    return this.sheet.zoom;
  }

  setZoom(zoom: number, anchor?: { x: number; y: number }): void {
    const z = Math.max(0.25, Math.min(4, Math.round(zoom * 100) / 100));
    const old = this.sheet.zoom;
    if (z === old) return;
    const vt = this.transform();
    const ax = anchor?.x ?? vt.mainLeft;
    const ay = anchor?.y ?? vt.mainTop;
    const sx = vt.sheetX(ax);
    const sy = vt.sheetY(ay);
    this.doc.setZoom(z);
    this.updateExtent();
    const nvt = this.transform();
    const hw = nvt.v.headerW;
    const hh = nvt.v.headerH;
    this.scroller.scrollLeft = Math.max(0, sx * z - (ax - hw) - nvt.FW * z + nvt.FW * z);
    this.scroller.scrollTop = Math.max(0, sy * z - (ay - hh));
    this.cb.onZoom?.(z);
    this.requestRender();
  }

  /** Grows the scrollable area beyond the used range so the sheet feels endless. */
  private updateExtent(): void {
    const used = this.sheet.usedRange();
    const layout = SheetLayout.of(this.sheet);
    const z = this.zoom;
    this.extentRows = Math.min(MAX_ROWS, Math.max(this.extentRows, (used?.r1 ?? 0) + 60, this.sheet.freeze.rows + Math.ceil(this.height / (15 * z)) + 40));
    this.extentCols = Math.min(MAX_COLS, Math.max(this.extentCols, (used?.c1 ?? 0) + 12, this.sheet.freeze.cols + Math.ceil(this.width / (40 * z)) + 8));
    const v = this.viewState();
    const w = v.headerW + layout.colX(this.extentCols) * z;
    const h = v.headerH + layout.rowY(this.extentRows) * z;
    this.spacer.style.width = `${Math.min(w, 3e7)}px`;
    this.spacer.style.height = `${Math.min(h, 3e7)}px`;
  }

  private viewState(): ViewState {
    const z = this.zoom;
    const vtRows = Math.max(this.extentRows, 1000);
    return {
      width: this.width,
      height: this.height,
      zoom: z,
      scrollX: this.scroller.scrollLeft,
      scrollY: this.scroller.scrollTop,
      headerW: headerWidthFor(Math.min(vtRows, Math.floor((this.scroller.scrollTop + this.height) / (15 * z)) + 100), z),
      headerH: headerHeightFor(z),
    };
  }

  transform(): ViewTransform {
    return new ViewTransform(this.sheet, this.viewState());
  }

  private onScroll(): void {
    const z = this.zoom;
    const layout = SheetLayout.of(this.sheet);
    // Extend the sheet when scrolling near its end.
    if (this.scroller.scrollTop + this.height > layout.rowY(this.extentRows) * z - 400 && this.extentRows < MAX_ROWS) {
      this.extentRows = Math.min(MAX_ROWS, Math.ceil(this.extentRows * 1.5) + 100);
      this.updateExtent();
    }
    if (this.scroller.scrollLeft + this.width > layout.colX(this.extentCols) * z - 300 && this.extentCols < MAX_COLS) {
      this.extentCols = Math.min(MAX_COLS, Math.ceil(this.extentCols * 1.5) + 10);
      this.updateExtent();
    }
    this.hideTooltip();
    this.requestRender();
  }

  requestRender(): void {
    if (this.frame) return;
    this.frame = requestAnimationFrame(() => {
      this.frame = 0;
      this.paint();
    });
  }

  /** Synchronous repaint (tests, screenshots). */
  paint(): void {
    if (!this.width || !this.height) {
      this.resize();
      if (!this.width) return;
    }
    const vt = this.transform();
    const ctx = this.ctx;
    ctx.setTransform(this.dpr, 0, 0, this.dpr, 0, 0);
    const state: RenderState = {
      selection: this.doc.selection,
      editing: this.edit.active && this.edit.origin === 'cell' && !this.isFormulaEdit(),
      mobile: this.mobile,
      highlights: [...this.formulaHighlights(), ...this.extraHighlights],
      searchHit: this.searchHit,
      fillPreview: this.fillPreview,
      dpr: this.dpr,
      resizeGuide: this.resizeGuide,
    };
    render(ctx, this.doc, vt, state);
    this.positionImages(vt);
    this.positionHandles(vt);
    this.positionEditor(vt);
    this.positionValidationButton(vt);
    this.cb.onPaint?.(vt);
  }

  // ---------------------------------------------------------------------------------------------
  // Hit testing

  private local(e: { clientX: number; clientY: number }): { x: number; y: number } {
    const r = this.root.getBoundingClientRect();
    return { x: e.clientX - r.left, y: e.clientY - r.top };
  }

  private hit(x: number, y: number, touch = false): { area: 'corner' | 'col' | 'row' | 'cell' | 'none'; r: number; c: number; resize?: { axis: 'row' | 'col'; index: number } } {
    const vt = this.transform();
    const { headerW, headerH } = vt.v;
    if (x > this.width || y > this.height) return { area: 'none', r: 0, c: 0 };
    const tol = touch ? 10 : 4;
    if (y < headerH && x < headerW) return { area: 'corner', r: 0, c: 0 };
    if (y < headerH) {
      const c = vt.colAt(x);
      const rect = vt.cellRect(0, c);
      if (Math.abs(x - (rect.x + rect.w)) <= tol) return { area: 'col', r: 0, c, resize: { axis: 'col', index: c } };
      if (Math.abs(x - rect.x) <= tol && c > 0) return { area: 'col', r: 0, c, resize: { axis: 'col', index: this.prevVisibleCol(c - 1) } };
      return { area: 'col', r: 0, c };
    }
    if (x < headerW) {
      const r = vt.rowAt(y);
      const rect = vt.cellRect(r, 0);
      if (Math.abs(y - (rect.y + rect.h)) <= tol) return { area: 'row', r, c: 0, resize: { axis: 'row', index: r } };
      if (Math.abs(y - rect.y) <= tol && r > 0) return { area: 'row', r, c: 0, resize: { axis: 'row', index: this.prevVisibleRow(r - 1) } };
      return { area: 'row', r, c: 0 };
    }
    return { area: 'cell', r: vt.rowAt(y), c: vt.colAt(x) };
  }

  private prevVisibleCol(c: number): number {
    while (c > 0 && this.sheet.isColHidden(c)) c--;
    return c;
  }

  private prevVisibleRow(r: number): number {
    while (r > 0 && this.sheet.isRowHidden(r)) r--;
    return r;
  }

  private nearFillHandle(x: number, y: number): boolean {
    if (this.mobile || this.doc.isReadOnly() || this.edit.active) return false;
    const vt = this.transform();
    const pr = vt.rangeRect(this.doc.range);
    return Math.abs(x - (pr.x + pr.w)) <= 5 && Math.abs(y - (pr.y + pr.h)) <= 5;
  }

  // ---------------------------------------------------------------------------------------------
  // Pointer interaction

  private onPointerDown(e: PointerEvent): void {
    this.tabStart = null;
    this.cancelPendingTap();
    const p = this.local(e);
    this.pointers.set(e.pointerId, p);
    if (this.pointers.size === 2) {
      // Start pinch zoom; abandon any single pointer action.
      const [a, b] = [...this.pointers.values()];
      this.cancelTouchTimer();
      this.drag = null;
      this.pinch = { dist: Math.hypot(a.x - b.x, a.y - b.y), zoom: this.zoom, cx: (a.x + b.x) / 2, cy: (a.y + b.y) / 2 };
      return;
    }
    if (e.button === 2) return;
    const touch = isTouchPointer(e);
    const h = this.hit(p.x, p.y, touch);
    if (h.area === 'none') return;
    this.hideTooltip();
    // Pressing on the grid leaves the selected picture, chart or shape.
    if (this.selectedImage) this.selectImage(null);

    if (!touch) e.preventDefault();
    if (h.resize && !this.doc.isReadOnly()) {
      const axis = h.resize.axis;
      const sel = this.doc.range;
      const full = axis === 'col' ? sel.r0 === 0 && sel.r1 === MAX_ROWS - 1 : sel.c0 === 0 && sel.c1 === MAX_COLS - 1;
      const inSel = axis === 'col' ? h.resize.index >= sel.c0 && h.resize.index <= sel.c1 : h.resize.index >= sel.r0 && h.resize.index <= sel.r1;
      const indexes = full && inSel ? range(axis === 'col' ? sel.c0 : sel.r0, axis === 'col' ? sel.c1 : sel.r1) : [h.resize.index];
      const size = axis === 'col' ? this.sheet.colWidthPx(h.resize.index) : this.sheet.rowHeightPx(h.resize.index);
      this.drag = { kind: 'resize', axis, index: h.resize.index, startPos: axis === 'col' ? p.x : p.y, startSize: size, pointerId: e.pointerId, indexes };
      this.capture(e);
      return;
    }

    // Auto filter drop-down buttons in the header row.
    const af = this.sheet.autoFilter;
    if (af && h.area === 'cell' && h.r === af.ref.r0 && h.c >= af.ref.c0 && h.c <= af.ref.c1) {
      const b = filterButtonRect(this.transform().cellRect(h.r, h.c), this.zoom);
      const pad = touch ? 8 : 1;
      if (p.x >= b.x - pad && p.x <= b.x + b.w + pad && p.y >= b.y - pad && p.y <= b.y + b.h + pad) {
        e.preventDefault();
        if (this.edit.active) this.commitEdit();
        this.doc.selectCell(h.r, h.c);
        const root = this.root.getBoundingClientRect();
        this.cb.onFilterButton?.(h.c, root.left + b.x, root.top + b.y + b.h + 2);
        if (touch) this.swallowTouchEnd = true;
        return;
      }
    }

    if (touch && h.area === 'cell') {
      // Taps select; long presses select ranges or open the menu; moves scroll natively.
      const longPress = window.setTimeout(() => this.onLongPress(), 450);
      this.drag = { kind: 'touch', pointerId: e.pointerId, x: p.x, y: p.y, time: Date.now(), r: h.r, c: h.c, longPress, moved: false, selecting: false };
      return;
    }

    // Formula point mode: clicking cells inserts references.
    if (h.area === 'cell' && this.edit.active && this.isFormulaEdit() && this.canInsertRef()) {
      this.insertPointRef(makeRange(h.r, h.c));
      this.drag = { kind: 'point', pointerId: e.pointerId, startR: h.r, startC: h.c };
      this.capture(e);
      return;
    }
    if (this.edit.active) this.commitEdit();

    if (h.area === 'cell' && this.nearFillHandle(p.x, p.y)) {
      this.drag = { kind: 'fill', pointerId: e.pointerId };
      this.capture(e);
      return;
    }

    if (h.area === 'corner') {
      this.doc.setSelection({ active: { r: 0, c: 0 }, anchor: { r: 0, c: 0 }, ranges: [makeRange(0, 0, MAX_ROWS - 1, MAX_COLS - 1)] });
      this.focusIdle();
      return;
    }
    if (h.area === 'col' || h.area === 'row') {
      const additive = e.ctrlKey || e.metaKey;
      const idx = h.area === 'col' ? h.c : h.r;
      if (e.shiftKey) {
        const sel = this.doc.selection;
        if (h.area === 'col') this.doc.selectCols(sel.anchor.c, idx);
        else this.doc.selectRows(sel.anchor.r, idx);
      } else if (h.area === 'col') this.doc.selectCols(idx, idx, additive);
      else this.doc.selectRows(idx, idx, additive);
      this.drag = { kind: 'header', axis: h.area, start: e.shiftKey ? (h.area === 'col' ? this.doc.selection.anchor.c : this.doc.selection.anchor.r) : idx, pointerId: e.pointerId, additive };
      this.capture(e);
      this.focusIdle();
      return;
    }

    // Links open with Ctrl+click (or a plain click when viewing).
    const link = this.sheet.hyperlinkAt(h.r, h.c);
    if (link && (e.ctrlKey || e.metaKey)) {
      // Ctrl+click follows the link (which may switch sheets) instead of selecting.
      this.cb.onLinkTap?.({ url: link.target, location: link.location, x: e.clientX, y: e.clientY });
      this.focusIdle();
      return;
    }
    if (link && this.doc.isReadOnly()) this.cb.onLinkTap?.({ url: link.target, location: link.location, x: e.clientX, y: e.clientY });

    if (e.shiftKey) this.doc.extendTo(h.r, h.c);
    else if (e.ctrlKey || e.metaKey) this.doc.select(makeRange(h.r, h.c), { r: h.r, c: h.c }, true);
    else this.doc.selectCell(h.r, h.c);
    const sel = this.doc.selection;
    this.drag = { kind: 'select', startR: e.shiftKey ? sel.anchor.r : h.r, startC: e.shiftKey ? sel.anchor.c : h.c, pointerId: e.pointerId, additive: e.ctrlKey || e.metaKey };
    this.capture(e);
    this.focusIdle();
  }

  private capture(e: PointerEvent): void {
    try {
      (e.target as Element).setPointerCapture(e.pointerId);
    } catch {
      /* ignore */
    }
  }

  private onLongPress(): void {
    const d = this.drag;
    if (!d || d.kind !== 'touch' || d.moved) return;
    d.longPress = null;
    d.selecting = true;
    navigator.vibrate?.(10);
    const sel = this.doc.range;
    const inside = d.r >= sel.r0 && d.r <= sel.r1 && d.c >= sel.c0 && d.c <= sel.c1;
    if (!inside) this.doc.selectCell(d.r, d.c);
  }

  private cancelTouchTimer(): void {
    const d = this.drag;
    if (d?.kind === 'touch' && d.longPress) {
      clearTimeout(d.longPress);
      d.longPress = null;
    }
  }

  private onTouchMove(e: TouchEvent): void {
    // While a long-press selection or a handle drag is active, the page must not scroll.
    const d = this.drag;
    if ((d?.kind === 'touch' && d.selecting) || d?.kind === 'handle' || d?.kind === 'resize' || d?.kind === 'image' || this.pinch || e.touches.length > 1) {
      if (e.cancelable) e.preventDefault();
    }
  }

  private onHandleDown(e: PointerEvent, corner: 'start' | 'end'): void {
    e.preventDefault();
    e.stopPropagation();
    const sel = this.doc.range;
    // Anchor the selection at the opposite corner.
    const anchor = corner === 'end' ? { r: sel.r0, c: sel.c0 } : { r: sel.r1, c: sel.c1 };
    this.doc.setSelection({ active: this.doc.selection.active, anchor, ranges: [sel] });
    this.drag = { kind: 'handle', corner, pointerId: e.pointerId };
    this.capture(e);
  }

  private onPointerMove(e: PointerEvent): void {
    const p = this.local(e);
    if (this.pointers.has(e.pointerId)) this.pointers.set(e.pointerId, p);
    if (this.pinch && this.pointers.size >= 2) {
      const [a, b] = [...this.pointers.values()];
      const dist = Math.hypot(a.x - b.x, a.y - b.y);
      if (this.pinch.dist > 0) this.setZoom((this.pinch.zoom * dist) / this.pinch.dist, { x: this.pinch.cx, y: this.pinch.cy });
      return;
    }
    const d = this.drag;
    if (!d) {
      if (!isTouchPointer(e)) this.hover(p.x, p.y, e);
      return;
    }
    if (d.pointerId !== e.pointerId) return;
    switch (d.kind) {
      case 'touch': {
        if (!d.selecting) {
          if (Math.hypot(p.x - d.x, p.y - d.y) > 8) {
            d.moved = true;
            this.cancelTouchTimer();
          }
          return;
        }
        const h = this.hit(p.x, p.y);
        this.doc.extendTo(h.r, h.c);
        this.updateAutoScroll(p);
        return;
      }
      case 'select': {
        const vt = this.transform();
        const r = vt.rowAt(p.y);
        const c = vt.colAt(p.x);
        const range = this.sheet.expandToMerges(makeRange(d.startR, d.startC, r, c));
        const sel = this.doc.selection;
        const ranges = d.additive ? [...sel.ranges.slice(0, -1), range] : [range];
        this.doc.setSelection({ active: sel.active, anchor: { r: d.startR, c: d.startC }, ranges });
        this.updateAutoScroll(p);
        return;
      }
      case 'point': {
        const vt = this.transform();
        this.insertPointRef(makeRange(d.startR, d.startC, vt.rowAt(p.y), vt.colAt(p.x)));
        this.updateAutoScroll(p);
        return;
      }
      case 'handle': {
        const vt = this.transform();
        this.doc.extendTo(vt.rowAt(p.y), vt.colAt(p.x));
        this.updateAutoScroll(p);
        return;
      }
      case 'header': {
        const vt = this.transform();
        if (d.axis === 'col') {
          const c = vt.colAt(p.x);
          this.doc.selectCols(d.start, c, d.additive);
        } else {
          const r = vt.rowAt(p.y);
          this.doc.selectRows(d.start, r, d.additive);
        }
        this.updateAutoScroll(p);
        return;
      }
      case 'resize': {
        const delta = (d.axis === 'col' ? p.x : p.y) - d.startPos;
        const size = Math.max(0, d.startSize + delta / this.zoom);
        const vt = this.transform();
        const rect = vt.cellRect(d.axis === 'row' ? d.index : 0, d.axis === 'col' ? d.index : 0);
        this.resizeGuide = { axis: d.axis, pos: (d.axis === 'col' ? rect.x : rect.y) + size * this.zoom };
        this.requestRender();
        return;
      }
      case 'fill': {
        const vt = this.transform();
        const r = vt.rowAt(p.y);
        const c = vt.colAt(p.x);
        const src = this.doc.range;
        const dr = r > src.r1 ? r - src.r1 : r < src.r0 ? r - src.r0 : 0;
        const dc = c > src.c1 ? c - src.c1 : c < src.c0 ? c - src.c0 : 0;
        if (Math.abs(dr) >= Math.abs(dc) && dr) this.fillPreview = makeRange(Math.min(src.r0, r), src.c0, Math.max(src.r1, r), src.c1);
        else if (dc) this.fillPreview = makeRange(src.r0, Math.min(src.c0, c), src.r1, Math.max(src.c1, c));
        else this.fillPreview = undefined;
        this.updateAutoScroll(p);
        this.requestRender();
        return;
      }
      case 'image': {
        const dx = (p.x - d.startX) / this.zoom;
        const dy = (p.y - d.startY) / this.zoom;
        const r = { ...d.rect };
        if (d.mode === 'move') {
          r.x += dx;
          r.y += dy;
        } else {
          if (d.mode.includes('e')) r.w = Math.max(8, d.rect.w + dx);
          if (d.mode.includes('s')) r.h = Math.max(8, d.rect.h + dy);
          if (d.mode.includes('w')) {
            r.w = Math.max(8, d.rect.w - dx);
            r.x = d.rect.x + d.rect.w - r.w;
          }
          if (d.mode.includes('n')) {
            r.h = Math.max(8, d.rect.h - dy);
            r.y = d.rect.y + d.rect.h - r.h;
          }
          if (d.mode.length === 2 && (e.shiftKey || (isTouchPointer(e) && d.objKind === 'image'))) {
            // Corners keep the aspect ratio with Shift (always on touch).
            const ratio = d.rect.w / d.rect.h;
            if (r.w / r.h > ratio) r.w = r.h * ratio;
            else r.h = r.w / ratio;
            if (d.mode.includes('w')) r.x = d.rect.x + d.rect.w - r.w;
            if (d.mode.includes('n')) r.y = d.rect.y + d.rect.h - r.h;
          }
        }
        if (Math.abs(dx) + Math.abs(dy) > 1) d.moved = true;
        const el = this.imageEls.get(d.id);
        if (el) this.placeImageEl(el, r, this.transform());
        return;
      }
    }
  }

  private onPointerUp(e: PointerEvent): void {
    this.pointers.delete(e.pointerId);
    if (this.pinch) {
      if (this.pointers.size < 2) this.pinch = null;
      return;
    }
    const d = this.drag;
    this.stopAutoScroll();
    if (!d || d.pointerId !== e.pointerId) return;
    this.drag = null;
    const p = this.local(e);
    switch (d.kind) {
      case 'touch': {
        this.cancelTouchTimer();
        // The grid handles this tap itself: suppress the browser's emulated mouse events and click,
        // which would otherwise land on a menu or editor opened by the tap ("ghost click").
        this.swallowTouchEnd = !d.moved || d.selecting;
        if (d.selecting) {
          // Long press without moving: show the menu.
          if (Math.hypot(p.x - d.x, p.y - d.y) <= 8) this.cb.onContextMenu?.(e.clientX, e.clientY, 'cell');
          return;
        }
        if (d.moved) return;
        this.onTap(d.r, d.c, e);
        return;
      }
      case 'resize': {
        this.resizeGuide = undefined;
        const delta = ((d.axis === 'col' ? p.x : p.y) - d.startPos) / this.zoom;
        if (Math.abs(delta) >= 1) {
          const size = Math.max(0, d.startSize + delta);
          if (d.axis === 'col') this.doc.setColWidth(d.indexes, size);
          else this.doc.setRowHeight(d.indexes, (size * 72) / 96);
        }
        this.requestRender();
        return;
      }
      case 'fill': {
        const target = this.fillPreview;
        this.fillPreview = undefined;
        if (target) this.doc.fill(target, e.ctrlKey || e.metaKey);
        this.requestRender();
        return;
      }
      case 'image': {
        const el = this.imageEls.get(d.id);
        if (el) {
          const rect = this.rectFromEl(el);
          if (Math.abs(rect.x - d.rect.x) > 0.5 || Math.abs(rect.y - d.rect.y) > 0.5 || Math.abs(rect.w - d.rect.w) > 0.5 || Math.abs(rect.h - d.rect.h) > 0.5) {
            this.doc.setObjectRect(d.id, { x: Math.max(0, rect.x), y: Math.max(0, rect.y), width: rect.w, height: rect.h });
          }
        }
        if (!d.moved && isTouchPointer(e)) {
          // Double tap on a chart or shape opens its settings.
          const now = Date.now();
          if (this.lastObjectTap.id === d.id && now - this.lastObjectTap.time < DOUBLE_TAP_MS * 1.5 && d.objKind !== 'image') {
            this.lastObjectTap = { id: '', time: 0 };
            this.cb.onObjectEdit?.(d.id);
          } else this.lastObjectTap = { id: d.id, time: now };
        }
        return;
      }
      default:
        return;
    }
  }

  private onPointerCancel(e: PointerEvent): void {
    this.pointers.delete(e.pointerId);
    if (this.pointers.size < 2) this.pinch = null;
    this.cancelTouchTimer();
    if (this.drag?.pointerId === e.pointerId) {
      this.drag = null;
      this.fillPreview = undefined;
      this.resizeGuide = undefined;
      this.stopAutoScroll();
      this.requestRender();
    }
  }

  private onTap(r: number, c: number, e: PointerEvent): void {
    const now = Date.now();
    const sel = this.doc.selection;
    const merge = this.sheet.mergeAt(r, c);
    const ar = merge ? merge.r0 : r;
    const ac = merge ? merge.c0 : c;
    const doubleTap = now - this.lastTap.time < DOUBLE_TAP_MS && this.lastTap.r === ar && this.lastTap.c === ac;
    this.lastTap = { time: now, r: ar, c: ac };
    this.cancelPendingTap();
    if (this.edit.active) {
      if (this.isFormulaEdit() && this.canInsertRef()) {
        this.insertPointRef(makeRange(r, c));
        return;
      }
      this.commitEdit();
    }
    const link = this.sheet.hyperlinkAt(ar, ac);
    if (doubleTap && !this.doc.isReadOnly()) {
      this.startEdit(undefined, 'edit');
      return;
    }
    const wasActive = sel.active.r === ar && sel.active.c === ac && sel.ranges.length === 1;
    if (!wasActive) this.doc.selectCell(r, c);
    if (!wasActive && !link) return;
    // Tapping the active cell opens its menu (or a link's popup), after waiting out a possible
    // second tap so that double-tap still edits.
    const x = e.clientX;
    const y = e.clientY;
    this.pendingTap = window.setTimeout(() => {
      this.pendingTap = 0;
      if (link) this.cb.onLinkTap?.({ url: link.target, location: link.location, x, y });
      else this.cb.onContextMenu?.(x, y, 'cell');
    }, this.doc.isReadOnly() ? 0 : DOUBLE_TAP_MS);
  }

  private cancelPendingTap(): void {
    if (this.pendingTap) clearTimeout(this.pendingTap);
    this.pendingTap = 0;
  }


  private onDoubleClick(e: MouseEvent): void {
    const p = this.local(e);
    const h = this.hit(p.x, p.y);
    if (h.resize) {
      if (h.resize.axis === 'col') {
        const sel = this.doc.range;
        const cols = sel.r0 === 0 && sel.r1 === MAX_ROWS - 1 && h.resize.index >= sel.c0 && h.resize.index <= sel.c1 ? range(sel.c0, sel.c1) : [h.resize.index];
        this.doc.autoFitColumns(cols);
      } else this.doc.autoFitRows([h.resize.index]);
      return;
    }
    if (h.area !== 'cell' || this.doc.isReadOnly()) return;
    if (this.edit.active) return;
    this.doc.selectCell(h.r, h.c);
    this.startEdit(undefined, 'edit', p);
  }

  private onContextMenu(e: MouseEvent): void {
    e.preventDefault();
    const p = this.local(e);
    const h = this.hit(p.x, p.y);
    if (this.drag?.kind === 'touch') return;
    if (h.area === 'cell') {
      const sel = this.doc.selection;
      const inside = sel.ranges.some((rg) => h.r >= rg.r0 && h.r <= rg.r1 && h.c >= rg.c0 && h.c <= rg.c1);
      if (!inside) this.doc.selectCell(h.r, h.c);
    } else if (h.area === 'col') {
      const rg = this.doc.range;
      if (!(rg.r0 === 0 && rg.r1 === MAX_ROWS - 1 && h.c >= rg.c0 && h.c <= rg.c1)) this.doc.selectCols(h.c, h.c);
    } else if (h.area === 'row') {
      const rg = this.doc.range;
      if (!(rg.c0 === 0 && rg.c1 === MAX_COLS - 1 && h.r >= rg.r0 && h.r <= rg.r1)) this.doc.selectRows(h.r, h.r);
    }
    if (this.edit.active) this.commitEdit();
    this.cb.onContextMenu?.(e.clientX, e.clientY, h.area === 'none' ? 'cell' : h.area);
  }

  private onWheel(e: WheelEvent): void {
    if (e.ctrlKey || e.metaKey) {
      e.preventDefault();
      const p = this.local(e);
      const factor = Math.exp(-e.deltaY / 400);
      this.setZoom(this.zoom * factor, p);
    }
  }

  private hover(x: number, y: number, e: PointerEvent): void {
    const h = this.hit(x, y);
    let cursor = 'cell';
    if (h.resize) cursor = h.resize.axis === 'col' ? 'col-resize' : 'row-resize';
    else if (h.area === 'col') cursor = 's-resize';
    else if (h.area === 'row') cursor = 'e-resize';
    else if (h.area === 'corner') cursor = 'default';
    else if (h.area === 'cell' && this.nearFillHandle(x, y)) cursor = 'crosshair';
    else if (h.area === 'cell' && this.sheet.hyperlinkAt(h.r, h.c) && (e.ctrlKey || e.metaKey || this.doc.isReadOnly())) cursor = 'pointer';
    this.scroller.style.cursor = cursor;
    if (h.area === 'cell') {
      const m = this.sheet.mergeAt(h.r, h.c);
      const t = this.sheet.commentAt(m?.r0 ?? h.r, m?.c0 ?? h.c);
      const link = this.sheet.hyperlinkAt(h.r, h.c);
      if (t) this.showTooltip(m ?? makeRange(h.r, h.c), t.comments.map((cm) => `<b>${escapeHtml(cm.author)}</b><br>${escapeHtml(cm.text).replace(/\n/g, '<br>')}`).join('<hr>'));
      else if (link) this.showTooltip(makeRange(h.r, h.c), `${escapeHtml(link.tooltip ?? link.target ?? link.location ?? '')}<br><small>Ctrl+click to follow</small>`);
      else this.hideTooltip();
    } else this.hideTooltip();
  }

  private showTooltip(range: Range, html: string): void {
    const vt = this.transform();
    const rect = vt.rangeRect(range);
    this.tooltip.innerHTML = html;
    this.tooltip.style.display = 'block';
    const left = Math.min(rect.x + rect.w + 6, this.width - 260);
    this.tooltip.style.left = `${Math.max(4, left)}px`;
    this.tooltip.style.top = `${Math.max(4, Math.min(rect.y, this.height - 120))}px`;
  }

  private hideTooltip(): void {
    this.tooltip.style.display = 'none';
  }

  // ---------------------------------------------------------------------------------------------
  // Auto scroll while dragging near the edges

  private updateAutoScroll(p: { x: number; y: number }): void {
    const vt = this.transform();
    const margin = 24;
    let dx = 0;
    let dy = 0;
    if (p.x > this.width - margin) dx = (p.x - (this.width - margin)) / 2 + 2;
    else if (p.x < vt.mainLeft && p.x < this.width) dx = -((vt.mainLeft - p.x) / 2 + 2);
    if (p.y > this.height - margin) dy = (p.y - (this.height - margin)) / 2 + 2;
    else if (p.y < vt.mainTop) dy = -((vt.mainTop - p.y) / 2 + 2);
    if (!dx && !dy) {
      this.stopAutoScroll();
      return;
    }
    this.autoScroll = { dx, dy, x: p.x, y: p.y };
    if (!this.autoScrollFrame) this.autoScrollFrame = requestAnimationFrame(() => this.autoScrollTick());
  }

  private autoScrollTick(): void {
    this.autoScrollFrame = 0;
    const a = this.autoScroll;
    if (!a || !this.drag) return;
    this.scroller.scrollLeft += a.dx;
    this.scroller.scrollTop += a.dy;
    // Re-run the drag update at the same pointer position.
    const fake = { clientX: a.x + this.root.getBoundingClientRect().left, clientY: a.y + this.root.getBoundingClientRect().top, pointerId: this.drag.pointerId, pointerType: 'mouse', shiftKey: false } as PointerEvent;
    this.onPointerMove(fake);
    this.autoScrollFrame = requestAnimationFrame(() => this.autoScrollTick());
  }

  private stopAutoScroll(): void {
    this.autoScroll = null;
    if (this.autoScrollFrame) cancelAnimationFrame(this.autoScrollFrame);
    this.autoScrollFrame = 0;
  }

  /** Scrolls so that a cell is fully visible. */
  scrollIntoView(r: number, c: number): void {
    const vt = this.transform();
    const layout = vt.layout;
    const z = this.zoom;
    if (r >= vt.fr) {
      const top = (layout.rowY(r) - vt.FH) * z;
      const bottom = (layout.rowY(r + 1) - vt.FH) * z;
      const viewH = this.height - vt.mainTop;
      if (top < this.scroller.scrollTop) this.scroller.scrollTop = top;
      else if (bottom > this.scroller.scrollTop + viewH) this.scroller.scrollTop = bottom - viewH;
    }
    if (c >= vt.fc) {
      const left = (layout.colX(c) - vt.FW) * z;
      const right = (layout.colX(c + 1) - vt.FW) * z;
      const viewW = this.width - vt.mainLeft;
      if (left < this.scroller.scrollLeft) this.scroller.scrollLeft = left;
      else if (right > this.scroller.scrollLeft + viewW) this.scroller.scrollLeft = Math.min(left, right - viewW);
    }
  }

  // ---------------------------------------------------------------------------------------------
  // Mobile selection handles

  private positionHandles(vt: ViewTransform): void {
    const show = this.mobile && !this.edit.active;
    const range = this.doc.range;
    const rect = vt.rangeRect(range);
    const [start, end] = this.handles;
    const visible = (x: number, y: number) => x >= vt.v.headerW - 2 && y >= vt.v.headerH - 2 && x <= this.width + 2 && y <= this.height + 2;
    const place = (el: HTMLDivElement, x: number, y: number) => {
      el.style.display = show && visible(x, y) ? 'block' : 'none';
      el.style.transform = `translate(${x}px, ${y}px)`;
    };
    place(start, rect.x, rect.y);
    place(end, rect.x + rect.w, rect.y + rect.h);
  }

  // ---------------------------------------------------------------------------------------------
  // Pictures, charts and shapes

  /** Pixel rectangle (100% zoom) of a picture, chart or shape of the active sheet. */
  private objectRect(id: string): Rect | null {
    const hit = this.doc.objectById(id);
    if (!hit) return null;
    const layout = SheetLayout.of(this.sheet);
    if (hit.kind === 'image') {
      const r = layout.imageRect(hit.obj);
      return { x: r.x, y: r.y, w: r.w, h: r.h };
    }
    return layout.anchorRect(hit.obj.from, hit.obj.to);
  }

  private createObjectEl(id: string, kind: 'image' | 'chart' | 'shape'): HTMLDivElement {
    const el = document.createElement('div');
    el.className = `xgrid-image xgrid-object xgrid-${kind}obj`;
    el.dataset.id = id;
    el.dataset.kind = kind;
    el.setAttribute('data-testid', `object-${kind}`);
    const body = document.createElement('div');
    body.className = 'xgrid-object-body';
    el.appendChild(body);
    if (kind === 'image') {
      const img = this.sheet.images.find((i) => i.id === id)!;
      const url = URL.createObjectURL(new Blob([img.data as BlobPart], { type: img.mime }));
      el.dataset.url = url;
      const pic = document.createElement('img');
      pic.src = url;
      pic.alt = img.descr ?? img.name;
      pic.draggable = false;
      body.appendChild(pic);
    } else {
      el.setAttribute('role', 'img');
    }
    for (const handle of ['nw', 'n', 'ne', 'e', 'se', 's', 'sw', 'w']) {
      const h = document.createElement('span');
      h.className = `xgrid-image-handle h-${handle}`;
      h.dataset.handle = handle;
      el.appendChild(h);
    }
    el.addEventListener('pointerdown', (e) => this.onImagePointerDown(e, id));
    el.addEventListener('pointermove', (e) => this.onPointerMove(e));
    el.addEventListener('pointerup', (e) => this.onPointerUp(e));
    el.addEventListener('pointercancel', (e) => this.onPointerCancel(e));
    el.addEventListener('dblclick', (e) => {
      e.preventDefault();
      e.stopPropagation();
      if (el.dataset.kind !== 'image') this.cb.onObjectEdit?.(id);
    });
    el.addEventListener('contextmenu', (e) => {
      e.preventDefault();
      e.stopPropagation();
      this.selectImage(id);
      this.cb.onContextMenu?.(e.clientX, e.clientY, 'image');
    });
    return el;
  }

  /** Draws the chart or shape (only when its markup changed: data, settings or size). */
  private renderObjectContent(el: HTMLElement, id: string): void {
    const hit = this.doc.objectById(id);
    if (!hit || hit.kind === 'image') return;
    const rect = this.objectRect(id);
    if (!rect) return;
    let markup: string;
    if (hit.kind === 'chart') {
      markup = chartSvg(hit.obj, readChartData(this.doc, hit.obj, this.sheet), rect.w, rect.h);
      el.setAttribute('aria-label', hit.obj.title || hit.obj.name || 'Chart');
    } else {
      markup = shapeSvg(hit.obj, rect.w, rect.h);
      el.setAttribute('aria-label', hit.obj.text || hit.obj.name || 'Shape');
    }
    if (this.objectMarkup.get(id) === markup) return;
    this.objectMarkup.set(id, markup);
    (el.firstElementChild as HTMLElement).innerHTML = markup;
  }

  private syncImages(): void {
    const sheet = this.sheet;
    const objects: { id: string; kind: 'image' | 'chart' | 'shape'; z: number }[] = [
      ...sheet.images.map((o, i) => ({ id: o.id, kind: 'image' as const, z: o.z ?? i })),
      ...sheet.charts.map((o) => ({ id: o.id, kind: 'chart' as const, z: o.z })),
      ...sheet.shapes.map((o) => ({ id: o.id, kind: 'shape' as const, z: o.z })),
    ];
    const ids = new Set(objects.map((o) => o.id));
    for (const [id, el] of this.imageEls) {
      if (ids.has(id) && el.dataset.sheet === String(sheet.uid)) continue;
      el.remove();
      this.imageEls.delete(id);
      this.objectMarkup.delete(id);
      if (el.dataset.url) URL.revokeObjectURL(el.dataset.url);
    }
    for (const o of objects) {
      let el = this.imageEls.get(o.id);
      if (!el) {
        el = this.createObjectEl(o.id, o.kind);
        el.dataset.sheet = String(sheet.uid);
        this.imageEls.set(o.id, el);
      }
      el.style.zIndex = String(1 + Math.max(0, Math.round(o.z)));
      if (o.kind !== 'image') this.renderObjectContent(el, o.id);
    }
    // Other drawing objects of the file (kept unchanged, shown read-only).
    for (const el of [...this.mainImages.querySelectorAll('.xgrid-chart'), ...this.frozenImages.querySelectorAll('.xgrid-chart')]) el.remove();
    for (const ph of chartPlaceholders(this.doc, sheet)) {
      const el = document.createElement('div');
      el.className = 'xgrid-chart';
      el.dataset.rect = JSON.stringify(ph.rect);
      el.appendChild(ph.element);
      this.mainImages.appendChild(el);
    }
    if (this.selectedImage && !ids.has(this.selectedImage)) this.selectImage(null);
    this.requestRender();
  }

  private onImagePointerDown(e: PointerEvent, id: string): void {
    e.stopPropagation();
    if (e.button === 2) return;
    if (this.edit.active) this.commitEdit();
    this.selectImage(id);
    if (this.doc.isReadOnly()) return;
    e.preventDefault();
    const rect = this.objectRect(id);
    const hit = this.doc.objectById(id);
    if (!rect || !hit) return;
    const p = this.local(e);
    const handle = (e.target as HTMLElement).dataset.handle;
    this.drag = { kind: 'image', id, objKind: hit.kind, mode: handle ?? 'move', startX: p.x, startY: p.y, rect, pointerId: e.pointerId, moved: false };
    this.capture(e);
  }

  selectImage(id: string | null): void {
    this.selectedImage = id;
    for (const [iid, el] of this.imageEls) el.classList.toggle('selected', iid === id);
    this.cb.onImageSelect?.(id);
    if (id) this.focusIdle();
  }

  get selectedImageId(): string | null {
    return this.selectedImage;
  }

  /** Moves the selected object by a few pixels (arrow keys). */
  private nudgeObject(dx: number, dy: number): void {
    const id = this.selectedImage;
    const rect = id ? this.objectRect(id) : null;
    if (!id || !rect || this.doc.isReadOnly()) return;
    this.doc.setObjectRect(id, { x: Math.max(0, rect.x + dx), y: Math.max(0, rect.y + dy), width: rect.w, height: rect.h });
  }

  private placeImageEl(el: HTMLElement, rect: Rect, vt: ViewTransform): void {
    const frozen = rect.x < vt.FW || rect.y < vt.FH;
    const x = vt.screenX(rect.x);
    const y = vt.screenY(rect.y);
    el.style.transform = `translate(${x}px, ${y}px)`;
    el.style.width = `${rect.w * vt.zoom}px`;
    el.style.height = `${rect.h * vt.zoom}px`;
    el.dataset.sheetRect = `${rect.x},${rect.y},${rect.w},${rect.h}`;
    const parent = frozen ? this.frozenImages : this.mainImages;
    if (el.parentElement !== parent) parent.appendChild(el);
  }

  private rectFromEl(el: HTMLElement): Rect {
    const [x, y, w, h] = (el.dataset.sheetRect ?? '0,0,0,0').split(',').map(Number);
    return { x, y, w, h };
  }

  private positionImages(vt: ViewTransform): void {
    const clipMain = `${vt.mainLeft}px`;
    this.mainImages.style.clipPath = `inset(${vt.mainTop}px 0 0 ${clipMain})`;
    this.frozenImages.style.clipPath = `inset(${vt.v.headerH}px 0 0 ${vt.v.headerW}px)`;
    for (const [id, el] of this.imageEls) {
      if (this.drag?.kind === 'image' && this.drag.id === id) continue;
      const r = this.objectRect(id);
      if (r) this.placeImageEl(el, r, vt);
    }
    for (const el of this.mainImages.querySelectorAll<HTMLElement>('.xgrid-chart')) {
      const r = JSON.parse(el.dataset.rect ?? '{}') as Rect;
      el.style.transform = `translate(${vt.screenX(r.x)}px, ${vt.screenY(r.y)}px)`;
      el.style.width = `${r.w * vt.zoom}px`;
      el.style.height = `${r.h * vt.zoom}px`;
    }
  }

  // ---------------------------------------------------------------------------------------------
  // Data validation lists

  private validationList(r: number, c: number): string[] | null {
    const dv = this.sheet.validations.find((v) => v.type === 'list' && !v.hideDropDown && v.sqref.some((rg) => r >= rg.r0 && r <= rg.r1 && c >= rg.c0 && c <= rg.c1));
    if (!dv?.formula1) return null;
    const f = dv.formula1.trim();
    if (f.startsWith('"')) return f.slice(1, -1).split(',').map((s) => s.trim());
    const v = this.doc.engine.evaluateText(f.replace(/^=/, ''), this.sheet, r, c);
    const out: string[] = [];
    if (v && typeof v === 'object' && 'areas' in v) {
      for (const area of v.areas) {
        for (let rr = area.range.r0; rr <= Math.min(area.range.r1, area.range.r0 + 500); rr++) {
          for (let cc = area.range.c0; cc <= area.range.c1; cc++) {
            const val = this.doc.engine.cellValue(area.sheet, rr, cc);
            if (val !== null && val !== '') out.push(formatValue(val, this.doc.wb.styles.get(area.sheet.styleAt(rr, cc)).numFmt).text);
          }
        }
      }
    }
    return out;
  }

  private positionValidationButton(vt: ViewTransform): void {
    const { r, c } = this.doc.selection.active;
    const list = !this.doc.isReadOnly() && !this.edit.active ? this.sheet.validations.length && this.validationList(r, c) : null;
    if (!list || !list.length) {
      this.dvButton.style.display = 'none';
      return;
    }
    const m = this.sheet.mergeAt(r, c);
    const rect = m ? vt.rangeRect(m) : vt.cellRect(r, c);
    const s = Math.max(16, Math.min(rect.h, 22));
    this.dvButton.style.display = 'flex';
    this.dvButton.style.width = `${s}px`;
    this.dvButton.style.height = `${s}px`;
    this.dvButton.style.transform = `translate(${rect.x + rect.w + 2}px, ${rect.y + rect.h - s}px)`;
  }

  private openValidationList(): void {
    const { r, c } = this.doc.selection.active;
    const list = this.validationList(r, c);
    if (!list) return;
    this.cb.onValidationList?.(list, this.dvButton.getBoundingClientRect(), (value) => {
      this.doc.setInput(r, c, value);
      this.focusIdle();
    });
  }

  // ---------------------------------------------------------------------------------------------
  // Cell editor

  /** Focus target while not editing: the hidden editor on desktop (keyboard, IME, clipboard). */
  focusIdle(): void {
    if (this.edit.active) return;
    if (this.mobile) {
      this.scroller.focus({ preventScroll: true });
      return;
    }
    this.syncEditorIdle();
    this.editor.focus({ preventScroll: true });
  }

  /** Gives the focus back to the cell being edited (after a dialog), leaving its caret where it was. */
  focusEdit(): void {
    if (this.edit.active && this.edit.origin === 'cell') this.editor.focus({ preventScroll: true });
  }

  private syncEditorIdle(): void {
    if (this.edit.active) return;
    this.editor.value = '';
    this.editor.classList.remove('editing');
    this.editor.readOnly = this.mobile;
    this.requestRender();
  }

  get editState(): EditState {
    return this.edit;
  }

  isEditing(): boolean {
    return this.edit.active;
  }

  private isFormulaEdit(): boolean {
    return this.edit.active && this.edit.text.startsWith('=');
  }

  /** Starts editing the active cell. `initial` replaces the content (typing), otherwise it is edited. */
  startEdit(initial?: string, mode: 'enter' | 'edit' = initial === undefined ? 'edit' : 'enter', clickPoint?: { x: number; y: number }, origin: 'cell' | 'bar' = 'cell'): void {
    if (this.doc.isReadOnly()) return;
    const sel = this.doc.selection;
    const sheet = this.sheet;
    const anchor = this.doc.anchorOf(sel.active.r, sel.active.c);
    const r = anchor && sheet.getCell(anchor.r, anchor.c)?.arr && !sheet.getCell(anchor.r, anchor.c)?.dyn ? anchor.r : sel.active.r;
    const c = anchor && sheet.getCell(anchor.r, anchor.c)?.arr && !sheet.getCell(anchor.r, anchor.c)?.dyn ? anchor.c : sel.active.c;
    const text = initial ?? this.doc.editText(r, c);
    this.edit = { active: true, text, mode, origin, caret: text.length, r, c, sheet };
    this.pointRef = null;
    this.editor.readOnly = false;
    this.editor.value = text;
    this.editor.classList.add('editing');
    this.scrollIntoView(r, c);
    this.paint();
    if (origin === 'cell') {
      this.editor.focus({ preventScroll: true });
      const pos = clickPoint ? this.caretFromPoint(clickPoint) : text.length;
      this.editor.setSelectionRange(pos, pos);
      this.edit.caret = pos;
    }
    this.emitEdit();
  }

  private caretFromPoint(p: { x: number; y: number }): number {
    const text = this.edit.text;
    if (text.includes('\n')) return text.length;
    const rect = this.editor.getBoundingClientRect();
    const rootRect = this.root.getBoundingClientRect();
    const x = p.x - (rect.left - rootRect.left) - 3;
    const font = this.editor.style.font;
    let best = text.length;
    for (let i = 0; i <= text.length; i++) {
      if (textWidth(text.slice(0, i), font) >= x) {
        best = i;
        break;
      }
    }
    return best;
  }

  /** Called by the formula bar while it is being edited. */
  setEditFromBar(text: string, caret: number): void {
    if (!this.edit.active) this.startEdit(text, 'edit', undefined, 'bar');
    this.edit.text = text;
    this.edit.caret = caret;
    this.edit.origin = 'bar';
    this.edit.mode = 'edit';
    this.editor.value = text;
    this.pointRef = null;
    this.requestRender();
    this.emitEdit();
  }

  private onCompositionStart(): void {
    if (!this.edit.active && !this.doc.isReadOnly()) {
      const sel = this.doc.selection;
      this.edit = { active: true, text: '', mode: 'enter', origin: 'cell', caret: 0, r: sel.active.r, c: sel.active.c, sheet: this.sheet };
      this.editor.classList.add('editing');
      this.paint();
      this.emitEdit();
    }
  }

  private onEditorInput(): void {
    if (!this.edit.active) {
      if (this.doc.isReadOnly()) {
        this.editor.value = '';
        return;
      }
      // Typing on a selected cell starts editing with the typed text. A bare line break is a
      // stray Enter from another control that handed focus to the grid; it never starts an edit.
      const text = this.editor.value;
      if (/^[\r\n]+$/.test(text)) {
        this.editor.value = '';
        return;
      }
      const sel = this.doc.selection;
      this.edit = { active: true, text, mode: 'enter', origin: 'cell', caret: text.length, r: sel.active.r, c: sel.active.c, sheet: this.sheet };
      this.editor.classList.add('editing');
      this.paint();
    }
    this.edit.text = this.editor.value;
    this.edit.caret = this.editor.selectionStart ?? this.edit.text.length;
    this.edit.origin = 'cell';
    this.pointRef = null;
    this.paint();
    this.emitEdit();
  }

  private updateCaret(): void {
    this.edit.caret = this.editor.selectionStart ?? this.edit.caret;
    this.emitEdit();
  }

  private emitEdit(): void {
    this.cb.onEdit?.({ ...this.edit });
  }

  /**
   * Moves the active cell after an entry. Like Excel, Enter after a run of Tabs returns to the
   * column where the run started (row-by-row data entry).
   */
  private advance(dr: number, dc: number): void {
    const sel = this.doc.selection;
    const single = sel.ranges.length === 1 && sel.ranges[0].r0 === sel.ranges[0].r1 && sel.ranges[0].c0 === sel.ranges[0].c1;
    const { r, c } = sel.active;
    if (single && dr === 0 && dc !== 0) {
      if (!this.tabStart || this.tabStart.r !== r) this.tabStart = { r, c };
    } else if (single && dr === 1 && dc === 0 && this.tabStart && this.tabStart.r === r) {
      const col = this.tabStart.c;
      this.tabStart = null;
      this.doc.selectCell(r + 1, col);
      return;
    } else this.tabStart = null;
    this.doc.cycleActive(dr, dc);
  }

  /** Writes the edit to the cell and optionally moves the selection. */
  commitEdit(move?: { dr: number; dc: number }, force = false): void {
    if (!this.edit.active) return;
    const { r, c, text, sheet } = this.edit;
    if (!force && sheet && this.cb.validateEdit && text !== this.doc.editText(r, c, sheet)) {
      const taken = this.cb.validateEdit(r, c, text, sheet, (accept) => {
        if (accept) this.commitEdit(move, true);
        else if (this.edit.active) {
          this.editor.focus({ preventScroll: true });
          this.editor.setSelectionRange(0, this.editor.value.length);
        }
      });
      if (taken) return;
    }
    this.edit = { ...this.edit, active: false };
    this.pointRef = null;
    this.editor.classList.remove('editing');
    this.editor.value = '';
    if (sheet && sheet !== this.sheet) this.doc.setActiveSheet(this.doc.wb.sheets.indexOf(sheet));
    if (sheet && text !== this.doc.editText(r, c, sheet)) this.doc.setInput(r, c, text, sheet);
    if (move) this.advance(move.dr, move.dc);
    this.emitEdit();
    this.syncEditorIdle();
    this.focusIdle();
    this.requestRender();
  }

  cancelEdit(): void {
    if (!this.edit.active) return;
    const sheet = this.edit.sheet;
    this.edit = { ...this.edit, active: false };
    this.pointRef = null;
    this.editor.classList.remove('editing');
    this.editor.value = '';
    if (sheet && sheet !== this.sheet) this.doc.setActiveSheet(this.doc.wb.sheets.indexOf(sheet));
    this.emitEdit();
    this.syncEditorIdle();
    this.focusIdle();
    this.requestRender();
  }

  /** Replaces the current edit text (autocomplete) and places the caret. */
  replaceEditText(text: string, caret: number): void {
    if (!this.edit.active) return;
    this.edit.text = text;
    this.edit.caret = caret;
    this.editor.value = text;
    if (this.edit.origin === 'cell') {
      this.editor.focus({ preventScroll: true });
      this.editor.setSelectionRange(caret, caret);
    }
    this.pointRef = null;
    this.paint();
    this.emitEdit();
  }

  private positionEditor(vt: ViewTransform): void {
    const ed = this.editor;
    const sel = this.doc.selection;
    if (!this.edit.active) {
      // Idle: a tiny transparent box at the active cell (keeps IME candidate windows nearby).
      const rect = vt.cellRect(sel.active.r, sel.active.c);
      ed.style.transform = `translate(${Math.max(0, rect.x)}px, ${Math.max(0, rect.y)}px)`;
      ed.style.width = '2px';
      ed.style.height = '2px';
      return;
    }
    if (this.edit.sheet !== this.sheet) {
      ed.style.transform = 'translate(-10000px, 0)';
      return;
    }
    const { r, c } = this.edit;
    const m = this.sheet.mergeAt(r, c);
    const rect = m ? vt.rangeRect(m) : vt.cellRect(r, c);
    const style = this.doc.styleAt(r, c);
    const font = cssFont(style.font, this.zoom);
    const wrap = !!style.align?.wrap;
    const lines = this.edit.text.split('\n');
    const pad = 6;
    const contentW = Math.max(...lines.map((l) => textWidth(l, font))) + pad * 2 + 8;
    const maxW = this.width - rect.x - 4;
    const w = wrap ? rect.w : Math.max(rect.w, Math.min(contentW, maxW));
    const lh = lineHeightPx(style.font.size, this.zoom);
    let visualLines = lines.length;
    if (wrap) {
      visualLines = 0;
      for (const l of lines) visualLines += Math.max(1, Math.ceil(textWidth(l, font) / Math.max(10, w - pad * 2)));
    }
    const h = Math.max(rect.h, visualLines * lh + 6);
    ed.style.transform = `translate(${rect.x}px, ${rect.y}px)`;
    ed.style.width = `${w}px`;
    ed.style.height = `${Math.min(h, this.height - rect.y)}px`;
    ed.style.font = font;
    ed.style.lineHeight = `${lh}px`;
    ed.style.color = `#${resolveColor(style.font.color, this.doc.wb.theme, '000000')}`;
    ed.style.background = fillColor(style, this.doc.wb.theme) ?? '#ffffff';
    const hAlign = style.align?.horizontal;
    ed.style.textAlign = hAlign === 'center' || hAlign === 'right' ? hAlign : 'left';
    ed.style.whiteSpace = wrap ? 'pre-wrap' : 'pre';
  }

  // ---------------------------------------------------------------------------------------------
  // Formula point mode

  /** Whether a reference can be inserted at the caret (after =, (, operator, comma...). */
  private canInsertRef(): boolean {
    const text = this.edit.text;
    const caret = this.edit.origin === 'cell' ? (this.editor.selectionStart ?? this.edit.caret) : this.edit.caret;
    if (this.pointRef && caret === this.pointRef.end) return true;
    const before = text.slice(0, caret).replace(/\s+$/, '');
    return /[=(,;:+\-*/^&<>]$/.test(before) || before === '=';
  }

  private insertPointRef(range: Range): void {
    const text = this.edit.text;
    const caret = this.edit.origin === 'cell' ? (this.editor.selectionStart ?? this.edit.caret) : this.edit.caret;
    let start = caret;
    let end = caret;
    if (this.pointRef && caret === this.pointRef.end) {
      start = this.pointRef.start;
      end = this.pointRef.end;
    }
    const prefix = this.edit.sheet && this.edit.sheet !== this.sheet ? `${quoteSheetName(this.sheet.name)}!` : '';
    const ref = prefix + (range.r0 === range.r1 && range.c0 === range.c1 ? cellName(range.r0, range.c0) : rangeName(range));
    const next = text.slice(0, start) + ref + text.slice(end);
    this.pointRef = { start, end: start + ref.length };
    this.edit.text = next;
    this.edit.caret = start + ref.length;
    this.editor.value = next;
    if (this.edit.origin === 'cell') {
      this.editor.focus({ preventScroll: true });
      this.editor.setSelectionRange(this.edit.caret, this.edit.caret);
    }
    this.requestRender();
    this.emitEdit();
  }

  private formulaHighlights(): Highlight[] {
    if (!this.edit.active || !this.edit.text.startsWith('=')) return [];
    const out: Highlight[] = [];
    let i = 0;
    try {
      for (const t of tokenize(this.edit.text.slice(1))) {
        if (t.type !== 'ref' || !t.ref || t.ref.kind === 'error') continue;
        const sheet = t.ref.sheet !== undefined ? this.doc.wb.sheetByName(t.ref.sheet) : this.edit.sheet;
        if (sheet !== this.sheet) {
          i++;
          continue;
        }
        out.push({ range: makeRange(t.ref.r0, t.ref.c0, t.ref.r1, t.ref.c1), color: REF_COLORS[i++ % REF_COLORS.length], fill: true });
      }
    } catch {
      /* incomplete formula */
    }
    return out;
  }

  // ---------------------------------------------------------------------------------------------
  // Keyboard

  private onKeyDown(e: KeyboardEvent): void {
    if (this.cb.editorKey?.(e)) {
      e.preventDefault();
      return;
    }
    const ctrl = e.ctrlKey || e.metaKey;
    const key = e.key;
    if (this.edit.active) {
      this.editingKey(e, ctrl);
      return;
    }
    if (e.isComposing) return;
    const doc = this.doc;
    const shift = e.shiftKey;
    const handled = () => e.preventDefault();
    if (this.selectedImage && (key === 'Delete' || key === 'Backspace')) {
      if (!doc.isReadOnly()) doc.deleteObject(this.selectedImage);
      this.selectImage(null);
      handled();
      return;
    }
    if (this.selectedImage && key.startsWith('Arrow') && !ctrl) {
      const step = shift ? 10 : 1;
      this.nudgeObject(key === 'ArrowLeft' ? -step : key === 'ArrowRight' ? step : 0, key === 'ArrowUp' ? -step : key === 'ArrowDown' ? step : 0);
      handled();
      return;
    }
    if (this.selectedImage && key === 'Escape') {
      this.selectImage(null);
      handled();
      return;
    }
    switch (key) {
      case 'ArrowUp':
      case 'ArrowDown':
      case 'ArrowLeft':
      case 'ArrowRight': {
        const [dr, dc] = key === 'ArrowUp' ? [-1, 0] : key === 'ArrowDown' ? [1, 0] : key === 'ArrowLeft' ? [0, -1] : [0, 1];
        this.tabStart = null;
        doc.move(dr, dc, shift, ctrl);
        handled();
        return;
      }
      case 'Tab':
        this.advance(0, shift ? -1 : 1);
        handled();
        return;
      case 'Enter':
        if (e.altKey) return;
        this.advance(shift ? -1 : 1, 0);
        handled();
        return;
      case 'Home':
        if (ctrl) doc.selectCell(this.sheet.freeze.rows, this.sheet.freeze.cols);
        else doc.selectCell(doc.selection.active.r, 0);
        handled();
        return;
      case 'End': {
        if (ctrl) {
          const used = this.sheet.usedRange(false);
          doc.selectCell(used?.r1 ?? 0, used?.c1 ?? 0);
          handled();
        }
        return;
      }
      case 'PageDown':
      case 'PageUp': {
        const rows = Math.max(1, Math.floor((this.height - this.transform().mainTop) / (20 * this.zoom)) - 1);
        if (e.altKey) doc.move(0, (key === 'PageDown' ? 1 : -1) * 8, shift);
        else if (ctrl) this.cb.onCommand?.(key === 'PageDown' ? 'next-sheet' : 'prev-sheet');
        else doc.move(key === 'PageDown' ? rows : -rows, 0, shift);
        handled();
        return;
      }
      case 'F2':
        this.startEdit(undefined, 'edit');
        handled();
        return;
      case 'Delete':
        doc.clear('contents');
        handled();
        return;
      case 'Backspace':
        doc.clear('contents');
        this.startEdit('', 'enter');
        handled();
        return;
      case 'Escape':
        doc.clearClip();
        this.extraHighlights = this.extraHighlights.filter((h) => !h.dashed);
        this.requestRender();
        return;
      case 'ContextMenu': {
        const rect = this.transform().cellRect(doc.selection.active.r, doc.selection.active.c);
        const root = this.root.getBoundingClientRect();
        this.cb.onContextMenu?.(root.left + rect.x + rect.w / 2, root.top + rect.y + rect.h, 'cell');
        handled();
        return;
      }
      default:
        break;
    }
    if (shift && key === 'F10') {
      const rect = this.transform().cellRect(doc.selection.active.r, doc.selection.active.c);
      const root = this.root.getBoundingClientRect();
      this.cb.onContextMenu?.(root.left + rect.x + rect.w / 2, root.top + rect.y + rect.h, 'cell');
      handled();
      return;
    }
    if (shift && key === 'F2') {
      this.cb.onCommand?.('comment');
      handled();
      return;
    }
    if (ctrl) {
      const k = key.toLowerCase();
      if (k === 'z' && !shift) doc.undo();
      else if (k === 'y' || (k === 'z' && shift)) doc.redo();
      else if (k === 'a') doc.selectAll();
      else if (k === 'b') doc.toggleFont('bold');
      else if (k === 'i') doc.toggleFont('italic');
      else if (k === 'u') doc.toggleFont('underline');
      else if (k === '5') doc.toggleFont('strike');
      else if (k === 'd') this.fillFrom('down');
      else if (k === 'r') this.fillFrom('right');
      else if (k === 'f') this.cb.onCommand?.('find');
      else if (k === 'h') this.cb.onCommand?.('replace');
      else if (k === 'k') this.cb.onCommand?.('link');
      else if (k === 's') this.cb.onCommand?.('save');
      else if (k === 'o') this.cb.onCommand?.('open');
      else if (k === 'p') this.cb.onCommand?.('print');
      else if (key === '\\') this.cb.onCommand?.('clear-format');
      else if (key === ' ' && !shift) doc.selectCols(doc.range.c0, doc.range.c1);
      else if (key === ';' || key === ':') {
        const now = new Date();
        const text = shift || key === ':' ? `${now.getHours()}:${String(now.getMinutes()).padStart(2, '0')}` : `${now.getMonth() + 1}/${now.getDate()}/${now.getFullYear()}`;
        doc.setInput(doc.selection.active.r, doc.selection.active.c, text);
      } else if (k === 'c' || k === 'x' || k === 'v') return;
      else return;
      handled();
      return;
    }
    if (shift && key === ' ') {
      doc.selectRows(doc.range.r0, doc.range.r1);
      handled();
      return;
    }
    if (e.altKey && (key === '=' || key === '+')) {
      this.autoSum();
      handled();
    }
  }

  private editingKey(e: KeyboardEvent, ctrl: boolean): void {
    const key = e.key;
    if (e.isComposing) return;
    const finish = (dr: number, dc: number) => {
      e.preventDefault();
      this.commitEdit({ dr, dc });
    };
    switch (key) {
      case 'Enter':
        if (e.altKey || (ctrl && !e.shiftKey && false)) {
          // Alt+Enter: a line break inside the cell.
          e.preventDefault();
          const el = this.editor;
          const s = el.selectionStart ?? el.value.length;
          const t = el.value.slice(0, s) + '\n' + el.value.slice(el.selectionEnd ?? s);
          this.replaceEditText(t, s + 1);
          return;
        }
        if (ctrl) {
          // Ctrl+Enter: fill the selection with the entry.
          e.preventDefault();
          const text = this.edit.text;
          this.cancelEdit();
          this.doc.transact('Fill selection', () => {
            for (const rg of this.doc.selection.ranges) for (let r = rg.r0; r <= Math.min(rg.r1, rg.r0 + 5000); r++) for (let c = rg.c0; c <= Math.min(rg.c1, rg.c0 + 200); c++) this.doc.setInput(r, c, text);
          });
          return;
        }
        finish(e.shiftKey ? -1 : 1, 0);
        return;
      case 'Tab':
        finish(0, e.shiftKey ? -1 : 1);
        return;
      case 'Escape':
        e.preventDefault();
        this.cancelEdit();
        return;
      case 'ArrowUp':
      case 'ArrowDown':
      case 'ArrowLeft':
      case 'ArrowRight': {
        if (this.edit.mode !== 'enter') return;
        if (this.isFormulaEdit() && this.canInsertRef()) {
          // Arrow keys point at cells while building a formula.
          e.preventDefault();
          const [dr, dc] = key === 'ArrowUp' ? [-1, 0] : key === 'ArrowDown' ? [1, 0] : key === 'ArrowLeft' ? [0, -1] : [0, 1];
          const base = this.pointRef ? this.pointedCell() : { r: this.edit.r, c: this.edit.c };
          const r = Math.max(0, base.r + dr);
          const c = Math.max(0, base.c + dc);
          this.insertPointRef(makeRange(r, c));
          this.scrollIntoView(r, c);
          return;
        }
        const [dr, dc] = key === 'ArrowUp' ? [-1, 0] : key === 'ArrowDown' ? [1, 0] : key === 'ArrowLeft' ? [0, -1] : [0, 1];
        finish(dr, dc);
        return;
      }
      case 'F2':
        e.preventDefault();
        this.edit.mode = this.edit.mode === 'enter' ? 'edit' : 'enter';
        this.emitEdit();
        return;
      case 'F4': {
        // Toggle absolute/relative for the reference at the caret.
        e.preventDefault();
        this.toggleAbsolute();
        return;
      }
      default:
        if (ctrl && key.toLowerCase() === 'z') {
          // Let the textarea undo typing.
          return;
        }
    }
  }

  private pointedCell(): { r: number; c: number } {
    const ref = this.pointRef ? this.edit.text.slice(this.pointRef.start, this.pointRef.end).replace(/^.*!/, '') : '';
    const m = /^\$?([A-Z]+)\$?(\d+)/.exec(ref);
    if (!m) return { r: this.edit.r, c: this.edit.c };
    let c = 0;
    for (const ch of m[1]) c = c * 26 + ch.charCodeAt(0) - 64;
    return { r: Number(m[2]) - 1, c: c - 1 };
  }

  private toggleAbsolute(): void {
    const text = this.edit.text;
    const caret = this.editor.selectionStart ?? this.edit.caret;
    const re = /(\$?)([A-Za-z]{1,3})(\$?)(\d+)/g;
    let m: RegExpExecArray | null;
    while ((m = re.exec(text))) {
      const start = m.index;
      const end = start + m[0].length;
      if (caret < start || caret > end) continue;
      const absC = m[1] === '$';
      const absR = m[3] === '$';
      const order: [boolean, boolean][] = [[false, false], [true, true], [false, true], [true, false]];
      const idx = order.findIndex(([a, b]) => a === absC && b === absR);
      const [nc, nr] = order[(idx + 1) % 4];
      const ref = `${nc ? '$' : ''}${m[2]}${nr ? '$' : ''}${m[4]}`;
      this.replaceEditText(text.slice(0, start) + ref + text.slice(end), start + ref.length);
      return;
    }
  }

  private fillFrom(direction: 'down' | 'right'): void {
    const range = this.doc.range;
    if (direction === 'down') {
      if (range.r0 === range.r1) {
        if (range.r0 === 0) return;
        this.doc.select(makeRange(range.r0 - 1, range.c0, range.r0 - 1, range.c1));
        this.doc.fill(makeRange(range.r0 - 1, range.c0, range.r1, range.c1), true);
      } else {
        this.doc.select(makeRange(range.r0, range.c0, range.r0, range.c1));
        this.doc.fill(range, true);
      }
    } else if (range.c0 === range.c1) {
      if (range.c0 === 0) return;
      this.doc.select(makeRange(range.r0, range.c0 - 1, range.r1, range.c0 - 1));
      this.doc.fill(makeRange(range.r0, range.c0 - 1, range.r1, range.c1), true);
    } else {
      this.doc.select(makeRange(range.r0, range.c0, range.r1, range.c0));
      this.doc.fill(range, true);
    }
  }

  /** AutoSum: sums the numbers above (or to the left of) the active cell. */
  autoSum(fn = 'SUM'): void {
    const { r, c } = this.doc.selection.active;
    const numeric = (rr: number, cc: number) => typeof this.doc.cell(rr, cc)?.v === 'number';
    let top = r - 1;
    while (top >= 0 && numeric(top, c)) top--;
    let ref: string | null = null;
    if (top < r - 1) ref = rangeName(makeRange(top + 1, c, r - 1, c));
    else {
      let left = c - 1;
      while (left >= 0 && numeric(r, left)) left--;
      if (left < c - 1) ref = rangeName(makeRange(r, left + 1, r, c - 1));
    }
    this.startEdit(`=${fn}(${ref ?? ''})`, 'enter');
    if (!ref) this.replaceEditText(`=${fn}()`, fn.length + 2);
  }

  // ---------------------------------------------------------------------------------------------
  // Clipboard (system clipboard events work without secure-context APIs)

  private onCopy(e: ClipboardEvent, cut: boolean): void {
    if (this.edit.active) return;
    if (this.selectedImage) {
      e.preventDefault();
      return;
    }
    const payload = cut ? this.doc.cut() : this.doc.copy();
    if (!payload) return;
    e.preventDefault();
    e.clipboardData?.setData('text/plain', payload.text);
    e.clipboardData?.setData('text/html', payload.html);
    this.markClip();
  }

  /** Dashed outline around the copied range. */
  markClip(): void {
    this.extraHighlights = [...this.extraHighlights.filter((h) => !h.dashed), { range: this.doc.range, sheet: this.sheet, color: '#1a73e8', dashed: true }];
    this.requestRender();
  }

  private onPaste(e: ClipboardEvent): void {
    if (this.edit.active) return;
    e.preventDefault();
    const data = e.clipboardData;
    const files = data ? [...data.files].filter((f) => f.type.startsWith('image/')) : [];
    if (files.length) {
      void this.insertImageFile(files[0]);
      return;
    }
    const html = data?.getData('text/html') || undefined;
    const text = data?.getData('text/plain') || undefined;
    // Nothing on the system clipboard (or no access to it): fall back to cells copied here.
    if (html === undefined && text === undefined) {
      if (!this.doc.hasClip()) return;
      this.doc.paste({});
    } else this.doc.paste({ text, html });
    this.extraHighlights = this.extraHighlights.filter((h) => !h.dashed);
    this.requestRender();
  }

  async insertImageFile(file: Blob, name?: string): Promise<void> {
    const data = new Uint8Array(await file.arrayBuffer());
    const url = URL.createObjectURL(file);
    const size = await new Promise<{ w: number; h: number }>((resolve) => {
      const img = new Image();
      img.onload = () => resolve({ w: img.naturalWidth, h: img.naturalHeight });
      img.onerror = () => resolve({ w: 200, h: 200 });
      img.src = url;
    });
    URL.revokeObjectURL(url);
    const vt = this.transform();
    const sel = this.doc.selection.active;
    const layout = vt.layout;
    const maxW = Math.max(120, (this.width - vt.mainLeft) / this.zoom * 0.6);
    const scale = Math.min(1, maxW / size.w, 400 / size.h);
    this.doc.insertImage(data, file.type || 'image/png', { x: layout.colX(sel.c), y: layout.rowY(sel.r), width: Math.round(size.w * scale), height: Math.round(size.h * scale) }, name);
  }

  // ---------------------------------------------------------------------------------------------

  /** Screen rectangle (client coordinates) of the active cell, for popups. */
  activeCellClientRect(): DOMRect {
    const vt = this.transform();
    const { r, c } = this.doc.selection.active;
    const m = this.sheet.mergeAt(r, c);
    const rect = m ? vt.rangeRect(m) : vt.cellRect(r, c);
    const root = this.root.getBoundingClientRect();
    return new DOMRect(root.left + rect.x, root.top + rect.y, rect.w, rect.h);
  }

  editorClientRect(): DOMRect {
    return this.editor.getBoundingClientRect();
  }

  /** Visible cell range (for tests and the find panel). */
  visibleRange(): Range {
    const { rows, cols } = this.transform().visibleMain();
    return makeRange(rows[0], cols[0], rows[1], cols[1]);
  }

  intersectsView(r: Range): boolean {
    return rangesIntersect(r, this.visibleRange());
  }
}

function range(a: number, b: number): number[] {
  const out: number[] = [];
  for (let i = Math.min(a, b); i <= Math.max(a, b); i++) out.push(i);
  return out;
}

function escapeHtml(s: string): string {
  return s.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
}

export type { SheetImage };
