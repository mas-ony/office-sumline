/**
 * Screen <-> sheet coordinate transform with frozen panes. Sheet coordinates are pixels at 100%
 * zoom; screen coordinates are CSS pixels relative to the grid's top-left corner.
 */
import type { Sheet } from '../core/model/sheet';
import { SheetLayout } from '../core/layout';
import type { Range } from '../core/refs';

export interface ViewState {
  width: number;
  height: number;
  zoom: number;
  scrollX: number;
  scrollY: number;
  headerW: number;
  headerH: number;
}

export interface Rect {
  x: number;
  y: number;
  w: number;
  h: number;
}

export interface PaneSpec {
  /** Screen clip rectangle. */
  clip: Rect;
  rows: [number, number];
  cols: [number, number];
}

export class ViewTransform {
  readonly layout: SheetLayout;
  readonly fr: number;
  readonly fc: number;
  /** Frozen width/height in sheet pixels. */
  readonly FW: number;
  readonly FH: number;

  constructor(readonly sheet: Sheet, readonly v: ViewState) {
    this.layout = SheetLayout.of(sheet);
    this.fr = sheet.freeze.rows;
    this.fc = sheet.freeze.cols;
    this.FW = this.layout.colX(this.fc);
    this.FH = this.layout.rowY(this.fr);
  }

  get zoom(): number {
    return this.v.zoom;
  }

  /** Left edge of the scrollable area on screen. */
  get mainLeft(): number {
    return this.v.headerW + this.FW * this.v.zoom;
  }

  get mainTop(): number {
    return this.v.headerH + this.FH * this.v.zoom;
  }

  screenX(x: number): number {
    return this.v.headerW + x * this.v.zoom - (x >= this.FW ? this.v.scrollX : 0);
  }

  screenY(y: number): number {
    return this.v.headerH + y * this.v.zoom - (y >= this.FH ? this.v.scrollY : 0);
  }

  /** Sheet x at a screen x (frozen aware). */
  sheetX(sx: number): number {
    const local = sx - this.v.headerW;
    if (local < this.FW * this.v.zoom) return local / this.v.zoom;
    return (local + this.v.scrollX) / this.v.zoom;
  }

  sheetY(sy: number): number {
    const local = sy - this.v.headerH;
    if (local < this.FH * this.v.zoom) return local / this.v.zoom;
    return (local + this.v.scrollY) / this.v.zoom;
  }

  colAt(sx: number): number {
    return this.layout.colAt(Math.max(0, this.sheetX(sx)));
  }

  rowAt(sy: number): number {
    return this.layout.rowAt(Math.max(0, this.sheetY(sy)));
  }

  /** Screen rectangle of a cell range (may extend outside the view). */
  rangeRect(r: Range): Rect {
    const x0 = this.screenX(this.layout.colX(r.c0));
    const y0 = this.screenY(this.layout.rowY(r.r0));
    // The far edge belongs to the pane of the last line.
    const x1 = this.screenXEnd(r.c1);
    const y1 = this.screenYEnd(r.r1);
    return { x: x0, y: y0, w: x1 - x0, h: y1 - y0 };
  }

  private screenXEnd(c: number): number {
    const x = this.layout.colX(c);
    const end = this.layout.colX(c + 1);
    return this.v.headerW + end * this.v.zoom - (x >= this.FW ? this.v.scrollX : 0);
  }

  private screenYEnd(r: number): number {
    const y = this.layout.rowY(r);
    const end = this.layout.rowY(r + 1);
    return this.v.headerH + end * this.v.zoom - (y >= this.FH ? this.v.scrollY : 0);
  }

  cellRect(r: number, c: number): Rect {
    return this.rangeRect({ r0: r, c0: c, r1: r, c1: c });
  }

  /** First and last scrollable row/column visible on screen. */
  visibleMain(): { rows: [number, number]; cols: [number, number] } {
    const z = this.v.zoom;
    const firstRow = this.layout.rowAt(this.FH + this.v.scrollY / z);
    const lastRow = this.layout.rowAt(this.FH + (this.v.scrollY + this.v.height - this.mainTop) / z);
    const firstCol = this.layout.colAt(this.FW + this.v.scrollX / z);
    const lastCol = this.layout.colAt(this.FW + (this.v.scrollX + this.v.width - this.mainLeft) / z);
    return { rows: [Math.max(firstRow, this.fr), lastRow], cols: [Math.max(firstCol, this.fc), lastCol] };
  }

  /** Up to four panes (frozen corner, frozen rows, frozen columns, main) to draw. */
  panes(): PaneSpec[] {
    const { rows, cols } = this.visibleMain();
    const { width, height, headerW, headerH } = this.v;
    const out: PaneSpec[] = [];
    const ml = this.mainLeft;
    const mt = this.mainTop;
    if (this.fr && this.fc) out.push({ clip: { x: headerW, y: headerH, w: ml - headerW, h: mt - headerH }, rows: [0, this.fr - 1], cols: [0, this.fc - 1] });
    if (this.fr) out.push({ clip: { x: ml, y: headerH, w: width - ml, h: mt - headerH }, rows: [0, this.fr - 1], cols });
    if (this.fc) out.push({ clip: { x: headerW, y: mt, w: ml - headerW, h: height - mt }, rows, cols: [0, this.fc - 1] });
    out.push({ clip: { x: ml, y: mt, w: width - ml, h: height - mt }, rows, cols });
    return out;
  }

  /** Whether a row/column index is shown in the frozen part. */
  isFrozenRow(r: number): boolean {
    return r < this.fr;
  }

  isFrozenCol(c: number): boolean {
    return c < this.fc;
  }
}
