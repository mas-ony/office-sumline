/**
 * Sheet geometry at 100% zoom: positions of rows and columns in pixels, hit testing, and
 * conversion between pixel rectangles and drawing anchors.
 */
import { MAX_COLS, MAX_ROWS } from './refs';
import type { AnchorPoint, Sheet, SheetImage } from './model/sheet';
import { EMU_PER_PX } from './xlsx/util';

interface Axis {
  version: number;
  /** Sorted indexes with a non-default size. */
  idx: number[];
  /** Size of each of those indexes. */
  size: number[];
  /** Sum of (size - default) for idx[0..i-1]. */
  before: number[];
  dflt: number;
}

const layouts = new WeakMap<Sheet, SheetLayout>();

export class SheetLayout {
  private rows: Axis | null = null;
  private cols: Axis | null = null;
  private rowsKey = '';
  private colsKey = '';

  constructor(readonly sheet: Sheet) {}

  static of(sheet: Sheet): SheetLayout {
    let l = layouts.get(sheet);
    if (!l) {
      l = new SheetLayout(sheet);
      layouts.set(sheet, l);
    }
    return l;
  }

  /** Must be called after row heights or column widths change. */
  invalidate(): void {
    this.rows = null;
    this.cols = null;
  }

  private rowAxis(): Axis {
    const key = `${this.sheet.layoutVersion}:${this.sheet.defaultRowHeightPt}`;
    if (this.rows && this.rowsKey === key) return this.rows;
    const dflt = this.sheet.rowHeightPx(-1);
    const idx = [...this.sheet.rows.keys()].sort((a, b) => a - b);
    const size = idx.map((r) => this.sheet.rowHeightPx(r));
    this.rows = buildAxis(idx, size, dflt);
    this.rowsKey = key;
    return this.rows;
  }

  private colAxis(): Axis {
    const key = `${this.sheet.layoutVersion}:${this.sheet.defaultColWidthPx}`;
    if (this.cols && this.colsKey === key) return this.cols;
    const dflt = this.sheet.defaultColWidthPx;
    const idx = [...this.sheet.cols.keys()].sort((a, b) => a - b);
    const size = idx.map((c) => this.sheet.colWidthPx(c));
    this.cols = buildAxis(idx, size, dflt);
    this.colsKey = key;
    return this.cols;
  }

  rowHeight(r: number): number {
    return this.sheet.rowHeightPx(r);
  }

  colWidth(c: number): number {
    return this.sheet.colWidthPx(c);
  }

  /** Top of row r. */
  rowY(r: number): number {
    return position(this.rowAxis(), r);
  }

  /** Left of column c. */
  colX(c: number): number {
    return position(this.colAxis(), c);
  }

  /** Row containing pixel y (clamped to the sheet). */
  rowAt(y: number): number {
    return Math.min(MAX_ROWS - 1, indexAt(this.rowAxis(), y));
  }

  colAt(x: number): number {
    return Math.min(MAX_COLS - 1, indexAt(this.colAxis(), x));
  }

  totalHeight(rows = MAX_ROWS): number {
    return this.rowY(rows);
  }

  totalWidth(cols = MAX_COLS): number {
    return this.colX(cols);
  }

  /** Pixel rectangle of an image (100% zoom). */
  imageRect(img: SheetImage): { x: number; y: number; w: number; h: number } {
    const pt = (p: AnchorPoint) => ({ x: this.colX(p.c) + p.dx / EMU_PER_PX, y: this.rowY(p.r) + p.dy / EMU_PER_PX });
    if (img.anchorType === 'absolute' && img.pos) {
      return { x: img.pos.x / EMU_PER_PX, y: img.pos.y / EMU_PER_PX, w: (img.ext?.cx ?? 0) / EMU_PER_PX, h: (img.ext?.cy ?? 0) / EMU_PER_PX };
    }
    const a = pt(img.from);
    if (img.anchorType === 'twoCell' && img.to) {
      const b = pt(img.to);
      return { x: a.x, y: a.y, w: Math.max(1, b.x - a.x), h: Math.max(1, b.y - a.y) };
    }
    return { x: a.x, y: a.y, w: (img.ext?.cx ?? 0) / EMU_PER_PX, h: (img.ext?.cy ?? 0) / EMU_PER_PX };
  }

  /** Pixel rectangle between two anchor points (charts and shapes). */
  anchorRect(from: AnchorPoint, to: AnchorPoint): { x: number; y: number; w: number; h: number } {
    const x0 = this.colX(from.c) + from.dx / EMU_PER_PX;
    const y0 = this.rowY(from.r) + from.dy / EMU_PER_PX;
    const x1 = this.colX(to.c) + to.dx / EMU_PER_PX;
    const y1 = this.rowY(to.r) + to.dy / EMU_PER_PX;
    return { x: x0, y: y0, w: Math.max(1, x1 - x0), h: Math.max(1, y1 - y0) };
  }

  /** Anchor point (cell + offset in EMU) of a pixel position. */
  anchorAt(x: number, y: number): AnchorPoint {
    const c = this.colAt(Math.max(0, x));
    const r = this.rowAt(Math.max(0, y));
    return { c, r, dx: Math.round((Math.max(0, x) - this.colX(c)) * EMU_PER_PX), dy: Math.round((Math.max(0, y) - this.rowY(r)) * EMU_PER_PX) };
  }
}

function buildAxis(idx: number[], size: number[], dflt: number): Axis {
  const before: number[] = new Array(idx.length + 1);
  before[0] = 0;
  for (let i = 0; i < idx.length; i++) before[i + 1] = before[i] + (size[i] - dflt);
  return { version: 0, idx, size, before, dflt };
}

/** Number of custom indexes strictly below i. */
function countBelow(axis: Axis, i: number): number {
  let lo = 0;
  let hi = axis.idx.length;
  while (lo < hi) {
    const mid = (lo + hi) >> 1;
    if (axis.idx[mid] < i) lo = mid + 1;
    else hi = mid;
  }
  return lo;
}

function position(axis: Axis, i: number): number {
  return i * axis.dflt + axis.before[countBelow(axis, i)];
}

function indexAt(axis: Axis, p: number): number {
  if (p <= 0) return 0;
  // Binary search on the monotonic position function.
  let lo = 0;
  let hi = Math.max(1, Math.ceil(p / Math.max(1, axis.dflt)) + 1);
  while (position(axis, hi) <= p) {
    lo = hi;
    hi *= 2;
    if (hi > 4 * 1048576) break;
  }
  while (lo < hi) {
    const mid = (lo + hi + 1) >> 1;
    if (position(axis, mid) <= p) lo = mid;
    else hi = mid - 1;
  }
  // Skip zero-size (hidden) lines.
  return lo;
}
