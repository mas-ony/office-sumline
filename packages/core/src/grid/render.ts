/** Canvas rendering of a sheet: cells, gridlines, borders, headers and selection. */
import { colName, MAX_COLS, MAX_ROWS, rangesIntersect, type Range } from '../core/refs';
import { CellError, type Scalar } from '../core/values';
import type { Cell, RichTextRun, Sheet } from '../core/model/sheet';
import type { BorderLine, CellStyle, FontStyle } from '../core/model/style';
import { resolveColor, type ColorSpec, type Theme } from '../core/model/color';
import { formatGeneral, formatValue } from '../core/format/display';
import type { SpreadsheetDocument, Selection } from '../core/document';
import { cssFont, lineHeightPx, textWidth, wrapText } from './fonts';
import type { PaneSpec, Rect, ViewTransform } from './viewport';
import { condFormatFor } from './condFormat';

export interface Highlight {
  range: Range;
  sheet?: Sheet;
  color: string;
  /** Dashed outline (copy marquee, fill preview). */
  dashed?: boolean;
  fill?: boolean;
}

export interface RenderState {
  selection: Selection;
  /** Hide the selection while the cell editor covers the active cell. */
  editing: boolean;
  mobile: boolean;
  highlights: Highlight[];
  /** Cells of the current search result. */
  searchHit?: { r: number; c: number };
  fillPreview?: Range;
  dpr: number;
  /** Columns whose header is being resized (feedback line). */
  resizeGuide?: { axis: 'row' | 'col'; pos: number };
}

export const COLORS = {
  grid: '#e3e3e3',
  header: '#f3f3f3',
  headerText: '#222222',
  headerBorder: '#cccccc',
  headerSelected: '#dcdcdc',
  headerSelectedText: '#000000',
  headerFullSelected: '#c8c8c8',
  headerMarker: '#444444',
  selection: '#000000',
  selectionFill: 'rgba(0,0,0,0.13)',
  text: '#000000',
  comment: '#d13438',
  threaded: '#c2185b',
  frozenLine: '#9a9a9a',
  filterButton: '#ffffff',
  filterActive: '#1a73e8',
};

/** The drop-down button of an auto filter header cell (screen coordinates of the grid). */
export function filterButtonRect(cell: Rect, zoom: number): Rect {
  const s = Math.round(Math.max(14, Math.min(20, 16 * zoom)));
  return { x: cell.x + cell.w - s - 2, y: cell.y + Math.max(1, (cell.h - s) / 2), w: s, h: s };
}

const colorCache = new WeakMap<object, string>();

function color(spec: ColorSpec | undefined, theme: Theme, fallback: string): string {
  if (!spec) return `#${fallback}`;
  const hit = colorCache.get(spec);
  if (hit) return hit;
  const c = `#${resolveColor(spec, theme, fallback)}`;
  colorCache.set(spec, c);
  return c;
}

/** Pattern fills approximated as a blend of the foreground on the background. */
const PATTERN_DENSITY: Record<string, number> = {
  solid: 1, darkGray: 0.75, mediumGray: 0.5, lightGray: 0.25, gray125: 0.125, gray0625: 0.0625,
  darkHorizontal: 0.5, darkVertical: 0.5, darkDown: 0.5, darkUp: 0.5, darkGrid: 0.5, darkTrellis: 0.75,
  lightHorizontal: 0.25, lightVertical: 0.25, lightDown: 0.25, lightUp: 0.25, lightGrid: 0.35, lightTrellis: 0.35,
};

function blend(fg: string, bg: string, a: number): string {
  const p = (h: string) => [parseInt(h.slice(1, 3), 16), parseInt(h.slice(3, 5), 16), parseInt(h.slice(5, 7), 16)];
  const [r1, g1, b1] = p(fg);
  const [r2, g2, b2] = p(bg);
  const m = (x: number, y: number) => Math.round(x * a + y * (1 - a)).toString(16).padStart(2, '0');
  return `#${m(r1, r2)}${m(g1, g2)}${m(b1, b2)}`;
}

export function fillColor(style: CellStyle, theme: Theme): string | null {
  const f = style.fill;
  if (!f) return null;
  if ('gradientXml' in f) return f.color ? color(f.color, theme, 'FFFFFF') : null;
  if (f.pattern === 'none') return null;
  const density = PATTERN_DENSITY[f.pattern] ?? 0.5;
  if (f.pattern === 'solid') return color(f.fg ?? f.bg, theme, 'FFFFFF');
  const fg = color(f.fg, theme, '000000');
  const bg = f.bg && !('auto' in f.bg) && !('indexed' in f.bg && f.bg.indexed === 64) ? color(f.bg, theme, 'FFFFFF') : '#FFFFFF';
  return blend(fg, bg, density);
}

interface CellText {
  text: string;
  color?: string;
  numeric: boolean;
}

const textCache = new Map<string, CellText>();

function cellText(v: Scalar, numFmt: string, date1904: boolean): CellText {
  if (v === null) return { text: '', numeric: false };
  const key = `${typeof v === 'object' ? (v as CellError).code : `${typeof v}:${String(v)}`}\u0000${numFmt}`;
  let t = textCache.get(key);
  if (!t) {
    const f = formatValue(v, numFmt, date1904);
    t = { text: f.text, color: f.color, numeric: f.numeric };
    if (textCache.size > 50000) textCache.clear();
    textCache.set(key, t);
  }
  return t;
}

export function clearTextCache(): void {
  textCache.clear();
}

function crisp(x: number, dpr: number, width = 1): number {
  const dev = Math.round(x * dpr);
  const lw = Math.max(1, Math.round(width * dpr));
  return (dev + (lw % 2 ? 0.5 : 0)) / dpr;
}

// -----------------------------------------------------------------------------------------------

interface DrawContext {
  ctx: CanvasRenderingContext2D;
  doc: SpreadsheetDocument;
  sheet: Sheet;
  vt: ViewTransform;
  theme: Theme;
  state: RenderState;
  z: number;
  dpr: number;
}

function isEmptyCell(sheet: Sheet, r: number, c: number): boolean {
  const cell = sheet.getCell(r, c);
  return !cell || (cell.v === null && cell.f === undefined) || cell.v === '';
}

export function render(ctx: CanvasRenderingContext2D, doc: SpreadsheetDocument, vt: ViewTransform, state: RenderState): void {
  const { width, height } = vt.v;
  const sheet = vt.sheet;
  const dc: DrawContext = { ctx, doc, sheet, vt, theme: doc.wb.theme, state, z: vt.zoom, dpr: state.dpr };
  ctx.save();
  ctx.fillStyle = '#ffffff';
  ctx.fillRect(0, 0, width, height);
  for (const pane of vt.panes()) drawPane(dc, pane);
  drawFrozenLines(dc);
  drawSelection(dc);
  drawFilterButtons(dc);
  drawHeaders(dc);
  ctx.restore();
}

function drawFilterButtons(dc: DrawContext): void {
  const { ctx, vt, sheet } = dc;
  const af = sheet.autoFilter;
  if (!af) return;
  const { rows, cols } = vt.visibleMain();
  const r = af.ref.r0;
  if (!(r < vt.fr || (r >= rows[0] && r <= rows[1]))) return;
  ctx.save();
  clipCells(dc);
  for (let c = af.ref.c0; c <= af.ref.c1; c++) {
    if (!(c < vt.fc || (c >= cols[0] && c <= cols[1]))) continue;
    const cell = vt.cellRect(r, c);
    if (cell.w < 18 || cell.h < 10) continue;
    const b = filterButtonRect(cell, vt.zoom);
    const active = af.columns.has(c);
    ctx.fillStyle = COLORS.filterButton;
    ctx.strokeStyle = active ? COLORS.filterActive : '#9e9e9e';
    ctx.lineWidth = 1;
    ctx.fillRect(b.x, b.y, b.w, b.h);
    ctx.strokeRect(crisp(b.x, dc.dpr), crisp(b.y, dc.dpr), b.w - 1, b.h - 1);
    ctx.fillStyle = active ? COLORS.filterActive : '#555555';
    ctx.beginPath();
    if (active) {
      // Funnel.
      const cx = b.x + b.w / 2;
      const top = b.y + b.h * 0.25;
      ctx.moveTo(b.x + b.w * 0.2, top);
      ctx.lineTo(b.x + b.w * 0.8, top);
      ctx.lineTo(cx + b.w * 0.08, b.y + b.h * 0.52);
      ctx.lineTo(cx + b.w * 0.08, b.y + b.h * 0.78);
      ctx.lineTo(cx - b.w * 0.08, b.y + b.h * 0.7);
      ctx.lineTo(cx - b.w * 0.08, b.y + b.h * 0.52);
    } else {
      const cx = b.x + b.w / 2;
      const cy = b.y + b.h / 2 + 1;
      ctx.moveTo(cx - b.w * 0.25, cy - b.h * 0.12);
      ctx.lineTo(cx + b.w * 0.25, cy - b.h * 0.12);
      ctx.lineTo(cx, cy + b.h * 0.16);
    }
    ctx.closePath();
    ctx.fill();
  }
  ctx.restore();
}

function drawPane(dc: DrawContext, pane: PaneSpec): void {
  const { ctx, vt, sheet, doc, z } = dc;
  const { clip } = pane;
  if (clip.w <= 0 || clip.h <= 0) return;
  const [r0, r1] = pane.rows;
  const [c0, c1] = pane.cols;
  if (r1 < r0 || c1 < c0) return;
  ctx.save();
  ctx.beginPath();
  ctx.rect(clip.x, clip.y, clip.w, clip.h);
  ctx.clip();

  const layout = vt.layout;
  const colLeft: number[] = [];
  const colW: number[] = [];
  for (let c = c0; c <= c1 + 1; c++) colLeft.push(vt.screenX(layout.colX(c)));
  for (let c = c0; c <= c1; c++) colW.push(colLeft[c - c0 + 1] - colLeft[c - c0]);
  // A column at the pane edge must still use the pane's scroll offset for its right edge.
  const rowTop: number[] = [];
  for (let r = r0; r <= r1 + 1; r++) rowTop.push(vt.screenY(layout.rowY(r)));
  const xOf = (c: number) => colLeft[c - c0];
  const yOf = (r: number) => rowTop[r - r0];

  // Gridlines.
  if (sheet.showGridLines) {
    ctx.strokeStyle = COLORS.grid;
    ctx.lineWidth = 1 / dc.dpr;
    ctx.beginPath();
    for (let r = r0; r <= r1 + 1; r++) {
      const y = crisp(yOf(r) - 0.5, dc.dpr, 1 / dc.dpr);
      ctx.moveTo(clip.x, y);
      ctx.lineTo(clip.x + clip.w, y);
    }
    for (let c = c0; c <= c1 + 1; c++) {
      const x = crisp(xOf(c) - 0.5, dc.dpr, 1 / dc.dpr);
      ctx.moveTo(x, clip.y);
      ctx.lineTo(x, clip.y + clip.h);
    }
    ctx.stroke();
  }

  const visible: Range = { r0, c0, r1, c1 };
  const merges = sheet.merges.filter((m) => rangesIntersect(m, visible));
  const covered = (r: number, c: number) => merges.some((m) => r >= m.r0 && r <= m.r1 && c >= m.c0 && c <= m.c1);

  // Fills (and blank merged areas hide inner gridlines).
  const styles = doc.wb.styles;
  const cf = condFormatFor(doc, sheet);
  for (let r = r0; r <= r1; r++) {
    if (rowTop[r - r0 + 1] - rowTop[r - r0] <= 0) continue;
    for (let c = c0; c <= c1; c++) {
      if (merges.length && covered(r, c)) continue;
      const style = styles.get(sheet.styleAt(r, c));
      let bg = fillColor(style, dc.theme);
      const rule = cf?.(r, c);
      if (rule?.fill) bg = rule.fill;
      if (bg) {
        ctx.fillStyle = bg;
        ctx.fillRect(xOf(c) - 1 / dc.dpr, yOf(r) - 1 / dc.dpr, colW[c - c0] + 1 / dc.dpr, rowTop[r - r0 + 1] - rowTop[r - r0] + 1 / dc.dpr);
      }
      if (rule?.bar) drawDataBar(dc, { x: xOf(c), y: yOf(r), w: colW[c - c0], h: rowTop[r - r0 + 1] - rowTop[r - r0] }, rule.bar);
    }
  }
  for (const m of merges) {
    const rect = vt.rangeRect(m);
    const style = styles.get(sheet.styleAt(m.r0, m.c0));
    ctx.fillStyle = fillColor(style, dc.theme) ?? '#ffffff';
    ctx.fillRect(rect.x - 1 / dc.dpr, rect.y - 1 / dc.dpr, rect.w + 1 / dc.dpr, rect.h + 1 / dc.dpr);
  }

  // Text.
  const overflowFrom = Math.max(0, c0 - 30);
  for (let r = r0; r <= r1; r++) {
    const h = rowTop[r - r0 + 1] - rowTop[r - r0];
    if (h <= 0) continue;
    const row = sheet.cells.get(r);
    if (!row) continue;
    for (let c = overflowFrom; c <= c1; c++) {
      const cell = row.get(c);
      if (!cell || cell.v === null || cell.v === '') continue;
      if (merges.length && covered(r, c)) continue;
      if (sheet.mergeAt(r, c)) continue;
      const x = c >= c0 ? xOf(c) : vt.screenX(layout.colX(c));
      const w = c >= c0 ? colW[c - c0] : layout.colWidth(c) * z;
      if (w <= 0) continue;
      drawCellText(dc, r, c, cell, { x, y: yOf(r), w, h }, c < c0, c0, c1);
    }
  }
  for (const m of merges) {
    const cell = sheet.getCell(m.r0, m.c0);
    if (!cell || cell.v === null) continue;
    const rect = vt.rangeRect(m);
    drawCellText(dc, m.r0, m.c0, cell, rect, false, c0, c1, true);
  }

  // Borders.
  for (let r = Math.max(0, r0 - 1); r <= r1; r++) {
    for (let c = Math.max(0, c0 - 1); c <= c1; c++) {
      const style = styles.get(sheet.styleAt(r, c));
      const b = style.border;
      if (!b) continue;
      const rect = vt.cellRect(r, c);
      const m = merges.length ? sheet.mergeAt(r, c) : undefined;
      if (b.top && (!m || r === m.r0)) drawBorderLine(dc, b.top, rect.x, rect.y, rect.x + rect.w, rect.y);
      if (b.bottom && (!m || r === m.r1)) drawBorderLine(dc, b.bottom, rect.x, rect.y + rect.h, rect.x + rect.w, rect.y + rect.h);
      if (b.left && (!m || c === m.c0)) drawBorderLine(dc, b.left, rect.x, rect.y, rect.x, rect.y + rect.h);
      if (b.right && (!m || c === m.c1)) drawBorderLine(dc, b.right, rect.x + rect.w, rect.y, rect.x + rect.w, rect.y + rect.h);
      if (b.diagonal && (b.diagonalDown || b.diagonalUp)) {
        if (b.diagonalDown) drawBorderLine(dc, b.diagonal, rect.x, rect.y, rect.x + rect.w, rect.y + rect.h);
        if (b.diagonalUp) drawBorderLine(dc, b.diagonal, rect.x, rect.y + rect.h, rect.x + rect.w, rect.y);
      }
    }
  }

  // Comment indicators.
  for (const t of sheet.comments) {
    if (t.r < r0 || t.r > r1 || t.c < c0 || t.c > c1) continue;
    const m = sheet.mergeAt(t.r, t.c);
    const rect = m ? vt.rangeRect(m) : vt.cellRect(t.r, t.c);
    const s = Math.max(5, 6 * z);
    ctx.fillStyle = t.kind === 'threaded' ? COLORS.threaded : COLORS.comment;
    ctx.beginPath();
    ctx.moveTo(rect.x + rect.w - s, rect.y);
    ctx.lineTo(rect.x + rect.w, rect.y);
    ctx.lineTo(rect.x + rect.w, rect.y + s);
    ctx.closePath();
    ctx.fill();
  }
  ctx.restore();
}

function drawDataBar(dc: DrawContext, rect: Rect, bar: { ratio: number; color: string }): void {
  const { ctx } = dc;
  const pad = 2;
  const w = Math.max(0, (rect.w - 2 * pad) * Math.min(1, Math.max(0, bar.ratio)));
  const g = ctx.createLinearGradient(rect.x, 0, rect.x + w, 0);
  g.addColorStop(0, bar.color);
  g.addColorStop(1, `${bar.color}22`);
  ctx.fillStyle = g;
  ctx.fillRect(rect.x + pad, rect.y + pad, w, rect.h - 2 * pad);
}

const DASHES: Record<string, number[]> = {
  dashed: [3, 1],
  dotted: [1, 1],
  hair: [1, 1],
  mediumDashed: [6, 2],
  dashDot: [6, 2, 2, 2],
  mediumDashDot: [6, 2, 2, 2],
  dashDotDot: [6, 2, 2, 2, 2, 2],
  mediumDashDotDot: [6, 2, 2, 2, 2, 2],
  slantDashDot: [8, 2, 2, 2],
};

function drawBorderLine(dc: DrawContext, line: BorderLine, x1: number, y1: number, x2: number, y2: number): void {
  const { ctx, dpr } = dc;
  const width = line.style === 'medium' || line.style.startsWith('medium') || line.style === 'slantDashDot' ? 2 : line.style === 'thick' ? 3 : 1;
  ctx.save();
  ctx.strokeStyle = color(line.color, dc.theme, '000000');
  ctx.lineWidth = width;
  const dash = DASHES[line.style];
  ctx.setLineDash(dash ? dash.map((d) => d * Math.max(1, width / 1.5)) : []);
  const horizontal = y1 === y2;
  if (line.style === 'double') {
    ctx.lineWidth = 1;
    ctx.beginPath();
    for (const off of [-1, 1]) {
      if (horizontal) {
        const y = crisp(y1 - 0.5 + off, dpr);
        ctx.moveTo(x1 - 1, y);
        ctx.lineTo(x2 + 1, y);
      } else {
        const x = crisp(x1 - 0.5 + off, dpr);
        ctx.moveTo(x, y1 - 1);
        ctx.lineTo(x, y2 + 1);
      }
    }
    ctx.stroke();
    ctx.restore();
    return;
  }
  ctx.beginPath();
  if (horizontal) {
    const y = crisp(y1 - 0.5, dpr, width);
    ctx.moveTo(x1 - (width > 1 ? 1 : 0.5), y);
    ctx.lineTo(x2 + (width > 1 ? 1 : 0), y);
  } else if (x1 === x2) {
    const x = crisp(x1 - 0.5, dpr, width);
    ctx.moveTo(x, y1 - (width > 1 ? 1 : 0.5));
    ctx.lineTo(x, y2 + (width > 1 ? 1 : 0));
  } else {
    ctx.moveTo(x1, y1);
    ctx.lineTo(x2, y2);
  }
  ctx.stroke();
  ctx.restore();
}

function effectiveFont(base: FontStyle, run?: Partial<FontStyle>): FontStyle {
  return run ? { ...base, ...run } : base;
}

function drawCellText(
  dc: DrawContext,
  r: number,
  c: number,
  cell: Cell,
  rect: Rect,
  offscreen: boolean,
  c0: number,
  c1: number,
  merged = false,
): void {
  const { ctx, doc, sheet, z } = dc;
  const style = doc.wb.styles.get(cell.s ?? sheet.styleAt(r, c));
  const t = cellText(cell.v, style.numFmt, doc.wb.date1904);
  if (!t.text) return;
  const align = style.align ?? {};
  const v = cell.v;
  const cf = condFormatFor(doc, sheet)?.(r, c);
  let fontStyle = style.font;
  if (cf?.font) fontStyle = { ...fontStyle, ...cf.font };
  let textColor = t.color ?? cf?.color ?? color(fontStyle.color, dc.theme, '000000');
  if (fontStyle.color && 'auto' in fontStyle.color) textColor = '#000000';
  const font = cssFont(fontStyle, z);
  const pad = 3 * z;
  let h = align.horizontal ?? 'general';
  if (h === 'general') h = typeof v === 'number' ? 'right' : typeof v === 'boolean' || v instanceof CellError ? 'center' : 'left';
  if (h === 'centerContinuous') h = 'center';
  const indent = (align.indent ?? 0) * 9 * z;
  const wrap = !!align.wrap || h === 'justify' || h === 'distributed';
  const rotation = align.rotation ?? 0;
  const lh = lineHeightPx(fontStyle.size, z);
  const sizePx = (fontStyle.size * 96) / 72 * z;

  ctx.font = font;
  ctx.fillStyle = textColor;
  ctx.textBaseline = 'alphabetic';

  if (rotation && rotation !== 0) {
    if (offscreen) return;
    drawRotated(dc, t.text, rect, rotation, font, sizePx, align.vertical ?? 'bottom', h);
    return;
  }

  let text = t.text;
  // Accounting formats: the part before the fill marker is left aligned.
  if (wrap) {
    if (offscreen) return;
    const lines = wrapText(text, font, Math.max(1, rect.w - 2 * pad - indent));
    const blockH = lines.length * lh;
    let y = blockTop(rect, blockH, align.vertical ?? 'bottom', pad / 2) + lh * 0.8;
    ctx.save();
    ctx.beginPath();
    ctx.rect(rect.x, rect.y, rect.w, rect.h);
    ctx.clip();
    for (const line of lines) {
      const w = textWidth(line, font);
      const x = h === 'right' ? rect.x + rect.w - pad - w - indent : h === 'center' || h === 'distributed' ? rect.x + (rect.w - w) / 2 : rect.x + pad + indent;
      ctx.fillText(line, x, y);
      decorate(ctx, fontStyle, x, y, w, sizePx, textColor);
      y += lh;
    }
    ctx.restore();
    return;
  }

  // Single line (multi-line text without wrap shows only explicit line breaks).
  const lines = text.split('\n');
  const multi = lines.length > 1;
  if (!multi) text = lines[0];
  let width = multi ? Math.max(...lines.map((l) => textWidth(l, font))) : textWidth(text, font);
  const avail = rect.w - 2 * pad - indent;

  if (t.numeric && width > avail && !merged) {
    // Numbers that do not fit: General shrinks precision, other formats show ###.
    if (style.numFmt === 'General' && typeof v === 'number') {
      for (let digits = 9; digits >= 1; digits--) {
        const s = shorten(v, digits);
        if (textWidth(s, font) <= avail) {
          text = s;
          width = textWidth(s, font);
          break;
        }
      }
    }
    if (width > avail) {
      const hashW = textWidth('#', font);
      text = '#'.repeat(Math.max(1, Math.floor(avail / hashW)));
      width = textWidth(text, font);
    }
  }

  // Overflow into empty neighbours (text only, not clipped, not numbers).
  let clipX0 = rect.x;
  let clipX1 = rect.x + rect.w;
  const clip = !!align.clip || h === 'fill';
  if (!t.numeric && !clip && !merged && width > avail) {
    const extendRight = h === 'left' || h === 'center';
    const extendLeft = h === 'right' || h === 'center';
    const need = width - avail;
    if (extendRight) {
      let cc = c + 1;
      let gained = 0;
      while (gained < (h === 'center' ? need / 2 : need) && cc < MAX_COLS && isEmptyCell(sheet, r, cc) && !sheet.mergeAt(r, cc) && cc <= c1 + 30) {
        gained += dc.vt.layout.colWidth(cc) * z;
        cc++;
      }
      clipX1 = rect.x + rect.w + gained;
    }
    if (extendLeft) {
      let cc = c - 1;
      let gained = 0;
      while (gained < (h === 'center' ? need / 2 : need) && cc >= 0 && isEmptyCell(sheet, r, cc) && !sheet.mergeAt(r, cc)) {
        gained += dc.vt.layout.colWidth(cc) * z;
        cc--;
      }
      clipX0 = rect.x - gained;
    }
  }
  if (offscreen && clipX1 <= dc.vt.screenX(dc.vt.layout.colX(c0))) return;
  void c1;

  let x: number;
  if (h === 'right') x = rect.x + rect.w - pad - width - indent;
  else if (h === 'center' || h === 'distributed') x = rect.x + (rect.w - width) / 2;
  else x = rect.x + pad + indent;
  if (h === 'fill' && width > 0 && !multi) {
    const reps = Math.max(1, Math.floor(avail / width));
    text = text.repeat(reps);
    width = textWidth(text, font);
  }
  const blockH = (multi ? lines.length : 1) * lh;
  let y = blockTop(rect, blockH, align.vertical ?? 'bottom', pad / 2) + lh * 0.8;

  ctx.save();
  ctx.beginPath();
  ctx.rect(clipX0, rect.y, clipX1 - clipX0, rect.h);
  ctx.clip();
  if (cell.rt && cell.rt.length && !multi) {
    drawRuns(dc, cell.rt, style.font, x, y, z, textColor);
  } else if (multi) {
    for (const line of lines) {
      const w = textWidth(line, font);
      const lx = h === 'right' ? rect.x + rect.w - pad - w - indent : h === 'center' ? rect.x + (rect.w - w) / 2 : rect.x + pad + indent;
      ctx.fillText(line, lx, y);
      decorate(ctx, fontStyle, lx, y, w, sizePx, textColor);
      y += lh;
    }
  } else if (text.includes('\u0001')) {
    const [left, right] = text.split('\u0001');
    ctx.fillText(left, rect.x + pad, y);
    const rw = textWidth(right, font);
    ctx.fillText(right, rect.x + rect.w - pad - rw, y);
  } else {
    ctx.fillText(text, x, y);
    decorate(ctx, fontStyle, x, y, width, sizePx, textColor);
  }
  ctx.restore();
}

function shorten(v: number, digits: number): string {
  const abs = Math.abs(v);
  if (abs >= 1e11 || (abs < 1e-4 && abs > 0)) {
    const s = v.toExponential(Math.max(0, digits - 4)).replace(/\.?0+e/, 'e');
    const [m, e] = s.split('e');
    return `${m}E${Number(e) < 0 ? '-' : '+'}${String(Math.abs(Number(e))).padStart(2, '0')}`;
  }
  const intDigits = Math.max(1, Math.floor(Math.log10(Math.max(1, abs))) + 1);
  if (intDigits > digits) {
    const s = v.toExponential(Math.max(0, digits - 5)).replace(/\.?0+e/, 'e');
    const [m, e] = s.split('e');
    return `${m}E+${String(Number(e)).padStart(2, '0')}`;
  }
  const dec = Math.max(0, digits - intDigits);
  const out = Number(v.toFixed(dec));
  return formatGeneral(out);
}

function blockTop(rect: Rect, blockH: number, valign: string, pad: number): number {
  if (valign === 'top') return rect.y + pad;
  if (valign === 'center' || valign === 'justify' || valign === 'distributed') return rect.y + (rect.h - blockH) / 2;
  return rect.y + rect.h - blockH - pad;
}

function decorate(ctx: CanvasRenderingContext2D, font: FontStyle, x: number, y: number, w: number, sizePx: number, colorCss: string): void {
  if (!font.underline && !font.strike) return;
  ctx.save();
  ctx.strokeStyle = colorCss;
  ctx.lineWidth = Math.max(1, sizePx / 14);
  ctx.beginPath();
  if (font.underline) {
    const uy = y + Math.max(1.5, sizePx * 0.12);
    ctx.moveTo(x, uy);
    ctx.lineTo(x + w, uy);
    if (font.underline === 'double' || font.underline === 'doubleAccounting') {
      ctx.moveTo(x, uy + 2);
      ctx.lineTo(x + w, uy + 2);
    }
  }
  if (font.strike) {
    const sy = y - sizePx * 0.3;
    ctx.moveTo(x, sy);
    ctx.lineTo(x + w, sy);
  }
  ctx.stroke();
  ctx.restore();
}

function drawRuns(dc: DrawContext, runs: RichTextRun[], base: FontStyle, x: number, y: number, z: number, fallbackColor: string): void {
  const { ctx } = dc;
  let cx = x;
  for (const run of runs) {
    const f = effectiveFont(base, run.font);
    const fontCss = cssFont({ ...f, size: f.vertAlign ? f.size * 0.7 : f.size }, z);
    ctx.font = fontCss;
    ctx.fillStyle = run.font?.color ? color(run.font.color, dc.theme, '000000') : fallbackColor;
    const dy = f.vertAlign === 'superscript' ? -f.size * 0.35 * z : f.vertAlign === 'subscript' ? f.size * 0.15 * z : 0;
    ctx.fillText(run.text, cx, y + dy);
    const w = textWidth(run.text, fontCss);
    decorate(ctx, f, cx, y + dy, w, (f.size * 96) / 72 * z, ctx.fillStyle as string);
    cx += w;
  }
}

function drawRotated(dc: DrawContext, text: string, rect: Rect, rotation: number, font: string, sizePx: number, valign: string, h: string): void {
  const { ctx } = dc;
  ctx.save();
  ctx.beginPath();
  ctx.rect(rect.x, rect.y, rect.w, rect.h);
  ctx.clip();
  if (rotation === 255) {
    // Stacked letters.
    const lh = sizePx * 1.15;
    const chars = [...text];
    const blockH = chars.length * lh;
    let y = blockTop(rect, blockH, valign, 2) + sizePx;
    for (const ch of chars) {
      const w = textWidth(ch, font);
      ctx.fillText(ch, rect.x + (rect.w - w) / 2, y);
      y += lh;
    }
    ctx.restore();
    return;
  }
  const angle = rotation <= 90 ? -rotation : rotation - 90;
  const rad = (angle * Math.PI) / 180;
  const w = textWidth(text, font);
  const cx = h === 'left' ? rect.x + 4 + Math.abs(Math.sin(rad)) * sizePx : h === 'right' ? rect.x + rect.w - 4 : rect.x + rect.w / 2;
  const cy = valign === 'top' ? rect.y + 4 : valign === 'center' ? rect.y + rect.h / 2 : rect.y + rect.h - 4;
  ctx.translate(cx, cy);
  ctx.rotate(rad);
  const ox = angle < 0 ? 0 : -w;
  ctx.fillText(text, h === 'center' ? -w / 2 : ox, sizePx * 0.35);
  ctx.restore();
}

function drawFrozenLines(dc: DrawContext): void {
  const { ctx, vt } = dc;
  ctx.strokeStyle = COLORS.frozenLine;
  ctx.lineWidth = 1;
  ctx.beginPath();
  if (vt.fr) {
    const y = crisp(vt.mainTop - 0.5, dc.dpr);
    ctx.moveTo(vt.v.headerW, y);
    ctx.lineTo(vt.v.width, y);
  }
  if (vt.fc) {
    const x = crisp(vt.mainLeft - 0.5, dc.dpr);
    ctx.moveTo(x, vt.v.headerH);
    ctx.lineTo(x, vt.v.height);
  }
  ctx.stroke();
}

/** Clips drawing to the cell area (not headers). */
function clipCells(dc: DrawContext): void {
  const { ctx, vt } = dc;
  ctx.beginPath();
  ctx.rect(vt.v.headerW, vt.v.headerH, vt.v.width - vt.v.headerW, vt.v.height - vt.v.headerH);
  ctx.clip();
}

/** Screen rectangle of a range clipped per frozen pane (ranges crossing the freeze line). */
function visibleRect(dc: DrawContext, range: Range): Rect {
  const rect = dc.vt.rangeRect(range);
  return rect;
}

function drawSelection(dc: DrawContext): void {
  const { ctx, state, vt, sheet } = dc;
  ctx.save();
  clipCells(dc);
  const sel = state.selection;
  const primary = sel.ranges[sel.ranges.length - 1];
  const activeMerge = sheet.mergeAt(sel.active.r, sel.active.c);
  const activeRange = activeMerge ?? { r0: sel.active.r, c0: sel.active.c, r1: sel.active.r, c1: sel.active.c };

  // Formula reference and other highlights.
  for (const hl of state.highlights) {
    if (hl.sheet && hl.sheet !== sheet) continue;
    const rect = visibleRect(dc, clampRange(hl.range));
    if (hl.fill) {
      ctx.fillStyle = `${hl.color}22`;
      ctx.fillRect(rect.x, rect.y, rect.w, rect.h);
    }
    ctx.strokeStyle = hl.color;
    ctx.lineWidth = hl.dashed ? 1.5 : 2;
    ctx.setLineDash(hl.dashed ? [5, 3] : []);
    ctx.strokeRect(rect.x + 0.5, rect.y + 0.5, rect.w - 1, rect.h - 1);
  }
  ctx.setLineDash([]);

  if (!state.editing || state.mobile) {
    // Selected cells are shaded (the active cell too) and outlined.
    ctx.fillStyle = COLORS.selectionFill;
    for (const range of sel.ranges) {
      const rect = visibleRect(dc, clampRange(range));
      ctx.fillRect(rect.x, rect.y, rect.w, rect.h);
    }
    const ar = visibleRect(dc, activeRange);
    const pr = visibleRect(dc, clampRange(primary));
    ctx.strokeStyle = COLORS.selection;
    ctx.lineWidth = 1;
    ctx.strokeRect(crisp(pr.x - 0.5, dc.dpr), crisp(pr.y - 0.5, dc.dpr), Math.round(pr.w), Math.round(pr.h));
    if (!(primary.r0 === activeRange.r0 && primary.c0 === activeRange.c0 && primary.r1 === activeRange.r1 && primary.c1 === activeRange.c1)) {
      ctx.strokeStyle = 'rgba(0,0,0,0.55)';
      ctx.setLineDash([2, 2]);
      ctx.strokeRect(ar.x + 0.5, ar.y + 0.5, ar.w - 1, ar.h - 1);
      ctx.setLineDash([]);
    }
    if (!state.mobile && !dc.doc.isReadOnly()) {
      // Fill handle: a small square on the bottom-right corner.
      const s = 5;
      ctx.fillStyle = COLORS.selection;
      ctx.fillRect(Math.round(pr.x + pr.w - s / 2), Math.round(pr.y + pr.h - s / 2), s, s);
      ctx.strokeStyle = '#ffffff';
      ctx.lineWidth = 1;
      ctx.strokeRect(Math.round(pr.x + pr.w - s / 2) - 0.5, Math.round(pr.y + pr.h - s / 2) - 0.5, s + 1, s + 1);
    }
  }
  if (state.fillPreview) {
    const rect = visibleRect(dc, state.fillPreview);
    ctx.strokeStyle = '#5f6368';
    ctx.lineWidth = 1;
    ctx.setLineDash([4, 3]);
    ctx.strokeRect(rect.x + 0.5, rect.y + 0.5, rect.w - 1, rect.h - 1);
    ctx.setLineDash([]);
  }
  if (state.searchHit) {
    const rect = vt.cellRect(state.searchHit.r, state.searchHit.c);
    ctx.strokeStyle = '#f9ab00';
    ctx.lineWidth = 2;
    ctx.strokeRect(rect.x + 1, rect.y + 1, rect.w - 2, rect.h - 2);
  }
  // Spill range outline for the active dynamic array.
  const anchor = dc.doc.anchorOf(sel.active.r, sel.active.c) ?? (dc.doc.cell(sel.active.r, sel.active.c)?.f !== undefined ? sel.active : null);
  if (anchor) {
    const spill = dc.doc.engine.spillOf(sheet, anchor.r, anchor.c);
    if (spill && (spill.r0 !== spill.r1 || spill.c0 !== spill.c1)) {
      const rect = visibleRect(dc, spill);
      ctx.strokeStyle = '#4a86e8';
      ctx.lineWidth = 1;
      ctx.setLineDash([3, 2]);
      ctx.strokeRect(rect.x + 0.5, rect.y + 0.5, rect.w - 1, rect.h - 1);
      ctx.setLineDash([]);
    }
  }
  if (state.resizeGuide) {
    ctx.strokeStyle = COLORS.selection;
    ctx.lineWidth = 1;
    ctx.setLineDash([4, 2]);
    ctx.beginPath();
    if (state.resizeGuide.axis === 'col') {
      ctx.moveTo(state.resizeGuide.pos, 0);
      ctx.lineTo(state.resizeGuide.pos, vt.v.height);
    } else {
      ctx.moveTo(0, state.resizeGuide.pos);
      ctx.lineTo(vt.v.width, state.resizeGuide.pos);
    }
    ctx.stroke();
    ctx.setLineDash([]);
  }
  ctx.restore();
}

function clampRange(r: Range): Range {
  return { r0: r.r0, c0: r.c0, r1: Math.min(r.r1, MAX_ROWS - 1), c1: Math.min(r.c1, MAX_COLS - 1) };
}

/** Shortens text with an ellipsis to fit a width. */
function fitText(ctx: CanvasRenderingContext2D, text: string, width: number): string {
  if (ctx.measureText(text).width <= width) return text;
  let lo = 0;
  let hi = text.length;
  while (lo < hi) {
    const mid = (lo + hi + 1) >> 1;
    if (ctx.measureText(`${text.slice(0, mid)}…`).width <= width) lo = mid;
    else hi = mid - 1;
  }
  return lo ? `${text.slice(0, lo)}…` : '';
}

function drawHeaders(dc: DrawContext): void {
  const { ctx, vt, state, sheet } = dc;
  const { width, height, headerW, headerH } = vt.v;
  const sel = state.selection;
  const font = `${Math.round(13 * Math.min(1.4, Math.max(0.8, vt.zoom)))}px Arial, "Arimo", "Liberation Sans", sans-serif`;
  ctx.save();
  ctx.font = font;
  ctx.textBaseline = 'middle';
  ctx.textAlign = 'center';

  const selRows = (r: number) => sel.ranges.some((x) => r >= x.r0 && r <= x.r1);
  const selCols = (c: number) => sel.ranges.some((x) => c >= x.c0 && c <= x.c1);
  const fullRow = (r: number) => sel.ranges.some((x) => x.c0 === 0 && x.c1 === MAX_COLS - 1 && r >= x.r0 && r <= x.r1);
  const fullCol = (c: number) => sel.ranges.some((x) => x.r0 === 0 && x.r1 === MAX_ROWS - 1 && c >= x.c0 && c <= x.c1);

  // Column headers.
  ctx.fillStyle = COLORS.header;
  ctx.fillRect(0, 0, width, headerH);
  const { rows, cols } = vt.visibleMain();
  const colRanges: [number, number][] = [];
  if (vt.fc) colRanges.push([0, vt.fc - 1]);
  colRanges.push(cols);
  for (const [a, b] of colRanges) {
    ctx.save();
    const clipX = a === 0 && vt.fc && b === vt.fc - 1 ? headerW : vt.mainLeft;
    const clipW = a === 0 && vt.fc && b === vt.fc - 1 ? vt.mainLeft - headerW : width - vt.mainLeft;
    ctx.beginPath();
    ctx.rect(clipX, 0, clipW, headerH);
    ctx.clip();
    for (let c = a; c <= b; c++) {
      const rect = vt.cellRect(0, c);
      if (rect.w <= 0) continue;
      const selected = selCols(c);
      if (selected) {
        ctx.fillStyle = fullCol(c) ? COLORS.headerFullSelected : COLORS.headerSelected;
        ctx.fillRect(rect.x, 0, rect.w, headerH);
      }
      ctx.fillStyle = selected ? COLORS.headerSelectedText : COLORS.headerText;
      if (rect.w > 12) {
        const title = sheet.colTitles?.get(c) ?? colName(c);
        ctx.fillText(fitText(ctx, title, rect.w - 6), rect.x + rect.w / 2, headerH / 2 + 0.5);
      }
      ctx.strokeStyle = COLORS.headerBorder;
      ctx.lineWidth = 1;
      ctx.beginPath();
      const x = crisp(rect.x + rect.w - 0.5, dc.dpr);
      ctx.moveTo(x, 0);
      ctx.lineTo(x, headerH);
      ctx.stroke();
      if (selected) {
        ctx.fillStyle = COLORS.headerMarker;
        ctx.fillRect(rect.x, headerH - 2, rect.w, 2);
      }
    }
    ctx.restore();
  }

  // Row headers.
  ctx.fillStyle = COLORS.header;
  ctx.fillRect(0, headerH, headerW, height - headerH);
  const rowRanges: [number, number][] = [];
  if (vt.fr) rowRanges.push([0, vt.fr - 1]);
  rowRanges.push(rows);
  for (const [a, b] of rowRanges) {
    ctx.save();
    const frozen = a === 0 && vt.fr && b === vt.fr - 1;
    ctx.beginPath();
    ctx.rect(0, frozen ? headerH : vt.mainTop, headerW, frozen ? vt.mainTop - headerH : height - vt.mainTop);
    ctx.clip();
    for (let r = a; r <= b; r++) {
      const rect = vt.cellRect(r, 0);
      if (rect.h <= 0) continue;
      const selected = selRows(r);
      if (selected) {
        ctx.fillStyle = fullRow(r) ? COLORS.headerFullSelected : COLORS.headerSelected;
        ctx.fillRect(0, rect.y, headerW, rect.h);
      }
      ctx.fillStyle = selected ? COLORS.headerSelectedText : COLORS.headerText;
      if (rect.h > 8) ctx.fillText(String(r + 1), headerW / 2, rect.y + rect.h / 2 + 0.5);
      ctx.strokeStyle = COLORS.headerBorder;
      ctx.lineWidth = 1;
      ctx.beginPath();
      const y = crisp(rect.y + rect.h - 0.5, dc.dpr);
      ctx.moveTo(0, y);
      ctx.lineTo(headerW, y);
      ctx.stroke();
      if (selected) {
        ctx.fillStyle = COLORS.headerMarker;
        ctx.fillRect(headerW - 2, rect.y, 2, rect.h);
      }
    }
    ctx.restore();
  }

  // Corner and header separators.
  ctx.fillStyle = COLORS.header;
  ctx.fillRect(0, 0, headerW, headerH);
  ctx.fillStyle = '#b6bac0';
  ctx.beginPath();
  ctx.moveTo(headerW - 4, headerH - 4);
  ctx.lineTo(headerW - 4, headerH - 12);
  ctx.lineTo(headerW - 12, headerH - 4);
  ctx.closePath();
  ctx.fill();
  ctx.strokeStyle = COLORS.headerBorder;
  ctx.lineWidth = 1;
  ctx.beginPath();
  ctx.moveTo(0, crisp(headerH - 0.5, dc.dpr));
  ctx.lineTo(width, crisp(headerH - 0.5, dc.dpr));
  ctx.moveTo(crisp(headerW - 0.5, dc.dpr), 0);
  ctx.lineTo(crisp(headerW - 0.5, dc.dpr), height);
  ctx.stroke();
  ctx.restore();
}

export function headerWidthFor(lastRow: number, zoom: number): number {
  const digits = String(lastRow + 1).length;
  return Math.round(Math.max(50, 16 + digits * 8.5) * Math.min(1.4, Math.max(0.8, zoom)));
}

export function headerHeightFor(zoom: number): number {
  return Math.round(28 * Math.min(1.4, Math.max(0.8, zoom)));
}
