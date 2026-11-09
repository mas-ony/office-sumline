/**
 * Print layout: splits the printed area of one or more sheets into pages (paper size,
 * orientation, scaling, margins, page order, alignment, repeated frozen rows and columns,
 * headings, header and footer) and describes every page as a list of drawing operations in
 * points. The same pages are drawn as SVG (preview and printing) and written as PDF.
 */
import type { SpreadsheetDocument } from '../core/document';
import type { Sheet } from '../core/model/sheet';
import type { BorderLine, CellStyle } from '../core/model/style';
import { resolveColor } from '../core/model/color';
import { SheetLayout } from '../core/layout';
import { cellName, parseRange, splitSheetRef, type Range } from '../core/refs';
import { formatValue } from '../core/format/display';
import { CellError } from '../core/values';
import { readChartData } from '../core/chartData';
import { fillColor } from '../grid/render';
import { condFormatFor } from '../grid/condFormat';
import { chartSvg } from '../grid/chartRender';
import { shapeSvg } from '../grid/shapes';

export type PaperSize = 'A4' | 'Letter' | 'Legal' | 'A3' | 'A5' | 'Tabloid';

export const PAPER_SIZES: Record<PaperSize, { w: number; h: number; label: string }> = {
  A4: { w: 595.28, h: 841.89, label: 'A4' },
  Letter: { w: 612, h: 792, label: 'Letter' },
  Legal: { w: 612, h: 1008, label: 'Legal' },
  A3: { w: 841.89, h: 1190.55, label: 'A3' },
  A5: { w: 419.53, h: 595.28, label: 'A5' },
  Tabloid: { w: 792, h: 1224, label: 'Tabloid' },
};

export type MarginPreset = 'normal' | 'narrow' | 'wide';

export const MARGINS: Record<MarginPreset, { t: number; r: number; b: number; l: number; label: string }> = {
  normal: { t: 54, r: 50.4, b: 54, l: 50.4, label: 'Normal' },
  narrow: { t: 54, r: 18, b: 54, l: 18, label: 'Narrow' },
  wide: { t: 72, r: 72, b: 72, l: 72, label: 'Wide' },
};

export type ScaleMode = 'normal' | 'fitWidth' | 'fitHeight' | 'fitPage';

export interface PrintOptions {
  scope: 'sheet' | 'selection' | 'workbook';
  paper: PaperSize;
  orientation: 'portrait' | 'landscape';
  scale: ScaleMode;
  margins: MarginPreset;
  /** Over then down (left to right), or down then over (top to bottom). */
  order: 'leftToRight' | 'topToBottom';
  hAlign: 'left' | 'center' | 'right';
  vAlign: 'top' | 'center' | 'bottom';
  gridlines: boolean;
  notes: boolean;
  headings: boolean;
  pageNumbers: boolean;
  spreadsheetTitle: boolean;
  worksheetTitle: boolean;
  date: boolean;
  time: boolean;
  repeatRows: boolean;
  repeatCols: boolean;
  /** Pictures, charts and shapes. */
  objects: boolean;
}

export const DEFAULT_PRINT_OPTIONS: PrintOptions = {
  scope: 'sheet',
  paper: 'A4',
  orientation: 'landscape',
  scale: 'normal',
  margins: 'normal',
  order: 'leftToRight',
  hAlign: 'center',
  vAlign: 'center',
  gridlines: true,
  notes: false,
  headings: false,
  pageNumbers: false,
  spreadsheetTitle: false,
  worksheetTitle: false,
  date: false,
  time: false,
  repeatRows: true,
  repeatCols: false,
  objects: true,
};

export type FontFamily = 'sans' | 'serif' | 'mono';

export interface FontSpec {
  family: FontFamily;
  bold: boolean;
  italic: boolean;
  /** Points. */
  size: number;
}

export interface Box {
  x: number;
  y: number;
  w: number;
  h: number;
}

export type ImageSrc = { kind: 'svg'; markup: string; w: number; h: number } | { kind: 'bitmap'; data: Uint8Array; mime: string };

export type Op =
  | { k: 'rect'; x: number; y: number; w: number; h: number; fill: string }
  | { k: 'line'; x1: number; y1: number; x2: number; y2: number; color: string; width: number; dash?: number[] }
  /** Text from (x, baseline y). */
  | { k: 'text'; x: number; y: number; text: string; font: FontSpec; color: string; underline?: boolean; strike?: boolean; clip?: Box }
  | { k: 'image'; x: number; y: number; w: number; h: number; src: ImageSrc; clip?: Box };

export interface Page {
  width: number;
  height: number;
  ops: Op[];
  sheet: string;
}

/** Width of a text in points at a font's size. */
export type Measure = (text: string, font: FontSpec) => number;

/** Approximate measurement (tests and environments without a canvas). */
export const approxMeasure: Measure = (text, font) => {
  const k = font.family === 'mono' ? 0.6 : font.bold ? 0.58 : 0.52;
  return text.length * font.size * k;
};

export function fontFamilyOf(name: string): FontFamily {
  if (/courier|consol|mono|lucida console/i.test(name)) return 'mono';
  if (/times|georgia|cambria|garamond|book|serif|palatino|minion/i.test(name) && !/sans/i.test(name)) return 'serif';
  return 'sans';
}

const PX = 0.75;

interface Area {
  sheet: Sheet;
  range: Range | null;
}

/** The print area of a sheet: its defined Print_Area, else the used range with its objects. */
function sheetArea(doc: SpreadsheetDocument, sheet: Sheet, withObjects: boolean): Range | null {
  const index = doc.wb.sheets.indexOf(sheet);
  const named = doc.wb.names.find((n) => /^_xlnm\.Print_Area$/i.test(n.name) && n.localSheet === index);
  if (named) {
    const first = named.formula.split(',')[0];
    const r = parseRange(splitSheetRef(first).ref.replace(/\$/g, ''));
    if (r) return r;
  }
  let area = sheet.usedRange(true);
  if (withObjects) {
    const layout = SheetLayout.of(sheet);
    const rects = [
      ...sheet.images.map((i) => layout.imageRect(i)),
      ...sheet.charts.map((c) => layout.anchorRect(c.from, c.to)),
      ...sheet.shapes.map((s) => layout.anchorRect(s.from, s.to)),
    ];
    for (const r of rects) {
      const c1 = layout.colAt(r.x + r.w - 1);
      const r1 = layout.rowAt(r.y + r.h - 1);
      area = area
        ? { r0: Math.min(area.r0, layout.rowAt(r.y)), c0: Math.min(area.c0, layout.colAt(r.x)), r1: Math.max(area.r1, r1), c1: Math.max(area.c1, c1) }
        : { r0: layout.rowAt(r.y), c0: layout.colAt(r.x), r1, c1 };
    }
  }
  return area;
}

function areasFor(doc: SpreadsheetDocument, opts: PrintOptions): Area[] {
  if (opts.scope === 'selection') {
    const r = doc.range;
    const ext = doc.sheet.dataExtent();
    // Whole rows or columns print as far as the data goes.
    const range = { r0: r.r0, c0: r.c0, r1: Math.min(r.r1, Math.max(r.r0, ext.rows)), c1: Math.min(r.c1, Math.max(r.c0, ext.cols)) };
    return [{ sheet: doc.sheet, range }];
  }
  const sheets = opts.scope === 'workbook' ? doc.wb.sheets.filter((s) => s.state === 'visible' && !s.chartsheet) : [doc.sheet];
  return sheets.map((sheet) => ({ sheet, range: sheetArea(doc, sheet, opts.objects) }));
}

function chunks(sizes: { i: number; size: number }[], avail: number): number[][] {
  const out: number[][] = [];
  let cur: number[] = [];
  let used = 0;
  for (const { i, size } of sizes) {
    if (cur.length && used + size > avail + 0.01) {
      out.push(cur);
      cur = [];
      used = 0;
    }
    cur.push(i);
    used += size;
  }
  if (cur.length) out.push(cur);
  return out.length ? out : [[]];
}

function borderSpec(line: BorderLine | undefined, doc: SpreadsheetDocument): { color: string; width: number; dash?: number[]; double?: boolean } | null {
  if (!line) return null;
  const color = `#${resolveColor(line.color, doc.wb.theme, '000000')}`;
  switch (line.style) {
    case 'medium':
      return { color, width: 1 };
    case 'thick':
      return { color, width: 1.5 };
    case 'double':
      return { color, width: 0.5, double: true };
    case 'hair':
      return { color, width: 0.25 };
    case 'dotted':
      return { color, width: 0.5, dash: [1, 1] };
    case 'dashed':
      return { color, width: 0.5, dash: [3, 1.5] };
    case 'mediumDashed':
      return { color, width: 1, dash: [4, 2] };
    case 'dashDot':
    case 'dashDotDot':
    case 'slantDashDot':
      return { color, width: 0.5, dash: [3, 1.5, 1, 1.5] };
    case 'mediumDashDot':
    case 'mediumDashDotDot':
      return { color, width: 1, dash: [4, 2, 1, 2] };
    default:
      return { color, width: 0.5 };
  }
}

function wrapLines(text: string, maxW: number, font: FontSpec, measure: Measure): string[] {
  const out: string[] = [];
  for (const para of text.split('\n')) {
    const words = para.split(/(\s+)/);
    let line = '';
    for (const w of words) {
      if (!w) continue;
      const next = line + w;
      if (line && measure(next.trimEnd(), font) > maxW) {
        out.push(line.trimEnd());
        line = w.trimStart();
      } else line = next;
      // Words longer than the cell break anywhere.
      while (line && measure(line, font) > maxW && line.length > 1) {
        let k = line.length - 1;
        while (k > 1 && measure(line.slice(0, k), font) > maxW) k--;
        out.push(line.slice(0, k));
        line = line.slice(k);
      }
    }
    out.push(line.trimEnd());
  }
  return out;
}

export interface BuildContext {
  /** File name shown with "Show spreadsheet title". */
  title: string;
  now?: Date;
}

/** Lays out the printed pages. */
export function buildPages(doc: SpreadsheetDocument, opts: PrintOptions, measure: Measure, ctx: BuildContext): Page[] {
  const paper = PAPER_SIZES[opts.paper];
  const pw = opts.orientation === 'landscape' ? Math.max(paper.w, paper.h) : Math.min(paper.w, paper.h);
  const ph = opts.orientation === 'landscape' ? Math.min(paper.w, paper.h) : Math.max(paper.w, paper.h);
  const m = MARGINS[opts.margins];
  const contentW = pw - m.l - m.r;
  const contentH = ph - m.t - m.b;
  const pages: Page[] = [];
  const areas = areasFor(doc, opts);
  for (const { sheet, range } of areas) {
    const layout = SheetLayout.of(sheet);
    const colW = (c: number) => layout.colWidth(c) * PX;
    const rowH = (r: number) => layout.rowHeight(r) * PX;
    const area = range ?? { r0: 0, c0: 0, r1: 0, c1: 0 };
    const titleRows: number[] = [];
    const titleCols: number[] = [];
    if (opts.repeatRows && sheet.freeze.rows > 0 && opts.scope !== 'selection') for (let r = 0; r < Math.min(sheet.freeze.rows, area.r1 + 1); r++) titleRows.push(r);
    if (opts.repeatCols && sheet.freeze.cols > 0 && opts.scope !== 'selection') for (let c = 0; c < Math.min(sheet.freeze.cols, area.c1 + 1); c++) titleCols.push(c);
    const bodyRows: { i: number; size: number }[] = [];
    const bodyCols: { i: number; size: number }[] = [];
    for (let r = area.r0; r <= area.r1; r++) if (!titleRows.includes(r) && rowH(r) > 0) bodyRows.push({ i: r, size: rowH(r) });
    for (let c = area.c0; c <= area.c1; c++) if (!titleCols.includes(c) && colW(c) > 0) bodyCols.push({ i: c, size: colW(c) });
    const headW = opts.headings ? Math.max(20, approxMeasure(String(area.r1 + 1), { family: 'sans', bold: false, italic: false, size: 8 }) + 10) : 0;
    const headH = opts.headings ? 14 : 0;
    const titleW = titleCols.reduce((a, c) => a + colW(c), 0);
    const titleH = titleRows.reduce((a, r) => a + rowH(r), 0);
    const totalW = headW + titleW + bodyCols.reduce((a, c) => a + c.size, 0);
    const totalH = headH + titleH + bodyRows.reduce((a, r) => a + r.size, 0);
    let s = 1;
    if (opts.scale === 'fitWidth') s = Math.min(1, contentW / Math.max(1, totalW));
    else if (opts.scale === 'fitHeight') s = Math.min(1, contentH / Math.max(1, totalH));
    else if (opts.scale === 'fitPage') s = Math.min(1, contentW / Math.max(1, totalW), contentH / Math.max(1, totalH));
    const availW = Math.max(20, contentW / s - headW - titleW);
    const availH = Math.max(20, contentH / s - headH - titleH);
    const colChunks = chunks(bodyCols, availW);
    const rowChunks = chunks(bodyRows, availH);
    const order: [number[], number[]][] = [];
    if (opts.order === 'leftToRight') for (const rc of rowChunks) for (const cc of colChunks) order.push([rc, cc]);
    else for (const cc of colChunks) for (const rc of rowChunks) order.push([rc, cc]);
    const cf = condFormatFor(doc, sheet);
    for (const [rowsChunk, colsChunk] of order) {
      const rows = [...(rowsChunk.length && titleRows.length && rowsChunk[0] > titleRows[titleRows.length - 1] ? titleRows : []), ...rowsChunk];
      const cols = [...(colsChunk.length && titleCols.length && colsChunk[0] > titleCols[titleCols.length - 1] ? titleCols : []), ...colsChunk];
      const page = drawPage(doc, sheet, rows, cols, rowsChunk, colsChunk, { s, headW, headH, pw, ph, m, contentW, contentH, opts, measure, cf, area: range });
      page.sheet = sheet.name;
      pages.push(page);
    }
    if (opts.notes) {
      const notes = sheet.comments.filter((t) => !range || (t.r >= range.r0 && t.r <= range.r1 && t.c >= range.c0 && t.c <= range.c1));
      if (notes.length) pages.push(...notesPages(sheet, notes, pw, ph, m, measure));
    }
  }
  if (!pages.length) pages.push({ width: 595.28, height: 841.89, ops: [], sheet: doc.sheet.name });
  // Header and footer.
  const now = ctx.now ?? new Date();
  const total = pages.length;
  pages.forEach((page, i) => {
    const small: FontSpec = { family: 'sans', bold: false, italic: false, size: 9 };
    const left = [opts.spreadsheetTitle ? ctx.title : '', opts.worksheetTitle ? page.sheet : ''].filter(Boolean).join(' - ');
    const right = [opts.date ? now.toLocaleDateString() : '', opts.time ? now.toLocaleTimeString([], { hour: '2-digit', minute: '2-digit' }) : ''].filter(Boolean).join(' ');
    const y = Math.max(14, m.t / 2 + 3);
    if (left) page.ops.push({ k: 'text', x: m.l, y, text: left, font: small, color: '#333333' });
    if (right) page.ops.push({ k: 'text', x: page.width - m.r - measure(right, small), y, text: right, font: small, color: '#333333' });
    if (opts.pageNumbers) {
      const label = `Page ${i + 1} of ${total}`;
      page.ops.push({ k: 'text', x: (page.width - measure(label, small)) / 2, y: page.height - Math.max(12, m.b / 2 - 3), text: label, font: small, color: '#333333' });
    }
  });
  return pages;
}

interface DrawContext {
  s: number;
  headW: number;
  headH: number;
  pw: number;
  ph: number;
  m: { t: number; r: number; b: number; l: number };
  contentW: number;
  contentH: number;
  opts: PrintOptions;
  measure: Measure;
  cf: ReturnType<typeof condFormatFor>;
  area: Range | null;
}

function drawPage(doc: SpreadsheetDocument, sheet: Sheet, rows: number[], cols: number[], bodyRows: number[], bodyCols: number[], dc: DrawContext): Page {
  const { s, opts, measure } = dc;
  const layout = SheetLayout.of(sheet);
  const colW = (c: number) => layout.colWidth(c) * PX;
  const rowH = (r: number) => layout.rowHeight(r) * PX;
  const ops: Op[] = [];
  // Unscaled content size, then its position on the page.
  const cw = dc.headW + cols.reduce((a, c) => a + colW(c), 0);
  const ch = dc.headH + rows.reduce((a, r) => a + rowH(r), 0);
  const ox = dc.m.l + (opts.hAlign === 'center' ? Math.max(0, (dc.contentW - cw * s) / 2) : opts.hAlign === 'right' ? Math.max(0, dc.contentW - cw * s) : 0);
  const oy = dc.m.t + (opts.vAlign === 'center' ? Math.max(0, (dc.contentH - ch * s) / 2) : opts.vAlign === 'bottom' ? Math.max(0, dc.contentH - ch * s) : 0);
  const X = (u: number) => ox + u * s;
  const Y = (v: number) => oy + v * s;
  // Column and row positions (unscaled, relative to the content origin).
  const colX = new Map<number, number>();
  let u = dc.headW;
  for (const c of cols) {
    colX.set(c, u);
    u += colW(c);
  }
  const rowY = new Map<number, number>();
  let v = dc.headH;
  for (const r of rows) {
    rowY.set(r, v);
    v += rowH(r);
  }
  const theme = doc.wb.theme;
  const merges = sheet.merges.filter((mg) => rows.some((r) => r >= mg.r0 && r <= mg.r1) && cols.some((c) => c >= mg.c0 && c <= mg.c1));
  const mergeOf = (r: number, c: number) => merges.find((mg) => r >= mg.r0 && r <= mg.r1 && c >= mg.c0 && c <= mg.c1);
  /** Box of a cell (or of the visible part of its merge) in unscaled content coordinates. */
  const cellBox = (r: number, c: number): Box | null => {
    const mg = mergeOf(r, c);
    if (!mg) return { x: colX.get(c)!, y: rowY.get(r)!, w: colW(c), h: rowH(r) };
    const rs = rows.filter((x) => x >= mg.r0 && x <= mg.r1);
    const cs = cols.filter((x) => x >= mg.c0 && x <= mg.c1);
    if (rs[0] !== r || cs[0] !== c) return null;
    return { x: colX.get(c)!, y: rowY.get(r)!, w: cs.reduce((a, k) => a + colW(k), 0), h: rs.reduce((a, k) => a + rowH(k), 0) };
  };
  const styleOf = (r: number, c: number): CellStyle => doc.wb.styles.get(sheet.styleAt(r, c));

  // Fills.
  for (const r of rows) {
    for (const c of cols) {
      const box = cellBox(r, c);
      if (!box) continue;
      const style = styleOf(r, c);
      const fill = dc.cf?.(r, c)?.fill ?? fillColor(style, theme);
      if (fill) ops.push({ k: 'rect', x: X(box.x), y: Y(box.y), w: box.w * s, h: box.h * s, fill });
    }
  }
  // Gridlines (not inside merged cells).
  if (opts.gridlines) {
    const grid = '#d4d4d4';
    const bottom = Y(ch);
    const right = X(cw);
    for (const c of [...cols, -1]) {
      const x = c === -1 ? cw : colX.get(c)!;
      // Split vertical lines around merges.
      let y0 = dc.headH;
      for (const r of rows) {
        const mg = mergeOf(r, c);
        if (c !== -1 && mg && c > mg.c0) {
          if (rowY.get(r)! > y0) ops.push({ k: 'line', x1: X(x), y1: Y(y0), x2: X(x), y2: Y(rowY.get(r)!), color: grid, width: 0.25 });
          y0 = rowY.get(r)! + rowH(r);
        }
      }
      if (y0 < ch) ops.push({ k: 'line', x1: X(x), y1: Y(y0), x2: X(x), y2: bottom, color: grid, width: 0.25 });
    }
    for (const r of [...rows, -1]) {
      const y = r === -1 ? ch : rowY.get(r)!;
      let x0 = dc.headW;
      for (const c of cols) {
        const mg = mergeOf(r, c);
        if (r !== -1 && mg && r > mg.r0) {
          if (colX.get(c)! > x0) ops.push({ k: 'line', x1: X(x0), y1: Y(y), x2: X(colX.get(c)!), y2: Y(y), color: grid, width: 0.25 });
          x0 = colX.get(c)! + colW(c);
        }
      }
      if (x0 < cw) ops.push({ k: 'line', x1: X(x0), y1: Y(y), x2: right, y2: Y(y), color: grid, width: 0.25 });
    }
  }
  // Borders.
  const edge = (x1: number, y1: number, x2: number, y2: number, b: ReturnType<typeof borderSpec>, vertical: boolean) => {
    if (!b) return;
    if (b.double) {
      const d = 0.75;
      for (const k of [-d, d]) ops.push({ k: 'line', x1: X(x1) + (vertical ? k : 0), y1: Y(y1) + (vertical ? 0 : k), x2: X(x2) + (vertical ? k : 0), y2: Y(y2) + (vertical ? 0 : k), color: b.color, width: 0.4 });
      return;
    }
    ops.push({ k: 'line', x1: X(x1), y1: Y(y1), x2: X(x2), y2: Y(y2), color: b.color, width: b.width, dash: b.dash });
  };
  for (const r of rows) {
    for (const c of cols) {
      const box = cellBox(r, c);
      if (!box) continue;
      const b = styleOf(r, c).border;
      if (!b) continue;
      // Merged cells use the borders of their corner cells.
      const mg = mergeOf(r, c);
      const right = mg ? styleOf(r, mg.c1).border : b;
      const bottom = mg ? styleOf(mg.r1, c).border : b;
      edge(box.x, box.y, box.x + box.w, box.y, borderSpec(b.top, doc), false);
      edge(box.x, box.y, box.x, box.y + box.h, borderSpec(b.left, doc), true);
      edge(box.x + box.w, box.y, box.x + box.w, box.y + box.h, borderSpec(right?.right, doc), true);
      edge(box.x, box.y + box.h, box.x + box.w, box.y + box.h, borderSpec(bottom?.bottom, doc), false);
    }
  }
  // Text.
  for (const r of rows) {
    for (let ci = 0; ci < cols.length; ci++) {
      const c = cols[ci];
      const box = cellBox(r, c);
      if (!box) continue;
      const cell = sheet.getCell(r, c);
      if (!cell || cell.v === null || cell.v === '') continue;
      const style = styleOf(r, c);
      const cfr = dc.cf?.(r, c);
      const f = { ...style.font, ...(cfr?.font ?? {}) };
      const formatted = formatValue(cell.v, style.numFmt, doc.wb.date1904);
      const text = formatted.text;
      if (!text) continue;
      const font: FontSpec = { family: fontFamilyOf(f.name), bold: !!f.bold, italic: !!f.italic, size: Math.max(1, f.size * s) };
      const color = cfr?.color ?? formatted.color ?? `#${resolveColor(f.color, doc.wb.theme, '000000')}`;
      const a = style.align ?? {};
      let h = a.horizontal ?? 'general';
      if (h === 'general') h = formatted.numeric ? 'right' : typeof cell.v === 'boolean' || cell.v instanceof CellError ? 'center' : 'left';
      if (h === 'centerContinuous') h = 'center';
      if (h === 'fill' || h === 'distributed' || h === 'justify') h = 'left';
      const pad = 2 * s;
      const indent = (a.indent ?? 0) * 6.75 * s;
      const bx = X(box.x);
      const by = Y(box.y);
      const bw = box.w * s;
      const bh = box.h * s;
      const lineH = font.size * 1.2;
      let lines = [text];
      if (a.wrap) lines = wrapLines(text, Math.max(4, bw - 2 * pad - indent), font, measure);
      else if (text.includes('\n')) lines = text.split('\n');
      const blockH = lines.length * lineH;
      const vAlign = a.vertical ?? 'bottom';
      let firstBase: number;
      if (vAlign === 'top') firstBase = by + pad + font.size * 0.92;
      else if (vAlign === 'center' || vAlign === 'justify' || vAlign === 'distributed') firstBase = by + (bh - blockH) / 2 + font.size * 0.95;
      else firstBase = by + bh - pad - blockH + font.size * 0.95;
      // Overflow into empty neighbours for single-line text.
      let clip: Box = { x: bx, y: by, w: bw, h: bh };
      const width0 = measure(lines[0], font);
      if (!a.wrap && lines.length === 1 && width0 > bw - 2 * pad && !formatted.numeric && !mergeOf(r, c)) {
        let extendR = 0;
        let extendL = 0;
        if (h === 'left' || h === 'center') {
          for (let k = ci + 1; k < cols.length && width0 > bw - 2 * pad + extendR; k++) {
            const nc = cols[k];
            const n = sheet.getCell(r, nc);
            if ((n && n.v !== null && n.v !== '') || mergeOf(r, nc)) break;
            extendR += colW(nc) * s;
          }
        }
        if (h === 'right' || h === 'center') {
          for (let k = ci - 1; k >= 0 && width0 > bw - 2 * pad + extendL; k--) {
            const nc = cols[k];
            const n = sheet.getCell(r, nc);
            if ((n && n.v !== null && n.v !== '') || mergeOf(r, nc)) break;
            extendL += colW(nc) * s;
          }
        }
        clip = { x: bx - extendL, y: by, w: bw + extendL + extendR, h: bh };
      }
      lines.forEach((line, k) => {
        if (!line) return;
        const tw = measure(line, font);
        let x: number;
        if (h === 'right') x = bx + bw - pad - tw - indent;
        else if (h === 'center') x = bx + (bw - tw) / 2;
        else x = bx + pad + indent;
        ops.push({ k: 'text', x, y: firstBase + k * lineH, text: line, font, color, underline: !!f.underline, strike: !!f.strike, clip });
      });
    }
  }
  // Pictures, charts and shapes over the body area.
  if (opts.objects) {
    const firstCol = bodyCols[0];
    const firstRow = bodyRows[0];
    if (firstCol !== undefined && firstRow !== undefined) {
      const bodyX0 = colX.get(firstCol)!;
      const bodyY0 = rowY.get(firstRow)!;
      const lastCol = bodyCols[bodyCols.length - 1];
      const lastRow = bodyRows[bodyRows.length - 1];
      const clip: Box = { x: X(bodyX0), y: Y(bodyY0), w: (colX.get(lastCol)! + colW(lastCol) - bodyX0) * s, h: (rowY.get(lastRow)! + rowH(lastRow) - bodyY0) * s };
      const sheetX0 = layout.colX(firstCol) * PX;
      const sheetY0 = layout.rowY(firstRow) * PX;
      const place = (rect: { x: number; y: number; w: number; h: number }) => ({ x: X(bodyX0 + rect.x * PX - sheetX0), y: Y(bodyY0 + rect.y * PX - sheetY0), w: rect.w * PX * s, h: rect.h * PX * s });
      const visible = (b: Box) => b.x < clip.x + clip.w && b.x + b.w > clip.x && b.y < clip.y + clip.h && b.y + b.h > clip.y;
      const objects: { z: number; op: () => Op | null }[] = [
        ...sheet.images.map((img, i) => ({
          z: img.z ?? i,
          op: () => {
            const b = place(layout.imageRect(img));
            return visible(b) ? ({ k: 'image', ...b, src: { kind: 'bitmap', data: img.data, mime: img.mime }, clip } as Op) : null;
          },
        })),
        ...sheet.charts.map((chart) => ({
          z: chart.z,
          op: () => {
            const rect = layout.anchorRect(chart.from, chart.to);
            const b = place(rect);
            if (!visible(b)) return null;
            const markup = chartSvg(chart, readChartData(doc, chart, sheet), rect.w, rect.h);
            return { k: 'image', ...b, src: { kind: 'svg', markup, w: rect.w, h: rect.h }, clip } as Op;
          },
        })),
        ...sheet.shapes.map((shape) => ({
          z: shape.z,
          op: () => {
            const rect = layout.anchorRect(shape.from, shape.to);
            const b = place(rect);
            if (!visible(b)) return null;
            return { k: 'image', ...b, src: { kind: 'svg', markup: shapeSvg(shape, rect.w, rect.h), w: rect.w, h: rect.h }, clip } as Op;
          },
        })),
      ];
      objects.sort((p, q) => p.z - q.z);
      for (const o of objects) {
        const op = o.op();
        if (op) ops.push(op);
      }
    }
  }
  // Row and column headings.
  if (opts.headings) {
    const hf: FontSpec = { family: 'sans', bold: false, italic: false, size: 8 * s };
    const bg = '#f3f3f3';
    const line = '#bfbfbf';
    ops.push({ k: 'rect', x: X(0), y: Y(0), w: cw * s, h: dc.headH * s, fill: bg });
    ops.push({ k: 'rect', x: X(0), y: Y(0), w: dc.headW * s, h: ch * s, fill: bg });
    for (const c of cols) {
      const label = cellName(0, c).replace(/\d+$/, '');
      const x = X(colX.get(c)!);
      const w = colW(c) * s;
      ops.push({ k: 'text', x: x + (w - measure(label, hf)) / 2, y: Y(0) + dc.headH * s * 0.72, text: label, font: hf, color: '#444444', clip: { x, y: Y(0), w, h: dc.headH * s } });
      ops.push({ k: 'line', x1: x + w, y1: Y(0), x2: x + w, y2: Y(dc.headH), color: line, width: 0.4 });
    }
    for (const r of rows) {
      const label = String(r + 1);
      const y = Y(rowY.get(r)!);
      const h = rowH(r) * s;
      ops.push({ k: 'text', x: X(0) + (dc.headW * s - measure(label, hf)) / 2, y: y + h / 2 + hf.size * 0.35, text: label, font: hf, color: '#444444', clip: { x: X(0), y, w: dc.headW * s, h } });
      ops.push({ k: 'line', x1: X(0), y1: y + h, x2: X(dc.headW), y2: y + h, color: line, width: 0.4 });
    }
    ops.push({ k: 'line', x1: X(0), y1: Y(dc.headH), x2: X(cw), y2: Y(dc.headH), color: line, width: 0.6 });
    ops.push({ k: 'line', x1: X(dc.headW), y1: Y(0), x2: X(dc.headW), y2: Y(ch), color: line, width: 0.6 });
  }
  return { width: dc.pw, height: dc.ph, ops, sheet: sheet.name };
}

function notesPages(sheet: Sheet, notes: Sheet['comments'], pw: number, ph: number, m: { t: number; r: number; b: number; l: number }, measure: Measure): Page[] {
  const pages: Page[] = [];
  const title: FontSpec = { family: 'sans', bold: true, italic: false, size: 12 };
  const label: FontSpec = { family: 'sans', bold: true, italic: false, size: 10 };
  const body: FontSpec = { family: 'sans', bold: false, italic: false, size: 10 };
  let page: Page = { width: pw, height: ph, ops: [], sheet: sheet.name };
  let y = m.t + 14;
  page.ops.push({ k: 'text', x: m.l, y, text: `Notes - ${sheet.name}`, font: title, color: '#000000' });
  y += 22;
  for (const t of notes) {
    const text = t.comments.map((cm) => (t.kind === 'threaded' && cm.author ? `${cm.author}: ${cm.text}` : cm.text)).join('\n');
    const lines = wrapLines(text, pw - m.l - m.r - 60, body, measure);
    if (y + lines.length * 13 > ph - m.b) {
      pages.push(page);
      page = { width: pw, height: ph, ops: [], sheet: sheet.name };
      y = m.t + 14;
    }
    page.ops.push({ k: 'text', x: m.l, y, text: cellName(t.r, t.c), font: label, color: '#000000' });
    lines.forEach((l, i) => page.ops.push({ k: 'text', x: m.l + 60, y: y + i * 13, text: l, font: body, color: '#000000' }));
    y += lines.length * 13 + 8;
  }
  pages.push(page);
  return pages;
}
