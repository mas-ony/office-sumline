/**
 * Chart data: building a chart definition from a selected range (one series per row or per
 * column, the first row/column holding names) and reading the current values behind a chart's
 * references for rendering and for the cached values written to the file.
 */
import type { SpreadsheetDocument } from './document';
import type { ChartSeries, ChartType, SheetChart } from './model/drawing';
import type { Sheet } from './model/sheet';
import { cellName, colName, parseRange, quoteSheetName, splitSheetRef, type Range } from './refs';
import { RefValue } from './formula/runtime';
import { formatValue } from './format/display';
import type { Scalar } from './values';

export interface ChartDataSeries {
  name: string;
  values: (number | null)[];
  xValues?: (number | null)[];
  sizes?: (number | null)[];
  source: ChartSeries;
  index: number;
}

export interface ChartData {
  categories: string[];
  series: ChartDataSeries[];
  /** Number format of the first value cell (axis labels follow it). */
  numFmt: string;
}

/** Sheet-qualified absolute reference: 'My Sheet'!$A$1:$B$4. */
export function absRef(sheet: Sheet | string, r: Range): string {
  const name = typeof sheet === 'string' ? sheet : sheet.name;
  const a = cellName(r.r0, r.c0, true, true);
  const b = cellName(r.r1, r.c1, true, true);
  return `${quoteSheetName(name)}!${a === b ? a : `${a}:${b}`}`;
}

/** The sheet and range of a single-area reference text ("Sheet1!$A$1:$A$9", or local "A1:A9"). */
export function resolveRef(doc: SpreadsheetDocument, text: string | undefined, base?: Sheet): { sheet: Sheet; range: Range } | null {
  if (!text) return null;
  const { sheet: name, ref } = splitSheetRef(text.replace(/^=/, ''));
  const sheet = name !== undefined ? doc.wb.sheetByName(name) : base ?? doc.sheet;
  const range = parseRange(ref.replace(/\$/g, '').toUpperCase());
  if (sheet && range) return { sheet, range };
  // Defined names and other expressions.
  try {
    const v = doc.engine.evaluateText(text.replace(/^=/, ''), base ?? doc.sheet, 0, 0);
    if (v instanceof RefValue && v.areas.length) return { sheet: v.areas[0].sheet, range: v.areas[0].range };
  } catch {
    // Not a reference.
  }
  return null;
}

/** Cells of a reference in reading order, clipped to the sheet's data (whole columns stay small). */
function cellsOf(doc: SpreadsheetDocument, text: string | undefined, base?: Sheet): { sheet: Sheet; cells: [number, number][] } | null {
  const hit = resolveRef(doc, text, base);
  if (!hit) return null;
  const { sheet } = hit;
  const ext = sheet.dataExtent();
  // dataExtent() gives the last used row and column indexes.
  const range = { ...hit.range, r1: Math.min(hit.range.r1, Math.max(hit.range.r0, ext.rows)), c1: Math.min(hit.range.c1, Math.max(hit.range.c0, ext.cols)) };
  const cells: [number, number][] = [];
  for (let r = range.r0; r <= range.r1; r++) for (let c = range.c0; c <= range.c1; c++) cells.push([r, c]);
  return { sheet, cells: cells.slice(0, 10000) };
}

function toNumber(v: Scalar): number | null {
  return typeof v === 'number' && Number.isFinite(v) ? v : null;
}

export function readNumbers(doc: SpreadsheetDocument, ref: string | undefined, base?: Sheet): (number | null)[] {
  const hit = cellsOf(doc, ref, base);
  if (!hit) return [];
  return hit.cells.map(([r, c]) => toNumber(doc.engine.cellValue(hit.sheet, r, c)));
}

export function readTexts(doc: SpreadsheetDocument, ref: string | undefined, base?: Sheet): string[] {
  const hit = cellsOf(doc, ref, base);
  if (!hit) return [];
  return hit.cells.map(([r, c]) => {
    doc.engine.cellValue(hit.sheet, r, c);
    return doc.displayText(r, c, hit.sheet);
  });
}

/** Current values of a chart's series, names and categories. */
export function readChartData(doc: SpreadsheetDocument, chart: SheetChart, base?: Sheet): ChartData {
  let numFmt = 'General';
  const series = chart.series.map((s, index): ChartDataSeries => {
    const values = readNumbers(doc, s.values, base);
    const name = s.nameRef ? readTexts(doc, s.nameRef, base).join(' ') : s.name ?? '';
    if (index === 0) {
      const first = resolveRef(doc, s.values, base);
      if (first) numFmt = doc.wb.styles.get(first.sheet.styleAt(first.range.r0, first.range.c0)).numFmt || 'General';
    }
    return {
      name: name || `Series ${index + 1}`,
      values,
      xValues: s.xValues ? readNumbers(doc, s.xValues, base) : undefined,
      sizes: s.sizes ? readNumbers(doc, s.sizes, base) : undefined,
      source: s,
      index,
    };
  });
  const n = Math.max(0, ...series.map((s) => s.values.length));
  let categories = chart.categories ? readTexts(doc, chart.categories, base) : [];
  if (!categories.length) categories = Array.from({ length: n }, (_, i) => String(i + 1));
  return { categories, series, numFmt };
}

// -------------------------------------------------------------------------------------------------
// Building charts from a range

export interface ChartLine {
  /** Index of the row or column inside the range. */
  index: number;
  /** Display name: the first cell with headers, "Row 3" / "Column B" without. */
  name: string;
}

export interface ChartLayout {
  seriesIn: 'rows' | 'columns';
  headers: boolean;
  /** Line holding the category labels, or null. */
  labels: number | null;
  /** Lines plotted as series. */
  series: number[];
}

function isNumeric(v: Scalar): boolean {
  return typeof v === 'number';
}

/** The range with empty trailing rows and columns removed (whole-column selections stay small). */
export function trimRange(sheet: Sheet, range: Range): Range {
  const ext = sheet.dataExtent();
  return {
    r0: range.r0,
    c0: range.c0,
    r1: Math.max(range.r0, Math.min(range.r1, ext.rows)),
    c1: Math.max(range.c0, Math.min(range.c1, ext.cols)),
  };
}

/** Lines (rows or columns) of a range with their names. */
export function chartLines(doc: SpreadsheetDocument, sheet: Sheet, range: Range, seriesIn: 'rows' | 'columns', headers: boolean): ChartLine[] {
  const out: ChartLine[] = [];
  const count = seriesIn === 'rows' ? range.r1 - range.r0 + 1 : range.c1 - range.c0 + 1;
  for (let i = 0; i < count; i++) {
    const r = seriesIn === 'rows' ? range.r0 + i : range.r0;
    const c = seriesIn === 'rows' ? range.c0 : range.c0 + i;
    const text = headers ? doc.displayText(r, c, sheet).trim() : '';
    out.push({ index: i, name: text || (seriesIn === 'rows' ? `Row ${r + 1}` : `Column ${colName(c)}`) });
  }
  return out;
}

/** Default layout of a range: Excel's rule (more rows than columns: one series per column). */
export function guessLayout(doc: SpreadsheetDocument, sheet: Sheet, range: Range, type: ChartType): ChartLayout {
  const value = (r: number, c: number) => doc.engine.cellValue(sheet, r, c);
  const rows = range.r1 - range.r0 + 1;
  const cols = range.c1 - range.c0 + 1;
  // Text in the first row (beyond the first cell) or an empty corner: the first row holds names.
  let topText = false;
  for (let c = range.c0 + (cols > 1 ? 1 : 0); c <= range.c1; c++) {
    const v = value(range.r0, c);
    if (v !== null && v !== '' && !isNumeric(v)) topText = true;
  }
  let leftText = false;
  for (let r = range.r0 + (rows > 1 ? 1 : 0); r <= range.r1; r++) {
    const v = value(r, range.c0);
    if (v !== null && v !== '' && !isNumeric(v)) leftText = true;
  }
  const corner = value(range.r0, range.c0);
  const cornerEmpty = corner === null || corner === '';
  const headerRow = rows > 1 && (topText || (cornerEmpty && leftText));
  const labelCol = cols > 1 && (leftText || (cornerEmpty && topText));
  const dataRows = rows - (headerRow ? 1 : 0);
  const dataCols = cols - (labelCol ? 1 : 0);
  const xy = type === 'scatter' || type === 'bubble';
  const single = type === 'histogram' || type === 'box-whisker';
  const seriesIn: 'rows' | 'columns' = xy || single || type === 'stock' || type === 'waterfall' || type === 'funnel' || type === 'treemap' || type === 'pareto' || dataRows > dataCols ? 'columns' : 'rows';
  // In the orientation chosen, "headers" means the first cell of each line names it, and the
  // first line holds the category labels.
  const headers = seriesIn === 'columns' ? headerRow : labelCol;
  const hasLabels = seriesIn === 'columns' ? labelCol : headerRow;
  const count = seriesIn === 'columns' ? cols : rows;
  const all = Array.from({ length: count }, (_, i) => i);
  let labels: number | null = hasLabels ? 0 : null;
  let series = all.filter((i) => i !== labels);
  if (xy && labels === null && series.length > 1) {
    // Scatter: the first column holds the x values.
    labels = series.shift()!;
  }
  if (!series.length) {
    series = all;
    labels = null;
  }
  return { seriesIn, headers, labels, series };
}

/** Series and category references for a layout of a range. */
export function seriesForLayout(sheet: Sheet, range: Range, layout: ChartLayout, type: ChartType, previous: ChartSeries[] = []): { categories?: string; series: ChartSeries[] } {
  const { seriesIn, headers } = layout;
  const line = (i: number): { name?: Range; values: Range } => {
    if (seriesIn === 'columns') {
      const c = range.c0 + i;
      return { name: headers ? { r0: range.r0, c0: c, r1: range.r0, c1: c } : undefined, values: { r0: range.r0 + (headers ? 1 : 0), c0: c, r1: range.r1, c1: c } };
    }
    const r = range.r0 + i;
    return { name: headers ? { r0: r, c0: range.c0, r1: r, c1: range.c0 } : undefined, values: { r0: r, c0: range.c0 + (headers ? 1 : 0), r1: r, c1: range.c1 } };
  };
  const labels = layout.labels !== null ? line(layout.labels).values : undefined;
  const xy = type === 'scatter' || type === 'bubble';
  const ids = layout.series;
  const series: ChartSeries[] = [];
  if (type === 'bubble') {
    // Bubble: pairs of (y, size) columns after the x column.
    for (let k = 0; k < ids.length; k += 2) {
      const y = line(ids[k]);
      const size = ids[k + 1] !== undefined ? line(ids[k + 1]) : undefined;
      series.push({
        ...previous[series.length],
        nameRef: y.name ? absRef(sheet, y.name) : undefined,
        name: y.name ? undefined : `Series ${series.length + 1}`,
        values: absRef(sheet, y.values),
        xValues: labels ? absRef(sheet, labels) : undefined,
        sizes: size ? absRef(sheet, size.values) : undefined,
      });
    }
  } else {
    for (const i of ids) {
      const l = line(i);
      series.push({
        ...previous[series.length],
        nameRef: l.name ? absRef(sheet, l.name) : undefined,
        name: l.name ? undefined : `Series ${series.length + 1}`,
        values: absRef(sheet, l.values),
        xValues: xy && labels ? absRef(sheet, labels) : undefined,
      });
    }
  }
  return { categories: !xy && labels ? absRef(sheet, labels) : undefined, series };
}

/** The layout of an existing chart within its range, when its references match one. */
export function layoutOfChart(doc: SpreadsheetDocument, chart: SheetChart): ChartLayout | null {
  const hit = resolveRef(doc, chart.range);
  if (!hit) return null;
  for (const headers of [chart.headers, !chart.headers]) {
    const layout: ChartLayout = { seriesIn: chart.seriesIn, headers, labels: null, series: [] };
    const count = chart.seriesIn === 'columns' ? hit.range.c1 - hit.range.c0 + 1 : hit.range.r1 - hit.range.r0 + 1;
    const refs = Array.from({ length: count }, (_, i) => seriesForLayout(hit.sheet, hit.range, { ...layout, labels: i, series: [i] }, 'line'));
    const valuesOf = (i: number) => refs[i].series[0].values;
    const norm = (s?: string) => (s ?? '').replace(/\$/g, '').toUpperCase();
    const xy = chart.type === 'scatter' || chart.type === 'bubble';
    const catRef = xy ? chart.series[0]?.xValues : chart.categories;
    layout.labels = catRef ? (() => {
      const i = refs.findIndex((_, k) => norm(valuesOf(k)) === norm(catRef));
      return i >= 0 ? i : null;
    })() : null;
    if (catRef && layout.labels === null) continue;
    const ids: number[] = [];
    for (const s of chart.series) {
      const i = refs.findIndex((_, k) => norm(valuesOf(k)) === norm(s.values));
      if (i < 0) break;
      ids.push(i);
      if (chart.type === 'bubble' && s.sizes) {
        const j = refs.findIndex((_, k) => norm(valuesOf(k)) === norm(s.sizes));
        if (j >= 0) ids.push(j);
      }
    }
    if (ids.length < chart.series.length) continue;
    layout.series = ids;
    return layout;
  }
  return null;
}

/** Bounding range of every reference of a chart when they are all on one sheet. */
export function chartBounds(doc: SpreadsheetDocument, chart: Pick<SheetChart, 'series' | 'categories'>): string | undefined {
  let sheet: Sheet | null = null;
  let box: Range | null = null;
  const refs = [chart.categories, ...chart.series.flatMap((s) => [s.nameRef, s.values, s.xValues, s.sizes])];
  for (const ref of refs) {
    if (!ref) continue;
    const hit = resolveRef(doc, ref);
    if (!hit) return undefined;
    if (sheet && hit.sheet !== sheet) return undefined;
    sheet = hit.sheet;
    box = box ? { r0: Math.min(box.r0, hit.range.r0), c0: Math.min(box.c0, hit.range.c0), r1: Math.max(box.r1, hit.range.r1), c1: Math.max(box.c1, hit.range.c1) } : { ...hit.range };
  }
  return sheet && box ? absRef(sheet, box) : undefined;
}

/** A complete new chart (without position) for a range of the active sheet. */
export function chartFromRange(doc: SpreadsheetDocument, sheet: Sheet, selection: Range, type: ChartType): Omit<SheetChart, 'id' | 'z' | 'from' | 'to'> {
  const range = trimRange(sheet, selection);
  const layout = guessLayout(doc, sheet, range, type);
  const { categories, series } = seriesForLayout(sheet, range, layout, type);
  return {
    name: 'Chart',
    type,
    range: absRef(sheet, range),
    seriesIn: layout.seriesIn,
    headers: layout.headers,
    categories,
    series,
    showLegend: true,
    gridlines: true,
    valueLabels: false,
  };
}

/** Local text of a sheet-qualified reference when it points to `sheet` ("A1:C3"). */
export function localRef(ref: string | undefined, sheet: Sheet): string {
  if (!ref) return '';
  const { sheet: name, ref: local } = splitSheetRef(ref);
  const bare = local.replace(/\$/g, '');
  return name === undefined || name === sheet.name ? bare : `${quoteSheetName(name)}!${bare}`;
}

/** Formats a value the way axis labels show it. */
export function formatAxisValue(v: number, numFmt: string): string {
  if (numFmt && numFmt !== 'General' && numFmt !== '@') return formatValue(v, numFmt).text;
  const a = Math.abs(v);
  if (a >= 1e9) return `${Number((v / 1e9).toPrecision(3))}B`;
  if (a >= 1e6) return `${Number((v / 1e6).toPrecision(3))}M`;
  if (a >= 1e4) return `${Number((v / 1e3).toPrecision(4))}K`;
  return String(Number(v.toPrecision(10)));
}
