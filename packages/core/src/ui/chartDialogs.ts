/**
 * The chart dialogs' logic: inserting a chart for a data range, and the chart settings (data
 * layout, style and the series' styles) read from a chart and written back to it.
 */
import { absRef, chartBounds, chartFromRange, type ChartLayout, guessLayout, layoutOfChart, localRef, readChartData, resolveRef, seriesForLayout, trimRange } from '../core/chartData';
import type { SpreadsheetDocument } from '../core/document';
import { SheetLayout } from '../core/layout';
import type { ChartSeries, ChartType, LineDash, SheetChart, Side, Trendline } from '../core/model/drawing';
import type { Sheet } from '../core/model/sheet';
import type { Range } from '../core/refs';

const CHART_W = 400;
const CHART_H = 300;

/**
 * Inserts a chart of the data in `rangeText`: next to the data when it is on this sheet (below it on
 * phones), so the chart does not cover it; else at the active cell. Returns the chart, or why not.
 */
export function insertChart(doc: SpreadsheetDocument, rangeText: string, type: ChartType, phone: boolean): { chart: SheetChart | null } | { error: string } {
  const sheet = doc.sheet;
  const hit = resolveRef(doc, rangeText.trim(), sheet);
  if (!hit) return { error: 'Invalid data range.' };
  const def = chartFromRange(doc, hit.sheet, hit.range, type);
  const data = readChartData(doc, { ...def, id: '', z: 0, from: { r: 0, c: 0, dx: 0, dy: 0 }, to: { r: 0, c: 0, dx: 0, dy: 0 } });
  if (!data.series.some((s) => s.values.some((v) => v !== null))) return { error: 'The selected range has no numbers to plot.' };
  const layout = SheetLayout.of(sheet);
  const w = phone ? Math.min(CHART_W, Math.round(window.innerWidth * 0.8)) : CHART_W;
  const h = phone ? Math.round((w * 3) / 4) : CHART_H;
  const area = trimRange(hit.sheet, hit.range);
  // The first free column right of anything in the rows the chart will span.
  const spanRows = Math.ceil(h / 20) + 1;
  let lastCol = area.c1;
  for (let r = area.r0; r <= area.r0 + spanRows; r++) {
    const row = hit.sheet.cells.get(r);
    if (row) for (const [c, cell] of row) if (cell.v !== null && cell.v !== '' && c > lastCol) lastCol = c;
  }
  const at =
    hit.sheet !== sheet
      ? { x: layout.colX(doc.selection.active.c), y: layout.rowY(doc.selection.active.r) }
      : phone
        ? { x: layout.colX(area.c0), y: layout.rowY(area.r1 + 1) + 12 }
        : { x: layout.colX(lastCol + 1) + 16, y: layout.rowY(area.r0) };
  return { chart: doc.addChart({ ...def, name: `Chart ${sheet.charts.length + 1}` }, { x: at.x, y: at.y, width: w, height: h }) };
}

export const CHART_SIDES: { key: Side; label: string }[] = [
  { key: 'top', label: 'Top' },
  { key: 'bottom', label: 'Bottom' },
  { key: 'left', label: 'Left' },
  { key: 'right', label: 'Right' },
];

export const LINE_DASHES: { key: LineDash; label: string }[] = [
  { key: 'solid', label: 'Solid' },
  { key: 'dashed', label: 'Dashed' },
  { key: 'dotted', label: 'Dotted' },
  { key: 'dash-dot', label: 'Dash dot' },
];

export const TRENDLINES: { key: Trendline; label: string }[] = [
  { key: 'none', label: 'None' },
  { key: 'linear', label: 'Linear' },
  { key: 'exponential', label: 'Exponential' },
  { key: 'logarithmic', label: 'Logarithmic' },
  { key: 'polynomial', label: 'Polynomial' },
  { key: 'power', label: 'Power' },
  { key: 'moving-average', label: 'Moving average' },
];

/** The Configuration tab: chart type and the layout of its data. */
export interface ChartConfig {
  type: ChartType;
  seriesIn: 'rows' | 'columns';
  range: string;
  headers: boolean;
  labels: number | null;
  series: number[];
}

/** The Style tab (the axis limits as typed). */
export type ChartStyle = Pick<
  SheetChart,
  'title' | 'titlePosition' | 'subtitle' | 'legendTitle' | 'legendPosition' | 'showLegend' | 'gridlines' | 'valueLabels' | 'xAxisTitle' | 'yAxisTitle' | 'secondaryAxisTitle'
> & { yMin: string; yMax: string };

/** The Series tab, per series. */
export type SeriesStyle = Pick<ChartSeries, 'color' | 'lineWidth' | 'lineStyle' | 'trendline' | 'markers' | 'secondaryAxis'>;

export const numText = (n: number | undefined): string => (n === undefined ? '' : String(n));
export const parseNum = (s: string): number | undefined => (s.trim() === '' || !Number.isFinite(Number(s)) ? undefined : Number(s));

/**
 * The settings of a chart when the dialog opens. `custom`: its series do not follow a layout of
 * one range (they are listed, not edited).
 */
export function chartSettingsOf(doc: SpreadsheetDocument, chart: SheetChart | undefined, sheet: Sheet): { config: ChartConfig; custom: boolean; style: ChartStyle; seriesStyles: SeriesStyle[] } {
  const initialLayout = chart ? layoutOfChart(doc, chart) : null;
  return {
    config: {
      type: chart?.type ?? 'column',
      seriesIn: chart?.seriesIn ?? 'rows',
      range: localRef(chart?.range || (chart ? chartBounds(doc, chart) : '') || '', sheet),
      headers: chart?.headers ?? true,
      labels: initialLayout?.labels ?? null,
      series: initialLayout?.series ?? [],
    },
    custom: !initialLayout,
    style: {
      title: chart?.title ?? '',
      titlePosition: chart?.titlePosition,
      subtitle: chart?.subtitle ?? '',
      legendTitle: chart?.legendTitle ?? '',
      legendPosition: chart?.legendPosition,
      showLegend: chart?.showLegend ?? true,
      gridlines: chart?.gridlines ?? true,
      valueLabels: chart?.valueLabels ?? false,
      xAxisTitle: chart?.xAxisTitle ?? '',
      yAxisTitle: chart?.yAxisTitle ?? '',
      secondaryAxisTitle: chart?.secondaryAxisTitle ?? '',
      yMin: numText(chart?.yMin),
      yMax: numText(chart?.yMax),
    },
    seriesStyles: (chart?.series ?? []).map((s) => ({ color: s.color, lineWidth: s.lineWidth, lineStyle: s.lineStyle, trendline: s.trendline, markers: s.markers, secondaryAxis: s.secondaryAxis })),
  };
}

/**
 * The configuration with `patch`, its labels and series worked out again (when the range, the
 * orientation or the header option changes); `laidOut` when the range is valid, so the series follow it.
 */
export function relayoutChart(doc: SpreadsheetDocument, sheet: Sheet, config: ChartConfig, patch: Partial<ChartConfig>): { config: ChartConfig; laidOut: boolean } {
  const next = { ...config, ...patch };
  const hit = resolveRef(doc, next.range, sheet);
  if (!hit) return { config: next, laidOut: false };
  const area = trimRange(hit.sheet, hit.range);
  const guess = guessLayout(doc, hit.sheet, area, next.type);
  const count = next.seriesIn === 'columns' ? area.c1 - area.c0 + 1 : area.r1 - area.r0 + 1;
  const all = Array.from({ length: count }, (_, i) => i);
  const labels = next.seriesIn === guess.seriesIn ? guess.labels : count > 1 ? 0 : null;
  next.labels = labels;
  next.series = all.filter((i) => i !== labels);
  if (patch.headers === undefined && patch.seriesIn !== undefined) next.headers = next.seriesIn === guess.seriesIn ? guess.headers : next.headers;
  return { config: next, laidOut: true };
}

/** The changes to a chart from the dialog's settings (`target`: the configured range, when valid). */
export function chartSettingsPatch(
  chart: SheetChart,
  settings: { config: ChartConfig; custom: boolean; style: ChartStyle; seriesStyles: SeriesStyle[] },
  target: { sheet: Sheet; range: Range } | null,
): Partial<SheetChart> {
  const { config, custom, style, seriesStyles } = settings;
  const patch: Partial<SheetChart> = {
    type: config.type,
    title: style.title?.trim() || undefined,
    titlePosition: style.titlePosition,
    subtitle: style.subtitle?.trim() || undefined,
    legendTitle: style.legendTitle?.trim() || undefined,
    legendPosition: style.legendPosition,
    showLegend: style.showLegend,
    gridlines: style.gridlines,
    valueLabels: style.valueLabels,
    xAxisTitle: style.xAxisTitle?.trim() || undefined,
    yAxisTitle: style.yAxisTitle?.trim() || undefined,
    secondaryAxisTitle: style.secondaryAxisTitle?.trim() || undefined,
    yMin: parseNum(style.yMin),
    yMax: parseNum(style.yMax),
  };
  let series = chart.series;
  if (!custom && target) {
    const area = trimRange(target.sheet, target.range);
    const layout: ChartLayout = { seriesIn: config.seriesIn, headers: config.headers, labels: config.labels, series: config.series };
    const built = seriesForLayout(target.sheet, area, layout, config.type, chart.series);
    series = built.series;
    patch.categories = built.categories;
    patch.range = absRef(target.sheet, area);
    patch.seriesIn = config.seriesIn;
    patch.headers = config.headers;
  }
  patch.series = series.map((s, i) => {
    const st = seriesStyles[i] ?? {};
    return { ...s, color: st.color || undefined, lineWidth: st.lineWidth, lineStyle: st.lineStyle, trendline: st.trendline, markers: st.markers, secondaryAxis: st.secondaryAxis };
  });
  return patch;
}

/** The series styles with series `i` changed. */
export function withSeriesStyle(list: SeriesStyle[], i: number, p: Partial<SeriesStyle>): SeriesStyle[] {
  const out = list.slice();
  while (out.length <= i) out.push({});
  out[i] = { ...out[i], ...p };
  return out;
}
