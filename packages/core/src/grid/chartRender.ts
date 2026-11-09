/**
 * SVG rendering of charts from a chart definition and the current values behind its references.
 * The result is markup, so the same drawing serves the sheet, the print preview and the PDF export.
 */
import { formatAxisValue, type ChartData, type ChartDataSeries } from '../core/chartData';
import type { LineDash, SheetChart, Side, Trendline } from '../core/model/drawing';

/** Series colours of charts created in the editor. */
export const CHART_COLORS = ['#1E88E5', '#E53935', '#43A047', '#FDD835', '#8E24AA', '#00897B', '#5D4037', '#FB8C00', '#3949AB', '#C0CA33', '#D81B60', '#6D4C41'];

const FONT = 'Arimo, Arial, Helvetica, sans-serif';
const TEXT = '#333333';
const AXIS_TEXT = '#555555';
const GRID = '#e6e6e6';
const BASELINE = '#bdbdbd';

interface Box {
  x: number;
  y: number;
  w: number;
  h: number;
}

interface Scale {
  min: number;
  max: number;
  step: number;
  ticks: number[];
}

interface Ctx {
  chart: SheetChart;
  data: ChartData;
  w: number;
  h: number;
  out: string[];
  colors: string[];
}

const esc = (s: string) => s.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
const r1 = (n: number) => String(Math.round(n * 10) / 10);

/** Approximate text width (no layout engine is needed to draw the chart). */
export function textWidth(s: string, size: number, bold = false): number {
  let w = 0;
  for (const ch of s) w += /[A-Z0-9MW@%#&]/.test(ch) ? 0.64 : /[iljtf.,:;'!| ]/.test(ch) ? 0.3 : /[　-鿿가-힯]/.test(ch) ? 1 : 0.52;
  return w * size * (bold ? 1.08 : 1);
}

function clipText(s: string, max: number, size: number, bold = false): string {
  if (textWidth(s, size, bold) <= max) return s;
  let out = s;
  while (out.length > 1 && textWidth(`${out}…`, size, bold) > max) out = out.slice(0, -1);
  return out.length <= 1 ? (max > size ? `${s[0]}…` : '') : `${out}…`;
}

interface TextOpts {
  size?: number;
  weight?: number;
  anchor?: 'start' | 'middle' | 'end';
  fill?: string;
  rotate?: number;
  middle?: boolean;
}

function text(x: number, y: number, s: string, o: TextOpts = {}): string {
  if (!s) return '';
  const attrs = [`x="${r1(x)}"`, `y="${r1(y)}"`, `font-size="${o.size ?? 11}"`, `fill="${o.fill ?? AXIS_TEXT}"`];
  if (o.weight) attrs.push(`font-weight="${o.weight}"`);
  if (o.anchor && o.anchor !== 'start') attrs.push(`text-anchor="${o.anchor}"`);
  if (o.middle) attrs.push('dominant-baseline="central"');
  if (o.rotate) attrs.push(`transform="rotate(${o.rotate} ${r1(x)} ${r1(y)})"`);
  return `<text ${attrs.join(' ')}>${esc(s)}</text>`;
}

function line(x1: number, y1: number, x2: number, y2: number, stroke: string, width = 1, dash?: string): string {
  return `<line x1="${r1(x1)}" y1="${r1(y1)}" x2="${r1(x2)}" y2="${r1(y2)}" stroke="${stroke}" stroke-width="${width}"${dash ? ` stroke-dasharray="${dash}"` : ''}/>`;
}

function rect(x: number, y: number, w: number, h: number, fill: string, extra = ''): string {
  return `<rect x="${r1(x)}" y="${r1(y)}" width="${r1(Math.max(0, w))}" height="${r1(Math.max(0, h))}" fill="${fill}"${extra}/>`;
}

function tip(title: string): string {
  return title ? `<title>${esc(title)}</title>` : '';
}

function dashArray(d: LineDash | undefined, width: number): string | undefined {
  switch (d) {
    case 'dashed':
      return `${width * 3} ${width * 2}`;
    case 'dotted':
      return `${width} ${width * 1.5}`;
    case 'dash-dot':
      return `${width * 3} ${width * 1.5} ${width} ${width * 1.5}`;
    default:
      return undefined;
  }
}

export function niceScale(min: number, max: number, target: number, fixedMin?: number, fixedMax?: number): Scale {
  if (fixedMin !== undefined && Number.isFinite(fixedMin)) min = fixedMin;
  if (fixedMax !== undefined && Number.isFinite(fixedMax)) max = fixedMax;
  if (!Number.isFinite(min) || !Number.isFinite(max)) {
    min = 0;
    max = 1;
  }
  if (max === min) {
    if (max === 0) max = 1;
    else if (max > 0) min = 0;
    else max = 0;
  }
  if (max < min) [min, max] = [max, min];
  const raw = (max - min) / Math.max(1, target);
  const mag = 10 ** Math.floor(Math.log10(raw));
  const norm = raw / mag;
  const step = (norm <= 1 ? 1 : norm <= 2 ? 2 : norm <= 2.5 ? 2.5 : norm <= 5 ? 5 : 10) * mag;
  const lo = fixedMin !== undefined && Number.isFinite(fixedMin) ? fixedMin : Math.floor(min / step + 1e-9) * step;
  let hi = fixedMax !== undefined && Number.isFinite(fixedMax) ? fixedMax : Math.ceil(max / step - 1e-9) * step;
  if (hi <= lo) hi = lo + step;
  const ticks: number[] = [];
  for (let k = 0; k < 200; k++) {
    const v = Number((lo + k * step).toPrecision(12));
    if (v > hi + step * 1e-6) break;
    ticks.push(v);
  }
  return { min: lo, max: hi, step, ticks };
}

// -------------------------------------------------------------------------------------------------
// Statistics used by the statistical chart types

export function quartiles(values: number[]): { min: number; q1: number; median: number; q3: number; max: number; mean: number; lo: number; hi: number; outliers: number[] } | null {
  const v = values.filter((x) => Number.isFinite(x)).sort((a, b) => a - b);
  if (!v.length) return null;
  // Exclusive method (Excel's default for box and whisker charts).
  const at = (p: number) => {
    const pos = (v.length + 1) * p;
    if (pos <= 1) return v[0];
    if (pos >= v.length) return v[v.length - 1];
    const i = Math.floor(pos);
    return v[i - 1] + (pos - i) * (v[i] - v[i - 1]);
  };
  const q1 = at(0.25);
  const median = at(0.5);
  const q3 = at(0.75);
  const iqr = q3 - q1;
  const inside = v.filter((x) => x >= q1 - 1.5 * iqr && x <= q3 + 1.5 * iqr);
  const mean = v.reduce((a, b) => a + b, 0) / v.length;
  return {
    min: v[0],
    max: v[v.length - 1],
    q1,
    median,
    q3,
    mean,
    lo: inside.length ? inside[0] : q1,
    hi: inside.length ? inside[inside.length - 1] : q3,
    outliers: v.filter((x) => x < q1 - 1.5 * iqr || x > q3 + 1.5 * iqr),
  };
}

/** Histogram bins (Scott's rule, like Excel's automatic bins). */
export function histogramBins(values: number[]): { from: number; to: number; count: number }[] {
  const v = values.filter((x) => Number.isFinite(x));
  if (!v.length) return [];
  const min = Math.min(...v);
  const max = Math.max(...v);
  if (min === max) return [{ from: min, to: max, count: v.length }];
  const mean = v.reduce((a, b) => a + b, 0) / v.length;
  const sd = Math.sqrt(v.reduce((a, b) => a + (b - mean) ** 2, 0) / Math.max(1, v.length - 1));
  let width = (3.5 * sd) / Math.cbrt(v.length);
  if (!(width > 0)) width = (max - min) / 5;
  let count = Math.max(1, Math.min(40, Math.ceil((max - min) / width)));
  if (count === 1 && v.length > 1) count = 2;
  width = (max - min) / count;
  const bins = Array.from({ length: count }, (_, i) => ({ from: min + i * width, to: i === count - 1 ? max : min + (i + 1) * width, count: 0 }));
  for (const x of v) {
    // Bins are (from, to], the first one includes its lower bound.
    let i = Math.ceil((x - min) / width) - 1;
    if (i < 0) i = 0;
    if (i >= count) i = count - 1;
    bins[i].count++;
  }
  return bins;
}

/** A fitted trendline, or null when the data does not allow one. */
export function fitTrend(kind: Trendline, xs: number[], ys: number[]): ((x: number) => number) | null {
  const pts = xs.map((x, i) => [x, ys[i]] as const).filter(([x, y]) => Number.isFinite(x) && Number.isFinite(y));
  const linear = (p: readonly (readonly [number, number])[]) => {
    const n = p.length;
    if (n < 2) return null;
    let sx = 0;
    let sy = 0;
    let sxx = 0;
    let sxy = 0;
    for (const [x, y] of p) {
      sx += x;
      sy += y;
      sxx += x * x;
      sxy += x * y;
    }
    const d = n * sxx - sx * sx;
    if (!d) return null;
    const b = (n * sxy - sx * sy) / d;
    return { a: (sy - b * sx) / n, b };
  };
  switch (kind) {
    case 'linear': {
      const f = linear(pts);
      return f ? (x) => f.a + f.b * x : null;
    }
    case 'exponential': {
      const f = linear(pts.filter(([, y]) => y > 0).map(([x, y]) => [x, Math.log(y)] as const));
      return f ? (x) => Math.exp(f.a + f.b * x) : null;
    }
    case 'logarithmic': {
      const f = linear(pts.filter(([x]) => x > 0).map(([x, y]) => [Math.log(x), y] as const));
      return f ? (x) => (x > 0 ? f.a + f.b * Math.log(x) : NaN) : null;
    }
    case 'power': {
      const f = linear(pts.filter(([x, y]) => x > 0 && y > 0).map(([x, y]) => [Math.log(x), Math.log(y)] as const));
      return f ? (x) => (x > 0 ? Math.exp(f.a) * x ** f.b : NaN) : null;
    }
    case 'polynomial': {
      // Order 2 least squares through the normal equations.
      if (pts.length < 3) return null;
      const s = [0, 0, 0, 0, 0];
      const t = [0, 0, 0];
      for (const [x, y] of pts) {
        for (let k = 0; k < 5; k++) s[k] += x ** k;
        for (let k = 0; k < 3; k++) t[k] += y * x ** k;
      }
      const m = [
        [s[0], s[1], s[2], t[0]],
        [s[1], s[2], s[3], t[1]],
        [s[2], s[3], s[4], t[2]],
      ];
      for (let i = 0; i < 3; i++) {
        let p = i;
        for (let k = i + 1; k < 3; k++) if (Math.abs(m[k][i]) > Math.abs(m[p][i])) p = k;
        [m[i], m[p]] = [m[p], m[i]];
        if (Math.abs(m[i][i]) < 1e-12) return null;
        for (let k = 0; k < 3; k++) {
          if (k === i) continue;
          const f = m[k][i] / m[i][i];
          for (let j = i; j < 4; j++) m[k][j] -= f * m[i][j];
        }
      }
      const c = [m[0][3] / m[0][0], m[1][3] / m[1][1], m[2][3] / m[2][2]];
      return (x) => c[0] + c[1] * x + c[2] * x * x;
    }
    default:
      return null;
  }
}

/** Squarified treemap layout of positive weights inside a box (order of the input kept in the output). */
export function squarify(weights: number[], box: Box): Box[] {
  const out: Box[] = new Array(weights.length);
  const items = weights.map((w, i) => ({ w: Math.max(0, w), i })).filter((x) => x.w > 0).sort((a, b) => b.w - a.w);
  const total = items.reduce((a, b) => a + b.w, 0);
  for (let i = 0; i < weights.length; i++) out[i] = { x: box.x, y: box.y, w: 0, h: 0 };
  if (!total || box.w <= 0 || box.h <= 0) return out;
  const scale = (box.w * box.h) / total;
  let rest = { ...box };
  let row: { w: number; i: number }[] = [];
  const worst = (r: { w: number }[], side: number) => {
    const areas = r.map((x) => x.w * scale);
    const sum = areas.reduce((a, b) => a + b, 0);
    const mx = Math.max(...areas);
    const mn = Math.min(...areas);
    return Math.max((side * side * mx) / (sum * sum), (sum * sum) / (side * side * mn));
  };
  const layoutRow = (r: { w: number; i: number }[]) => {
    const sum = r.reduce((a, b) => a + b.w * scale, 0);
    if (rest.w >= rest.h) {
      const width = sum / rest.h;
      let y = rest.y;
      for (const it of r) {
        const h = (it.w * scale) / width;
        out[it.i] = { x: rest.x, y, w: width, h };
        y += h;
      }
      rest = { x: rest.x + width, y: rest.y, w: rest.w - width, h: rest.h };
    } else {
      const height = sum / rest.w;
      let x = rest.x;
      for (const it of r) {
        const w = (it.w * scale) / height;
        out[it.i] = { x, y: rest.y, w, h: height };
        x += w;
      }
      rest = { x: rest.x, y: rest.y + height, w: rest.w, h: rest.h - height };
    }
  };
  for (const it of items) {
    const side = Math.min(rest.w, rest.h);
    if (!row.length || worst([...row, it], side) <= worst(row, side)) row.push(it);
    else {
      layoutRow(row);
      row = [it];
    }
  }
  if (row.length) layoutRow(row);
  return out;
}

// -------------------------------------------------------------------------------------------------
// Frame: title, subtitle, legend and axis titles around the plot area

interface LegendItem {
  name: string;
  color: string;
  line?: boolean;
}

function legendItemsOf(ctx: Ctx): LegendItem[] {
  const { chart, data, colors } = ctx;
  switch (chart.type) {
    case 'pie':
    case 'doughnut':
    case 'treemap':
      return data.categories.map((c, i) => ({ name: c, color: colors[i % colors.length] }));
    case 'funnel':
      return data.series.slice(0, 1).map((s) => ({ name: s.name, color: seriesColor(ctx, s, 0) }));
    case 'waterfall':
      return [
        { name: 'Increase', color: colors[0] },
        { name: 'Decrease', color: colors[1] },
        { name: 'Total', color: colors[2] },
      ];
    case 'histogram':
      return data.series.slice(0, 1).map((s) => ({ name: s.name, color: seriesColor(ctx, s, 0) }));
    case 'pareto':
      return data.series.length ? [{ name: data.series[0].name, color: seriesColor(ctx, data.series[0], 0) }, { name: 'Cumulative %', color: colors[1], line: true }] : [];
    case 'stock':
      return [];
    default:
      return data.series.map((s, i) => ({ name: s.name, color: seriesColor(ctx, s, i), line: chart.type === 'line' || chart.type === 'radar' }));
  }
}

function seriesColor(ctx: Ctx, s: ChartDataSeries | undefined, i: number): string {
  return s?.source.color || ctx.colors[i % ctx.colors.length];
}

function drawFrame(ctx: Ctx): Box {
  const { chart, w, h, out } = ctx;
  let top = 10;
  let bottom = h - 10;
  let left = 10;
  let right = w - 10;
  const title = (chart.title ?? '').trim();
  const subtitle = (chart.subtitle ?? '').trim();
  const pos: Side = chart.titlePosition ?? 'top';
  if (title || subtitle) {
    const th = (title ? 22 : 0) + (subtitle ? 17 : 0);
    if (pos === 'top' || pos === 'bottom') {
      let y = pos === 'top' ? top : bottom - th;
      if (title) {
        out.push(text(w / 2, y + 15, clipText(title, w - 20, 15, true), { size: 15, weight: 600, anchor: 'middle', fill: TEXT }));
        y += 22;
      }
      if (subtitle) out.push(text(w / 2, y + 12, clipText(subtitle, w - 20, 12), { size: 12, anchor: 'middle', fill: '#757575' }));
      if (pos === 'top') top += th + 4;
      else bottom -= th + 4;
    } else {
      const x = pos === 'left' ? left + 8 : right - 8;
      const rot = pos === 'left' ? -90 : 90;
      let off = 0;
      if (title) {
        out.push(text(x + (pos === 'left' ? 8 : -8), h / 2, clipText(title, h - 20, 15, true), { size: 15, weight: 600, anchor: 'middle', fill: TEXT, rotate: rot }));
        off += 22;
      }
      if (subtitle) out.push(text(x + (pos === 'left' ? off + 6 : -off - 6), h / 2, clipText(subtitle, h - 20, 12), { size: 12, anchor: 'middle', fill: '#757575', rotate: rot }));
      if (pos === 'left') left += th + 6;
      else right -= th + 6;
    }
  }
  const items = chart.showLegend ? legendItemsOf(ctx) : [];
  const legendTitle = (chart.legendTitle ?? '').trim();
  if (items.length) {
    const lp: Side = chart.legendPosition ?? 'bottom';
    const size = 12;
    const itemW = (it: LegendItem) => 16 + Math.min(150, textWidth(it.name, size, true)) + 18;
    if (lp === 'top' || lp === 'bottom') {
      const avail = right - left;
      const rows: { items: LegendItem[]; width: number }[] = [{ items: [], width: legendTitle ? textWidth(legendTitle, size, true) + 14 : 0 }];
      for (const it of items) {
        const iw = itemW(it);
        const cur = rows[rows.length - 1];
        if (cur.items.length && cur.width + iw > avail) rows.push({ items: [it], width: iw });
        else {
          cur.items.push(it);
          cur.width += iw;
        }
      }
      const shown = rows.slice(0, 3);
      const lh = shown.length * 20;
      let y = lp === 'top' ? top + 10 : bottom - lh + 10;
      shown.forEach((row, ri) => {
        let x = left + Math.max(0, (avail - row.width + 18) / 2);
        if (ri === 0 && legendTitle) {
          out.push(text(x, y + 4, legendTitle, { size, weight: 700, fill: TEXT }));
          x += textWidth(legendTitle, size, true) + 14;
        }
        for (const it of row.items) {
          out.push(legendMarker(x + 5, y, it));
          out.push(text(x + 14, y + 4, clipText(it.name, 150, size, true), { size, weight: 700, fill: TEXT }));
          x += itemW(it);
        }
        y += 20;
      });
      if (lp === 'top') top += lh + 6;
      else bottom -= lh + 6;
    } else {
      const lw = Math.min((right - left) * 0.35, Math.max(...items.map((it) => itemW(it) - 18), legendTitle ? textWidth(legendTitle, size, true) : 0) + 8);
      const maxRows = Math.max(1, Math.floor((bottom - top) / 20) - (legendTitle ? 1 : 0));
      const shown = items.slice(0, maxRows);
      const lh = (shown.length + (legendTitle ? 1 : 0)) * 20;
      let y = top + Math.max(0, (bottom - top - lh) / 2) + 10;
      const x0 = lp === 'left' ? left + 4 : right - lw;
      if (legendTitle) {
        out.push(text(x0, y + 4, clipText(legendTitle, lw, size, true), { size, weight: 700, fill: TEXT }));
        y += 20;
      }
      for (const it of shown) {
        out.push(legendMarker(x0 + 5, y, it));
        out.push(text(x0 + 14, y + 4, clipText(it.name, lw - 16, size, true), { size, weight: 700, fill: TEXT }));
        y += 20;
      }
      if (lp === 'left') left += lw + 10;
      else right -= lw + 10;
    }
  }
  return { x: left, y: top, w: Math.max(10, right - left), h: Math.max(10, bottom - top) };
}

function legendMarker(x: number, y: number, it: LegendItem): string {
  if (it.line) return `${line(x - 6, y, x + 6, y, it.color, 2)}<circle cx="${r1(x)}" cy="${r1(y)}" r="3.5" fill="${it.color}"/>`;
  return `<circle cx="${r1(x)}" cy="${r1(y)}" r="5" fill="${it.color}"/>`;
}

// -------------------------------------------------------------------------------------------------
// Category charts: column, bar, line, area and their stacked forms, waterfall, histogram, pareto,
// box and whisker, stock

interface Mark {
  /** Category index (band). */
  i: number;
  series: number;
  from: number;
  to: number;
  color: string;
  label?: string;
  title: string;
}

function cartesian(ctx: Ctx, plot: Box): void {
  const { chart, data, out } = ctx;
  const type = chart.type;
  const horizontal = type === 'bar' || type === 'stacked-bar' || type === 'percent-bar';
  const stacked = type.startsWith('stacked-') || type.startsWith('percent-');
  const percent = type.startsWith('percent-');
  const isArea = type === 'area' || type === 'stacked-area' || type === 'percent-area';
  const isLine = type === 'line';
  const isColumnLike = !isArea && !isLine && type !== 'stock' && type !== 'box-whisker';
  const numFmt = percent ? '0%' : data.numFmt;

  // Categories and plotted series after the type-specific transformations.
  let categories = data.categories.slice();
  let series = data.series;
  let secondaryLine: { values: number[]; color: string; name: string } | null = null;
  const marks: Mark[] = [];
  const lines: { s: ChartDataSeries; color: string; pts: { i: number; v: number }[]; secondary: boolean }[] = [];
  let lo = 0;
  let hi = 0;
  const grow = (v: number) => {
    if (!Number.isFinite(v)) return;
    lo = Math.min(lo, v);
    hi = Math.max(hi, v);
  };
  let gapless = false;
  let boxes: { i: number; q: NonNullable<ReturnType<typeof quartiles>>; color: string; name: string }[] = [];
  let candles: { i: number; open?: number; high: number; low: number; close: number }[] = [];
  const secondaryIdx = new Set(series.filter((s) => s.source.secondaryAxis).map((s) => s.index));
  let slo = 0;
  let shi = 0;

  if (type === 'histogram') {
    const s = series[0];
    const all = (s?.values ?? []).filter((v): v is number => v !== null);
    const bins = histogramBins(all);
    const dec = (n: number) => String(Number(n.toPrecision(4)));
    categories = bins.map((b, i) => `${i === 0 ? '[' : '('}${dec(b.from)}, ${dec(b.to)}]`);
    bins.forEach((b, i) => {
      marks.push({ i, series: 0, from: 0, to: b.count, color: seriesColor(ctx, s, 0), label: String(b.count), title: `${categories[i]}: ${b.count}` });
      grow(b.count);
    });
    gapless = true;
  } else if (type === 'pareto') {
    const s = series[0];
    const items = categories.map((c, i) => ({ c, v: Math.max(0, s?.values[i] ?? 0) })).filter((x) => x.v > 0 || x.c);
    items.sort((a, b) => b.v - a.v);
    categories = items.map((x) => x.c);
    const total = items.reduce((a, b) => a + b.v, 0) || 1;
    let run = 0;
    const cum: number[] = [];
    items.forEach((x, i) => {
      marks.push({ i, series: 0, from: 0, to: x.v, color: seriesColor(ctx, s, 0), label: formatAxisValue(x.v, numFmt), title: `${x.c}: ${formatAxisValue(x.v, numFmt)}` });
      grow(x.v);
      run += x.v;
      cum.push((run / total) * 100);
    });
    secondaryLine = { values: cum, color: ctx.colors[1], name: 'Cumulative %' };
    slo = 0;
    shi = 100;
    gapless = true;
  } else if (type === 'waterfall') {
    const s = series[0];
    let run = 0;
    categories.forEach((c, i) => {
      const v = s?.values[i] ?? 0;
      if (v === null) return;
      const total = /^(total|subtotal|sum|net|balance|grand total)\b/i.test(c.trim()) && i > 0;
      let from: number;
      let to: number;
      let color: string;
      if (total) {
        from = 0;
        to = v;
        run = v;
        color = ctx.colors[2];
      } else {
        from = run;
        to = run + v;
        run = to;
        color = v >= 0 ? ctx.colors[0] : ctx.colors[1];
      }
      grow(from);
      grow(to);
      marks.push({ i, series: 0, from, to, color, label: formatAxisValue(v, numFmt), title: `${c}: ${formatAxisValue(v, numFmt)}` });
    });
  } else if (type === 'box-whisker') {
    boxes = series
      .map((s, i) => ({ i, q: quartiles(s.values.filter((v): v is number => v !== null)), color: seriesColor(ctx, s, i), name: s.name }))
      .filter((b): b is { i: number; q: NonNullable<ReturnType<typeof quartiles>>; color: string; name: string } => b.q !== null);
    categories = series.map((s) => s.name);
    lo = Infinity;
    hi = -Infinity;
    for (const b of boxes) {
      grow(b.q.min);
      grow(b.q.max);
    }
    if (!Number.isFinite(lo)) {
      lo = 0;
      hi = 1;
    }
  } else if (type === 'stock') {
    const get = (k: number, i: number) => series[k]?.values[i] ?? null;
    const four = series.length >= 4;
    categories.forEach((_, i) => {
      const high = get(four ? 1 : 0, i);
      const low = get(four ? 2 : 1, i);
      const close = get(four ? 3 : 2, i);
      const open = four ? get(0, i) : null;
      if (high === null || low === null || close === null) return;
      candles.push({ i, open: open ?? undefined, high, low, close });
    });
    lo = Infinity;
    hi = -Infinity;
    for (const c of candles) {
      grow(c.low);
      grow(c.high);
    }
    if (!Number.isFinite(lo)) {
      lo = 0;
      hi = 1;
    } else {
      // Prices: the axis does not need to start at zero.
      const pad = (hi - lo) * 0.1 || 1;
      lo = Math.max(lo > 0 ? 0 : -Infinity, lo - pad);
      hi += pad;
    }
  } else {
    const n = Math.max(categories.length, ...series.map((s) => s.values.length));
    while (categories.length < n) categories.push(String(categories.length + 1));
    for (let i = 0; i < n; i++) {
      let pos = 0;
      let neg = 0;
      let total = 0;
      if (percent) for (const s of series) total += Math.abs(s.values[i] ?? 0);
      for (const s of series) {
        const raw = s.values[i];
        if (raw === null || raw === undefined) continue;
        const secondary = secondaryIdx.has(s.index) && !stacked;
        let v = raw;
        if (percent) v = total ? (raw / total) * 100 : 0;
        let from = 0;
        let to = v;
        if (stacked) {
          if (v >= 0) {
            from = pos;
            pos += v;
            to = pos;
          } else {
            from = neg;
            neg += v;
            to = neg;
          }
        }
        if (secondary) {
          slo = Math.min(slo, from, to);
          shi = Math.max(shi, from, to);
        } else {
          grow(from);
          grow(to);
        }
        const label = percent ? `${Math.round(v)}%` : formatAxisValue(raw, data.numFmt);
        const title = `${categories[i] ?? ''}\n${s.name}: ${percent ? `${Number(v.toFixed(1))}%` : formatAxisValue(raw, data.numFmt)}`;
        if (isColumnLike) marks.push({ i, series: s.index, from, to, color: seriesColor(ctx, s, s.index), label, title });
      }
    }
    if (isLine || isArea) {
      series.forEach((s) => {
        const pts: { i: number; v: number }[] = [];
        for (let i = 0; i < n; i++) {
          const raw = s.values[i];
          if (raw === null || raw === undefined) {
            if (stacked) pts.push({ i, v: NaN });
            continue;
          }
          pts.push({ i, v: raw });
        }
        lines.push({ s, color: seriesColor(ctx, s, s.index), pts, secondary: secondaryIdx.has(s.index) && !stacked });
      });
      if (stacked) {
        // Accumulate the stacked values (percent: shares of the category total).
        for (let i = 0; i < n; i++) {
          let run = 0;
          let total = 0;
          for (const l of lines) total += Math.abs(l.pts.find((p) => p.i === i)?.v || 0);
          for (const l of lines) {
            const p = l.pts.find((q) => q.i === i);
            if (!p) continue;
            const v = Number.isFinite(p.v) ? p.v : 0;
            run += percent ? (total ? (v / total) * 100 : 0) : v;
            p.v = run;
            grow(run);
          }
        }
      }
    }
    series = series.filter((s) => s.values.length);
  }
  if (percent) {
    lo = Math.min(0, lo);
    hi = 100;
  }

  const n = Math.max(1, categories.length);
  const labelSize = 11;
  const fixedMin = chart.yMin;
  const fixedMax = chart.yMax;
  const axisLen0 = horizontal ? plot.w : plot.h;
  const scale = niceScale(lo, hi, Math.max(2, Math.min(10, Math.floor(axisLen0 / 50))), percent ? 0 : fixedMin, percent ? 100 : fixedMax);
  const hasSecondary = !!secondaryLine || (secondaryIdx.size > 0 && !stacked);
  const sscale = hasSecondary ? niceScale(slo, shi, Math.max(2, Math.min(10, Math.floor(axisLen0 / 50))), type === 'pareto' ? 0 : undefined, type === 'pareto' ? 100 : undefined) : null;
  const fmtTick = (v: number) => (percent ? `${Number(v.toPrecision(6))}%` : formatAxisValue(v, numFmt));
  const fmtSec = (v: number) => (type === 'pareto' ? `${Number(v.toPrecision(6))}%` : formatAxisValue(v, numFmt));

  // Axis titles.
  const xTitle = (chart.xAxisTitle ?? '').trim();
  const yTitle = (chart.yAxisTitle ?? '').trim();
  const sTitle = (chart.secondaryAxisTitle ?? '').trim();
  let px0 = plot.x;
  let px1 = plot.x + plot.w;
  let py0 = plot.y + 6;
  let py1 = plot.y + plot.h;
  if (xTitle) {
    out.push(text(plot.x + plot.w / 2, py1 - 3, clipText(xTitle, plot.w, 12, true), { size: 12, weight: 600, anchor: 'middle', fill: TEXT }));
    py1 -= 20;
  }
  if (yTitle) {
    out.push(text(plot.x + 9, plot.y + plot.h / 2, clipText(yTitle, plot.h, 12, true), { size: 12, weight: 600, anchor: 'middle', fill: TEXT, rotate: -90 }));
    px0 += 20;
  }
  if (sTitle && hasSecondary) {
    out.push(text(plot.x + plot.w - 9, plot.y + plot.h / 2, clipText(sTitle, plot.h, 12, true), { size: 12, weight: 600, anchor: 'middle', fill: TEXT, rotate: 90 }));
    px1 -= 20;
  }
  // Tick label space.
  const tickLabels = scale.ticks.map(fmtTick);
  const catLabelW = Math.min(plot.w * 0.35, Math.max(...categories.map((c) => textWidth(c, labelSize))) + 10);
  if (horizontal) {
    px0 += catLabelW;
    py1 -= 18;
    px1 -= Math.min(30, textWidth(tickLabels[tickLabels.length - 1] ?? '', labelSize) / 2);
  } else {
    px0 += Math.max(...tickLabels.map((t) => textWidth(t, labelSize))) + 10;
    py1 -= 20;
    if (sscale) px1 -= Math.max(...sscale.ticks.map((t) => textWidth(fmtSec(t), labelSize))) + 10;
    else px1 -= 6;
  }
  const pw = Math.max(10, px1 - px0);
  const ph = Math.max(10, py1 - py0);
  const valueLen = horizontal ? pw : ph;
  const toPos = (v: number, sc: Scale = scale) => ((Math.min(sc.max, Math.max(sc.min, v)) - sc.min) / (sc.max - sc.min)) * valueLen;
  // Screen coordinate of a value along the value axis.
  const vp = (v: number, sc: Scale = scale) => (horizontal ? px0 + toPos(v, sc) : py1 - toPos(v, sc));
  const band = (horizontal ? ph : pw) / n;
  // Areas run from edge to edge; other types use the middle of each band.
  const areaStep = n > 1 ? pw / (n - 1) : pw;
  const cp = (i: number) => (isArea ? px0 + (n > 1 ? i * areaStep : pw / 2) : horizontal ? py0 + band * (i + 0.5) : px0 + band * (i + 0.5));

  // Gridlines and value labels.
  scale.ticks.forEach((t, k) => {
    const p = vp(t);
    if (horizontal) {
      if (chart.gridlines) out.push(line(p, py0, p, py1, GRID));
      out.push(text(p, py1 + 14, tickLabels[k], { size: labelSize, anchor: 'middle' }));
    } else {
      if (chart.gridlines) out.push(line(px0, p, px1, p, GRID));
      out.push(text(px0 - 6, p + 4, tickLabels[k], { size: labelSize, anchor: 'end' }));
    }
  });
  if (sscale && !horizontal) {
    for (const t of sscale.ticks) out.push(text(px1 + 6, vp(t, sscale) + 4, fmtSec(t), { size: labelSize }));
  }
  // Category labels (thinned out when crowded).
  const space = horizontal ? 16 : Math.max(...categories.map((c) => Math.min(textWidth(c, labelSize), 90))) + 8;
  const every = Math.max(1, Math.ceil(space / Math.max(1, horizontal ? band : isArea ? areaStep : band)));
  categories.forEach((c, i) => {
    if (i % every) return;
    if (horizontal) out.push(text(px0 - 6, cp(i) + 4, clipText(c, catLabelW - 8, labelSize), { size: labelSize, anchor: 'end' }));
    else out.push(text(cp(i), py1 + 15, clipText(c, Math.max(band * every - 4, 20), labelSize), { size: labelSize, anchor: 'middle' }));
  });

  // Columns and bars.
  if (marks.length) {
    const groups = stacked || type === 'waterfall' || type === 'histogram' || type === 'pareto' ? 1 : Math.max(1, series.length);
    const inner = gapless ? band * 0.96 : band * (groups > 1 ? 0.74 : 0.6);
    const per = inner / groups;
    const gap = groups > 1 ? Math.min(per * 0.18, 8) : 0;
    const seriesSlot = new Map(series.map((s, k) => [s.index, k]));
    for (const m of marks) {
      const slot = groups === 1 ? 0 : seriesSlot.get(m.series) ?? 0;
      const sc = secondaryIdx.has(m.series) && !stacked && sscale ? sscale : scale;
      const a = vp(m.from, sc);
      const b = vp(m.to, sc);
      const start = (horizontal ? py0 : px0) + band * m.i + (band - inner) / 2 + per * slot + gap / 2;
      const size = Math.max(1, per - gap);
      if (horizontal) out.push(`${rect(Math.min(a, b), start, Math.abs(b - a), size, m.color)}`.replace('/>', `>${tip(m.title)}</rect>`));
      else out.push(`${rect(start, Math.min(a, b), size, Math.abs(b - a), m.color)}`.replace('/>', `>${tip(m.title)}</rect>`));
      if (chart.valueLabels && m.label) {
        if (horizontal) out.push(text(Math.max(a, b) + 4, start + size / 2 + 4, m.label, { size: 10, fill: TEXT }));
        else if (stacked) out.push(text(start + size / 2, (a + b) / 2 + 4, m.label, { size: 10, anchor: 'middle', fill: '#ffffff' }));
        else out.push(text(start + size / 2, (m.to >= m.from ? Math.min(a, b) : Math.max(a, b) + 12) - 4, m.label, { size: 10, anchor: 'middle', fill: TEXT }));
      }
    }
    if (type === 'waterfall') {
      for (let k = 0; k + 1 < marks.length; k++) {
        const y = vp(marks[k].to);
        const x0 = px0 + band * marks[k].i + (band - inner) / 2 + inner;
        const x1 = px0 + band * marks[k + 1].i + (band - inner) / 2;
        out.push(line(x0, y, x1, y, '#9e9e9e', 1, '3 2'));
      }
    }
  }

  // Box and whisker.
  for (const b of boxes) {
    const cx = px0 + band * (b.i + 0.5);
    const bw = Math.min(band * 0.5, 60);
    const yq1 = vp(b.q.q1);
    const yq3 = vp(b.q.q3);
    out.push(line(cx, vp(b.q.lo), cx, yq1, '#616161'), line(cx, yq3, cx, vp(b.q.hi), '#616161'));
    out.push(line(cx - bw / 4, vp(b.q.lo), cx + bw / 4, vp(b.q.lo), '#616161'), line(cx - bw / 4, vp(b.q.hi), cx + bw / 4, vp(b.q.hi), '#616161'));
    out.push(rect(cx - bw / 2, Math.min(yq1, yq3), bw, Math.abs(yq1 - yq3), b.color, ` stroke="#424242" stroke-width="0.8"`).replace('/>', `>${tip(`${b.name}\nQ1 ${formatAxisValue(b.q.q1, numFmt)}, median ${formatAxisValue(b.q.median, numFmt)}, Q3 ${formatAxisValue(b.q.q3, numFmt)}`)}</rect>`));
    out.push(line(cx - bw / 2, vp(b.q.median), cx + bw / 2, vp(b.q.median), '#ffffff', 2));
    const my = vp(b.q.mean);
    out.push(line(cx - 4, my - 4, cx + 4, my + 4, '#ffffff', 1.5), line(cx - 4, my + 4, cx + 4, my - 4, '#ffffff', 1.5));
    for (const o of b.q.outliers) out.push(`<circle cx="${r1(cx)}" cy="${r1(vp(o))}" r="3" fill="none" stroke="${b.color}" stroke-width="1.2"/>`);
  }

  // Stock: high-low lines with open/close candles (or close ticks).
  for (const c of candles) {
    const cx = px0 + band * (c.i + 0.5);
    out.push(line(cx, vp(c.high), cx, vp(c.low), '#616161', 1.2));
    if (c.open !== undefined) {
      const up = c.close >= c.open;
      const y0 = vp(Math.max(c.open, c.close));
      const y1 = vp(Math.min(c.open, c.close));
      const bw = Math.max(3, Math.min(band * 0.55, 24));
      out.push(rect(cx - bw / 2, y0, bw, Math.max(1, y1 - y0), up ? '#43A047' : '#E53935', ` stroke="#424242" stroke-width="0.6"`).replace('/>', `>${tip(`${categories[c.i]}\nOpen ${c.open}  High ${c.high}  Low ${c.low}  Close ${c.close}`)}</rect>`));
    } else {
      const y = vp(c.close);
      out.push(line(cx, y, cx + Math.min(band * 0.3, 10), y, '#424242', 2));
    }
  }

  // Lines and areas.
  const base = vp(Math.max(scale.min, Math.min(scale.max, 0)));
  const drawn = isArea ? [...lines].reverse() : lines;
  for (const l of drawn) {
    const sc = l.secondary && sscale ? sscale : scale;
    const segs: { x: number; y: number }[][] = [[]];
    for (const p of l.pts) {
      if (!Number.isFinite(p.v)) {
        if (segs[segs.length - 1].length) segs.push([]);
        continue;
      }
      segs[segs.length - 1].push({ x: cp(p.i), y: vp(p.v, sc) });
    }
    const width = l.s.source.lineWidth ?? 2;
    const dash = dashArray(l.s.source.lineStyle, width);
    for (const seg of segs) {
      if (!seg.length) continue;
      const d = seg.map((p, k) => `${k ? 'L' : 'M'}${r1(p.x)} ${r1(p.y)}`).join(' ');
      if (isArea) {
        // Stacked areas fill down to the previous series (drawn on top of it, so down to the axis works).
        out.push(`<path d="${d} L${r1(seg[seg.length - 1].x)} ${r1(base)} L${r1(seg[0].x)} ${r1(base)} Z" fill="${l.color}" fill-opacity="${stacked ? 0.85 : 0.35}"/>`);
      }
      out.push(`<path d="${d}" fill="none" stroke="${l.color}" stroke-width="${width}" stroke-linejoin="round" stroke-linecap="round"${dash ? ` stroke-dasharray="${dash}"` : ''}/>`);
    }
    const showMarkers = l.s.source.markers || (isLine && l.pts.length === 1);
    for (const p of l.pts) {
      if (!Number.isFinite(p.v)) continue;
      const x = cp(p.i);
      const y = vp(p.v, sc);
      const raw = l.s.values[p.i];
      const title = `${categories[p.i] ?? ''}\n${l.s.name}: ${raw === null || raw === undefined ? '' : formatAxisValue(raw, data.numFmt)}`;
      if (showMarkers) out.push(`<circle cx="${r1(x)}" cy="${r1(y)}" r="3.5" fill="${l.color}" stroke="#ffffff" stroke-width="1">${tip(title)}</circle>`);
      else out.push(`<circle cx="${r1(x)}" cy="${r1(y)}" r="6" fill="transparent">${tip(title)}</circle>`);
      if (chart.valueLabels && raw !== null && raw !== undefined) out.push(text(x, y - 7, formatAxisValue(raw, data.numFmt), { size: 10, anchor: 'middle', fill: TEXT }));
    }
  }

  // Pareto cumulative line.
  if (secondaryLine && sscale) {
    const pts = secondaryLine.values.map((v, i) => ({ x: cp(i), y: vp(v, sscale) }));
    if (pts.length) {
      out.push(`<path d="${pts.map((p, k) => `${k ? 'L' : 'M'}${r1(p.x)} ${r1(p.y)}`).join(' ')}" fill="none" stroke="${secondaryLine.color}" stroke-width="2"/>`);
      pts.forEach((p, i) => out.push(`<circle cx="${r1(p.x)}" cy="${r1(p.y)}" r="3" fill="${secondaryLine!.color}">${tip(`${categories[i]}: ${secondaryLine!.values[i].toFixed(1)}%`)}</circle>`));
    }
  }

  // Trendlines of plotted series.
  for (const s of series) {
    const kind = s.source.trendline;
    if (!kind || kind === 'none' || type === 'stock' || type === 'box-whisker' || type === 'histogram' || type === 'pareto' || type === 'waterfall') continue;
    const color = seriesColor(ctx, s, s.index);
    const sc = secondaryIdx.has(s.index) && !stacked && sscale ? sscale : scale;
    const xs: number[] = [];
    const ys: number[] = [];
    s.values.forEach((v, i) => {
      if (v === null) return;
      xs.push(i + 1);
      ys.push(v);
    });
    if (kind === 'moving-average') {
      const pts: string[] = [];
      for (let i = 1; i < s.values.length; i++) {
        const a = s.values[i - 1];
        const b = s.values[i];
        if (a === null || b === null) continue;
        pts.push(`${pts.length ? 'L' : 'M'}${r1(cp(i))} ${r1(vp((a + b) / 2, sc))}`);
      }
      if (pts.length > 1) out.push(`<path d="${pts.join(' ')}" fill="none" stroke="${color}" stroke-width="1.5" stroke-dasharray="5 3"/>`);
      continue;
    }
    const f = fitTrend(kind, xs, ys);
    if (!f) continue;
    const d: string[] = [];
    const steps = 40;
    for (let k = 0; k <= steps; k++) {
      const x = 1 + ((categories.length - 1) * k) / steps;
      const y = f(x);
      if (!Number.isFinite(y)) continue;
      const px = horizontal ? vp(y, sc) : cp(0) + ((x - 1) / Math.max(1, categories.length - 1)) * (cp(Math.max(0, categories.length - 1)) - cp(0));
      const py = horizontal ? cp(0) + ((x - 1) / Math.max(1, categories.length - 1)) * (cp(Math.max(0, categories.length - 1)) - cp(0)) : vp(y, sc);
      d.push(`${d.length ? 'L' : 'M'}${r1(px)} ${r1(py)}`);
    }
    if (d.length > 1) out.push(`<path d="${d.join(' ')}" fill="none" stroke="${color}" stroke-width="1.5" stroke-dasharray="5 3"/>`);
  }

  // Baseline at zero (or the bottom of the scale).
  const zero = vp(scale.min <= 0 && scale.max >= 0 ? 0 : scale.min);
  if (horizontal) out.push(line(zero, py0, zero, py1, BASELINE));
  else out.push(line(px0, zero, px1, zero, BASELINE));
}

// -------------------------------------------------------------------------------------------------
// Scatter and bubble

function xyChart(ctx: Ctx, plot: Box): void {
  const { chart, data, out } = ctx;
  const bubble = chart.type === 'bubble';
  const pts = data.series.map((s) => {
    const xs = s.xValues && s.xValues.some((v) => v !== null) ? s.xValues : s.values.map((_, i) => i + 1);
    return s.values.map((y, i) => ({ x: xs[i] ?? null, y, size: bubble ? s.sizes?.[i] ?? 1 : 1 }));
  });
  let xlo = Infinity;
  let xhi = -Infinity;
  let ylo = 0;
  let yhi = 0;
  let smax = 0;
  for (const list of pts) {
    for (const p of list) {
      if (p.x === null || p.y === null) continue;
      xlo = Math.min(xlo, p.x);
      xhi = Math.max(xhi, p.x);
      ylo = Math.min(ylo, p.y);
      yhi = Math.max(yhi, p.y);
      smax = Math.max(smax, Math.abs(p.size ?? 0));
    }
  }
  if (!Number.isFinite(xlo)) {
    xlo = 0;
    xhi = 1;
  }
  if (xlo > 0 && xhi > 0 && xlo / xhi < 0.6) xlo = 0;
  const labelSize = 11;
  const yScale = niceScale(ylo, yhi, Math.max(2, Math.floor(plot.h / 50)), chart.yMin, chart.yMax);
  const xScale = niceScale(xlo, xhi, Math.max(2, Math.floor(plot.w / 70)));
  const xTitle = (chart.xAxisTitle ?? '').trim();
  const yTitle = (chart.yAxisTitle ?? '').trim();
  let px0 = plot.x;
  let py1 = plot.y + plot.h;
  if (xTitle) {
    out.push(text(plot.x + plot.w / 2, py1 - 3, clipText(xTitle, plot.w, 12, true), { size: 12, weight: 600, anchor: 'middle', fill: TEXT }));
    py1 -= 20;
  }
  if (yTitle) {
    out.push(text(plot.x + 9, plot.y + plot.h / 2, clipText(yTitle, plot.h, 12, true), { size: 12, weight: 600, anchor: 'middle', fill: TEXT, rotate: -90 }));
    px0 += 20;
  }
  const yLabels = yScale.ticks.map((t) => formatAxisValue(t, data.numFmt));
  px0 += Math.max(...yLabels.map((t) => textWidth(t, labelSize))) + 10;
  py1 -= 20;
  const py0 = plot.y + 6;
  const px1 = plot.x + plot.w - 10;
  const pw = Math.max(10, px1 - px0);
  const ph = Math.max(10, py1 - py0);
  const X = (v: number) => px0 + ((v - xScale.min) / (xScale.max - xScale.min)) * pw;
  const Y = (v: number) => py1 - ((Math.min(yScale.max, Math.max(yScale.min, v)) - yScale.min) / (yScale.max - yScale.min)) * ph;
  yScale.ticks.forEach((t, k) => {
    if (chart.gridlines) out.push(line(px0, Y(t), px1, Y(t), GRID));
    out.push(text(px0 - 6, Y(t) + 4, yLabels[k], { size: labelSize, anchor: 'end' }));
  });
  for (const t of xScale.ticks) {
    if (chart.gridlines) out.push(line(X(t), py0, X(t), py1, GRID));
    out.push(text(X(t), py1 + 15, formatAxisValue(t, 'General'), { size: labelSize, anchor: 'middle' }));
  }
  out.push(line(px0, Y(Math.max(yScale.min, Math.min(yScale.max, 0))), px1, Y(Math.max(yScale.min, Math.min(yScale.max, 0))), BASELINE));
  const maxR = Math.max(4, Math.min(pw, ph) * 0.1);
  data.series.forEach((s, k) => {
    const color = seriesColor(ctx, s, k);
    const list = pts[k].filter((p): p is { x: number; y: number; size: number } => p.x !== null && p.y !== null);
    if (!bubble && s.source.lineWidth) {
      const dash = dashArray(s.source.lineStyle, s.source.lineWidth);
      out.push(`<path d="${list.map((p, i) => `${i ? 'L' : 'M'}${r1(X(p.x))} ${r1(Y(p.y))}`).join(' ')}" fill="none" stroke="${color}" stroke-width="${s.source.lineWidth}"${dash ? ` stroke-dasharray="${dash}"` : ''}/>`);
    }
    for (const p of list) {
      const title = `${s.name}\n(${formatAxisValue(p.x, 'General')}, ${formatAxisValue(p.y, data.numFmt)})${bubble ? ` size ${p.size}` : ''}`;
      const r = bubble ? Math.max(2, Math.sqrt(Math.abs(p.size) / (smax || 1)) * maxR) : 4;
      out.push(`<circle cx="${r1(X(p.x))}" cy="${r1(Y(p.y))}" r="${r1(r)}" fill="${color}"${bubble ? ' fill-opacity="0.75" stroke="#ffffff" stroke-width="1"' : ''}>${tip(title)}</circle>`);
      if (chart.valueLabels) out.push(text(X(p.x), Y(p.y) - r - 4, formatAxisValue(p.y, data.numFmt), { size: 10, anchor: 'middle', fill: TEXT }));
    }
    const kind = s.source.trendline;
    if (kind && kind !== 'none' && kind !== 'moving-average') {
      const f = fitTrend(kind, list.map((p) => p.x), list.map((p) => p.y));
      if (f && list.length) {
        const xa = Math.min(...list.map((p) => p.x));
        const xb = Math.max(...list.map((p) => p.x));
        const d: string[] = [];
        for (let i = 0; i <= 40; i++) {
          const x = xa + ((xb - xa) * i) / 40;
          const y = f(x);
          if (Number.isFinite(y)) d.push(`${d.length ? 'L' : 'M'}${r1(X(x))} ${r1(Y(y))}`);
        }
        if (d.length > 1) out.push(`<path d="${d.join(' ')}" fill="none" stroke="${color}" stroke-width="1.5" stroke-dasharray="5 3"/>`);
      }
    }
  });
}

// -------------------------------------------------------------------------------------------------
// Pie and doughnut

function arcPath(cx: number, cy: number, r: number, inner: number, a0: number, a1: number): string {
  const p = (a: number, rad: number) => `${r1(cx + rad * Math.cos(a))} ${r1(cy + rad * Math.sin(a))}`;
  const large = a1 - a0 > Math.PI ? 1 : 0;
  if (a1 - a0 >= Math.PI * 2 - 1e-6) {
    // Full circle (two halves).
    const mid = a0 + Math.PI;
    if (inner > 0) return `M${p(a0, r)} A${r1(r)} ${r1(r)} 0 1 1 ${p(mid, r)} A${r1(r)} ${r1(r)} 0 1 1 ${p(a0, r)} M${p(a0, inner)} A${r1(inner)} ${r1(inner)} 0 1 0 ${p(mid, inner)} A${r1(inner)} ${r1(inner)} 0 1 0 ${p(a0, inner)} Z`;
    return `M${p(a0, r)} A${r1(r)} ${r1(r)} 0 1 1 ${p(mid, r)} A${r1(r)} ${r1(r)} 0 1 1 ${p(a0, r)} Z`;
  }
  if (inner > 0) return `M${p(a0, r)} A${r1(r)} ${r1(r)} 0 ${large} 1 ${p(a1, r)} L${p(a1, inner)} A${r1(inner)} ${r1(inner)} 0 ${large} 0 ${p(a0, inner)} Z`;
  return `M${r1(cx)} ${r1(cy)} L${p(a0, r)} A${r1(r)} ${r1(r)} 0 ${large} 1 ${p(a1, r)} Z`;
}

function pieChart(ctx: Ctx, plot: Box): void {
  const { chart, data, out, colors } = ctx;
  const doughnut = chart.type === 'doughnut';
  const cx = plot.x + plot.w / 2;
  const cy = plot.y + plot.h / 2;
  const radius = Math.max(8, Math.min(plot.w, plot.h) / 2 - 6);
  const rings = doughnut ? Math.max(1, data.series.length) : 1;
  const hole = doughnut ? radius * 0.5 : 0;
  const ringW = (radius - hole) / rings;
  for (let k = 0; k < rings; k++) {
    const s = data.series[k];
    if (!s) continue;
    const vals = s.values.map((v) => Math.max(0, v ?? 0));
    const total = vals.reduce((a, b) => a + b, 0);
    if (!total) continue;
    const outer = radius - ringW * k;
    const inner = doughnut ? outer - ringW : 0;
    let a = -Math.PI / 2;
    vals.forEach((v, i) => {
      if (!v) return;
      const a1 = a + (v / total) * Math.PI * 2;
      const color = colors[i % colors.length];
      const pct = (v / total) * 100;
      const title = `${data.categories[i] ?? ''}\n${s.name}: ${formatAxisValue(s.values[i] ?? 0, data.numFmt)} (${pct.toFixed(1)}%)`;
      out.push(`<path d="${arcPath(cx, cy, outer, inner, a, a1)}" fill="${color}" stroke="#ffffff" stroke-width="1.5" fill-rule="evenodd">${tip(title)}</path>`);
      if (chart.valueLabels && a1 - a > 0.25) {
        const mid = (a + a1) / 2;
        const lr = doughnut ? (outer + inner) / 2 : outer * 0.65;
        out.push(text(cx + lr * Math.cos(mid), cy + lr * Math.sin(mid) + 4, `${Math.round(pct)}%`, { size: 11, weight: 700, anchor: 'middle', fill: '#ffffff' }));
      }
      a = a1;
    });
  }
}

// -------------------------------------------------------------------------------------------------
// Radar

function radarChart(ctx: Ctx, plot: Box): void {
  const { chart, data, out } = ctx;
  const filled = chart.type === 'filled-radar';
  const n = Math.max(3, data.categories.length);
  const cx = plot.x + plot.w / 2;
  const cy = plot.y + plot.h / 2 + 4;
  const labelPad = Math.min(60, Math.max(...data.categories.map((c) => textWidth(c, 11))) + 6);
  const radius = Math.max(10, Math.min(plot.w / 2 - labelPad, plot.h / 2 - 18));
  let lo = 0;
  let hi = 0;
  for (const s of data.series) for (const v of s.values) if (v !== null) {
    lo = Math.min(lo, v);
    hi = Math.max(hi, v);
  }
  const scale = niceScale(lo, hi, Math.max(2, Math.min(6, Math.floor(radius / 30))), chart.yMin, chart.yMax);
  const ang = (i: number) => -Math.PI / 2 + (i * Math.PI * 2) / n;
  const pt = (i: number, v: number) => {
    const r = ((Math.min(scale.max, Math.max(scale.min, v)) - scale.min) / (scale.max - scale.min)) * radius;
    return { x: cx + r * Math.cos(ang(i)), y: cy + r * Math.sin(ang(i)) };
  };
  if (chart.gridlines) {
    for (const t of scale.ticks) {
      const ring = Array.from({ length: n }, (_, i) => pt(i, t));
      out.push(`<path d="${ring.map((p, k) => `${k ? 'L' : 'M'}${r1(p.x)} ${r1(p.y)}`).join(' ')} Z" fill="none" stroke="${GRID}"/>`);
    }
  }
  for (let i = 0; i < n; i++) {
    const p = pt(i, scale.max);
    out.push(line(cx, cy, p.x, p.y, GRID));
    const lx = cx + (radius + 10) * Math.cos(ang(i));
    const ly = cy + (radius + 10) * Math.sin(ang(i));
    const cos = Math.cos(ang(i));
    out.push(text(lx, ly + 4, clipText(data.categories[i] ?? '', labelPad, 11), { size: 11, anchor: Math.abs(cos) < 0.3 ? 'middle' : cos > 0 ? 'start' : 'end' }));
  }
  for (const t of scale.ticks) {
    const p = pt(0, t);
    out.push(text(p.x + 4, p.y + 4, formatAxisValue(t, data.numFmt), { size: 9 }));
  }
  data.series.forEach((s, k) => {
    const color = seriesColor(ctx, s, k);
    const ptsList = Array.from({ length: n }, (_, i) => pt(i, s.values[i] ?? 0));
    const d = `${ptsList.map((p, i) => `${i ? 'L' : 'M'}${r1(p.x)} ${r1(p.y)}`).join(' ')} Z`;
    const width = s.source.lineWidth ?? 2;
    const dash = dashArray(s.source.lineStyle, width);
    out.push(`<path d="${d}" fill="${filled ? color : 'none'}" fill-opacity="${filled ? 0.35 : 0}" stroke="${color}" stroke-width="${width}" stroke-linejoin="round"${dash ? ` stroke-dasharray="${dash}"` : ''}/>`);
    ptsList.forEach((p, i) => {
      const v = s.values[i];
      const title = `${data.categories[i] ?? ''}\n${s.name}: ${v === null || v === undefined ? '' : formatAxisValue(v, data.numFmt)}`;
      out.push(`<circle cx="${r1(p.x)}" cy="${r1(p.y)}" r="${s.source.markers ? 3.5 : 5}" fill="${s.source.markers ? color : 'transparent'}">${tip(title)}</circle>`);
      if (chart.valueLabels && v !== null && v !== undefined) out.push(text(p.x, p.y - 6, formatAxisValue(v, data.numFmt), { size: 10, anchor: 'middle', fill: TEXT }));
    });
  });
}

// -------------------------------------------------------------------------------------------------
// Funnel and treemap

function funnelChart(ctx: Ctx, plot: Box): void {
  const { data, out } = ctx;
  const s = data.series[0];
  if (!s) return;
  const n = Math.max(1, data.categories.length);
  const labelW = Math.min(plot.w * 0.3, Math.max(...data.categories.map((c) => textWidth(c, 11))) + 10);
  const x0 = plot.x + labelW;
  const w = plot.w - labelW;
  const band = plot.h / n;
  const max = Math.max(1e-9, ...s.values.map((v) => Math.abs(v ?? 0)));
  const color = seriesColor(ctx, s, 0);
  data.categories.forEach((c, i) => {
    const v = s.values[i] ?? 0;
    const bw = (Math.abs(v) / max) * w;
    const y = plot.y + band * i + band * 0.06;
    const bh = band * 0.88;
    out.push(rect(x0 + (w - bw) / 2, y, bw, bh, color).replace('/>', `>${tip(`${c}: ${formatAxisValue(v, data.numFmt)}`)}</rect>`));
    out.push(text(x0 - 6, y + bh / 2 + 4, clipText(c, labelW - 8, 11), { size: 11, anchor: 'end' }));
    const label = formatAxisValue(v, data.numFmt);
    if (textWidth(label, 11) < bw - 6) out.push(text(x0 + w / 2, y + bh / 2 + 4, label, { size: 11, weight: 700, anchor: 'middle', fill: '#ffffff' }));
    else out.push(text(x0 + (w + bw) / 2 + 4, y + bh / 2 + 4, label, { size: 11, fill: TEXT }));
  });
}

function isLight(hex: string): boolean {
  const m = /^#?([0-9a-f]{6})$/i.exec(hex);
  if (!m) return false;
  const n = parseInt(m[1], 16);
  return 0.299 * ((n >> 16) & 255) + 0.587 * ((n >> 8) & 255) + 0.114 * (n & 255) > 170;
}

function treemapChart(ctx: Ctx, plot: Box): void {
  const { data, out, colors } = ctx;
  const s = data.series[0];
  if (!s) return;
  const boxes = squarify(s.values.map((v) => Math.abs(v ?? 0)), plot);
  boxes.forEach((b, i) => {
    if (b.w <= 0 || b.h <= 0) return;
    const color = colors[i % colors.length];
    const name = data.categories[i] ?? '';
    const v = s.values[i] ?? 0;
    out.push(rect(b.x, b.y, b.w, b.h, color, ' stroke="#ffffff" stroke-width="2"').replace('/>', `>${tip(`${name}: ${formatAxisValue(v, data.numFmt)}`)}</rect>`));
    if (b.w > 24 && b.h > 18) {
      const ink = isLight(color) ? '#212121' : '#ffffff';
      out.push(text(b.x + 6, b.y + 16, clipText(name, b.w - 10, 11, true), { size: 11, weight: 700, fill: ink }));
      if (b.h > 34) out.push(text(b.x + 6, b.y + 31, clipText(formatAxisValue(v, data.numFmt), b.w - 10, 11), { size: 11, fill: ink }));
    }
  });
}

// -------------------------------------------------------------------------------------------------

export interface ChartSvgOptions {
  /** Colours used for series without their own (default: the editor's palette). */
  palette?: string[];
  /** Draw a white background (default true). */
  background?: boolean;
}

/** The chart as an SVG document of `w` x `h` pixels. */
export function chartSvg(chart: SheetChart, data: ChartData, w: number, h: number, opts: ChartSvgOptions = {}): string {
  const W = Math.max(40, Math.round(w));
  const H = Math.max(30, Math.round(h));
  const ctx: Ctx = { chart, data, w: W, h: H, out: [], colors: opts.palette?.length ? opts.palette : chart.palette?.length ? chart.palette : CHART_COLORS };
  if (opts.background !== false) ctx.out.push(rect(0, 0, W, H, '#ffffff'));
  const plot = drawFrame(ctx);
  const empty = !data.series.length || data.series.every((s) => !s.values.some((v) => v !== null));
  if (empty) {
    ctx.out.push(text(W / 2, H / 2, 'No data', { size: 12, anchor: 'middle', fill: '#9e9e9e' }));
  } else if (plot.w > 20 && plot.h > 20) {
    switch (chart.type) {
      case 'pie':
      case 'doughnut':
        pieChart(ctx, plot);
        break;
      case 'scatter':
      case 'bubble':
        xyChart(ctx, plot);
        break;
      case 'radar':
      case 'filled-radar':
        radarChart(ctx, plot);
        break;
      case 'funnel':
        funnelChart(ctx, plot);
        break;
      case 'treemap':
        treemapChart(ctx, plot);
        break;
      default:
        cartesian(ctx, plot);
    }
  }
  return `<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 ${W} ${H}" width="${W}" height="${H}" font-family="${FONT}">${ctx.out.join('')}</svg>`;
}
