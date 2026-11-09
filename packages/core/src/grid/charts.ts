/**
 * Live rendering of the charts stored in a workbook (bar/column, line, area, pie/doughnut,
 * scatter). The chart XML is kept in the file (only its cell references follow structural edits);
 * values are read from the referenced ranges, so charts update while the data is edited.
 */
import type { SpreadsheetDocument } from '../core/document';
import { writeChartFormulas } from '../core/xlsx/chartRefs';
import type { Sheet } from '../core/model/sheet';
import { SheetLayout } from '../core/layout';
import { attr, child, childrenNamed, findAll, numAttr, parseXml, textContent, type XmlElement } from '../core/xml/xml';
import { relTypeIs } from '../core/xlsx/package';
import { EMU_PER_PX } from '../core/xlsx/util';
import { ArrayValue, RefValue } from '../core/formula/runtime';
import { formatValue } from '../core/format/display';
import type { Rect } from './viewport';
import type { Scalar } from '../core/values';

export interface ChartInfo {
  id: string;
  rect: Rect;
  title: string;
  type: string;
  element: HTMLElement;
  sourceRanges: Record<string, string[]>;
}

interface Series {
  name: string;
  values: (number | null)[];
  categories: string[];
  xValues?: (number | null)[];
  color?: string;
  refs: { val?: string; cat?: string; name?: string };
}

const PALETTE = ['#4472C4', '#ED7D31', '#A5A5A5', '#FFC000', '#5B9BD5', '#70AD47', '#264478', '#9E480E', '#636363', '#997300'];

function anchorRect(sheet: Sheet, anchor: XmlElement): Rect {
  const layout = SheetLayout.of(sheet);
  const point = (e: XmlElement | undefined) => {
    const c = Number(textContent(child(e, 'xdr:col') ?? { type: 'text', text: '0' })) || 0;
    const r = Number(textContent(child(e, 'xdr:row') ?? { type: 'text', text: '0' })) || 0;
    const dx = Number(textContent(child(e, 'xdr:colOff') ?? { type: 'text', text: '0' })) || 0;
    const dy = Number(textContent(child(e, 'xdr:rowOff') ?? { type: 'text', text: '0' })) || 0;
    return { x: layout.colX(c) + dx / EMU_PER_PX, y: layout.rowY(r) + dy / EMU_PER_PX };
  };
  if (anchor.name === 'xdr:twoCellAnchor') {
    const a = point(child(anchor, 'xdr:from'));
    const b = point(child(anchor, 'xdr:to'));
    return { x: a.x, y: a.y, w: Math.max(20, b.x - a.x), h: Math.max(20, b.y - a.y) };
  }
  const ext = child(anchor, 'xdr:ext');
  const w = (numAttr(ext, 'cx') ?? 0) / EMU_PER_PX;
  const h = (numAttr(ext, 'cy') ?? 0) / EMU_PER_PX;
  if (anchor.name === 'xdr:absoluteAnchor') {
    const pos = child(anchor, 'xdr:pos');
    return { x: (numAttr(pos, 'x') ?? 0) / EMU_PER_PX, y: (numAttr(pos, 'y') ?? 0) / EMU_PER_PX, w, h };
  }
  const a = point(child(anchor, 'xdr:from'));
  return { x: a.x, y: a.y, w, h };
}

function chartXml(doc: SpreadsheetDocument, sheet: Sheet, anchor: XmlElement): XmlElement | null {
  const pkg = doc.wb.source?.pkg;
  if (!pkg || !sheet.sourcePath) return null;
  const chartRef = findAll(anchor, 'c:chart')[0];
  const rid = attr(chartRef, 'r:id');
  if (!rid) return null;
  const drawingEl = child(sheet.sourceXml, 'drawing');
  const drawingRel = drawingEl ? pkg.getRels(sheet.sourcePath).find((r) => r.id === attr(drawingEl, 'r:id')) : undefined;
  if (!drawingRel) return null;
  const drawingPath = pkg.resolveTarget(sheet.sourcePath, drawingRel.target);
  const rel = pkg.getRels(drawingPath).find((r) => r.id === rid && relTypeIs(r.type, 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart'));
  if (!rel) return null;
  const chartPath = pkg.resolveTarget(drawingPath, rel.target);
  let text = pkg.getText(chartPath);
  if (!text) return null;
  // References as updated by edits (renamed sheets, inserted rows...).
  const refs = doc.wb.chartRefs.get(chartPath);
  if (refs) text = writeChartFormulas(text, refs);
  try {
    return parseXml(text).root;
  } catch {
    return null;
  }
}

function richText(e: XmlElement | undefined): string {
  if (!e) return '';
  return findAll(e, 'a:t').map((t) => textContent(t)).join('');
}

function evalRef(doc: SpreadsheetDocument, sheet: Sheet, ref: string | undefined): Scalar[] | null {
  if (!ref) return null;
  try {
    const v = doc.engine.evaluateText(ref, sheet, 0, 0);
    if (v instanceof RefValue) {
      const out: Scalar[] = [];
      for (const area of v.areas) {
        for (let r = area.range.r0; r <= area.range.r1; r++) for (let c = area.range.c0; c <= area.range.c1; c++) out.push(doc.engine.cellValue(area.sheet, r, c));
      }
      return out;
    }
    if (v instanceof ArrayValue) return v.data.flat();
    return [v];
  } catch {
    return null;
  }
}

function cached(e: XmlElement | undefined): Scalar[] {
  const cache = child(e, 'c:numCache') ?? child(e, 'c:strCache');
  if (!cache) return [];
  const count = numAttr(child(cache, 'c:ptCount'), 'val') ?? 0;
  const out: Scalar[] = new Array(count).fill(null);
  for (const pt of childrenNamed(cache, 'c:pt')) {
    const idx = numAttr(pt, 'idx') ?? 0;
    const v = textContent(child(pt, 'c:v') ?? pt);
    out[idx] = cache.name === 'c:numCache' ? Number(v) : v;
  }
  return out;
}

function dataOf(doc: SpreadsheetDocument, sheet: Sheet, e: XmlElement | undefined): { values: Scalar[]; ref?: string } {
  const numRef = child(e, 'c:numRef') ?? child(e, 'c:strRef');
  const ref = numRef ? textContent(child(numRef, 'c:f') ?? { type: 'text', text: '' }) : undefined;
  const live = evalRef(doc, sheet, ref);
  if (live) return { values: live, ref };
  const lit = child(e, 'c:numLit') ?? child(e, 'c:strLit');
  if (lit) return { values: cached({ type: 'element', name: 'x', attrs: {}, children: [{ ...lit, name: lit.name === 'c:numLit' ? 'c:numCache' : 'c:strCache' }] }), ref };
  return { values: cached(numRef), ref };
}

function seriesColor(ser: XmlElement, i: number, theme: Record<string, string>): string {
  const fill = child(child(ser, 'c:spPr'), 'a:solidFill') ?? child(child(child(ser, 'c:spPr'), 'a:ln'), 'a:solidFill');
  const srgb = attr(child(fill, 'a:srgbClr'), 'val');
  if (srgb) return `#${srgb}`;
  const scheme = attr(child(fill, 'a:schemeClr'), 'val');
  if (scheme && theme[scheme]) return `#${theme[scheme]}`;
  return PALETTE[i % PALETTE.length];
}

function readSeries(doc: SpreadsheetDocument, sheet: Sheet, plot: XmlElement): Series[] {
  const theme = doc.wb.theme.colors as unknown as Record<string, string>;
  return childrenNamed(plot, 'c:ser').map((ser, i) => {
    const tx = child(ser, 'c:tx');
    const nameData = dataOf(doc, sheet, tx);
    const name = nameData.values.length ? String(nameData.values[0] ?? '') : textContent(child(tx, 'c:v') ?? { type: 'text', text: '' }) || `Series ${i + 1}`;
    const val = dataOf(doc, sheet, child(ser, 'c:val') ?? child(ser, 'c:yVal'));
    const cat = dataOf(doc, sheet, child(ser, 'c:cat') ?? child(ser, 'c:xVal'));
    const values = val.values.map((v) => (typeof v === 'number' ? v : typeof v === 'string' && v.trim() !== '' && !Number.isNaN(Number(v)) ? Number(v) : null));
    const categories = cat.values.length ? cat.values.map((v) => (v === null ? '' : typeof v === 'number' ? formatValue(v, 'General').text : String(v))) : values.map((_, k) => String(k + 1));
    const xValues = child(ser, 'c:xVal') ? cat.values.map((v) => (typeof v === 'number' ? v : null)) : undefined;
    return { name, values, categories, xValues, color: seriesColor(ser, i, theme), refs: { val: val.ref, cat: cat.ref, name: nameData.ref } };
  });
}

const svgNs = 'http://www.w3.org/2000/svg';

function el(name: string, attrs: Record<string, string | number>, text?: string): SVGElement {
  const e = document.createElementNS(svgNs, name);
  for (const [k, v] of Object.entries(attrs)) e.setAttribute(k, String(v));
  if (text !== undefined) e.textContent = text;
  return e;
}

function niceScale(min: number, max: number, ticks = 5): { min: number; max: number; step: number } {
  if (min === max) {
    max = min + 1;
    min = Math.min(0, min);
  }
  const range = max - min;
  const raw = range / ticks;
  const mag = Math.pow(10, Math.floor(Math.log10(raw)));
  const norm = raw / mag;
  const step = (norm < 1.5 ? 1 : norm < 3 ? 2 : norm < 7 ? 5 : 10) * mag;
  return { min: Math.floor(min / step) * step, max: Math.ceil(max / step) * step, step };
}

function fmt(n: number): string {
  const a = Math.abs(n);
  if (a >= 1e9) return `${(n / 1e9).toFixed(1).replace(/\.0$/, '')}B`;
  if (a >= 1e6) return `${(n / 1e6).toFixed(1).replace(/\.0$/, '')}M`;
  if (a >= 1e4) return `${(n / 1e3).toFixed(0)}K`;
  // Axis steps can carry float noise (0.30000000000000004); whole numbers get no decimal point.
  const v = Number(n.toPrecision(12));
  return formatValue(v, Number.isInteger(v) ? '#,##0' : a < 1 ? '0.0#' : '#,##0.##').text;
}

function renderChart(root: XmlElement, w: number, h: number, doc: SpreadsheetDocument, sheet: Sheet): { svg: SVGSVGElement; title: string; type: string; series: Series[] } {
  const chart = child(root, 'c:chart');
  const plotArea = child(chart, 'c:plotArea');
  const plots = (plotArea ? [...plotArea.children] : []).filter((c): c is XmlElement => c.type === 'element' && /Chart$/.test(c.name));
  const plot = plots[0];
  const type = plot?.name.replace(/^c:/, '') ?? 'unknown';
  const titleEl = child(chart, 'c:title');
  const autoDeleted = attr(child(chart, 'c:autoTitleDeleted'), 'val') === '1';
  const series = plot ? readSeries(doc, sheet, plot) : [];
  let title = richText(child(titleEl, 'c:tx'));
  if (!title && titleEl && !autoDeleted && series.length === 1) title = series[0].name;
  const svg = el('svg', { viewBox: `0 0 ${w} ${h}`, width: '100%', height: '100%', class: 'xchart' }) as SVGSVGElement;
  svg.appendChild(el('rect', { x: 0, y: 0, width: w, height: h, fill: '#ffffff', stroke: '#d0d0d0' }));
  let top = 8;
  if (title) {
    svg.appendChild(el('text', { x: w / 2, y: 22, 'text-anchor': 'middle', 'font-size': 14, fill: '#404040', 'font-family': 'Carlito, Calibri, sans-serif' }, title));
    top = 34;
  }
  const legendOn = !!child(chart, 'c:legend');
  const legendH = legendOn ? 22 : 0;
  const area = { x: 44, y: top + 4, w: w - 60, h: h - top - 30 - legendH };
  if (!plot || !series.length || area.w < 20 || area.h < 20) {
    svg.appendChild(el('text', { x: w / 2, y: h / 2, 'text-anchor': 'middle', 'font-size': 12, fill: '#808080' }, 'Chart'));
    return { svg, title, type, series };
  }
  const isPie = type === 'pieChart' || type === 'doughnutChart' || type === 'pie3DChart';
  if (isPie) {
    const s = series[0];
    const vals = s.values.map((v) => Math.max(0, v ?? 0));
    const total = vals.reduce((a, b) => a + b, 0) || 1;
    const cx = area.x + area.w / 2 - 8;
    const cy = area.y + area.h / 2;
    const rad = Math.max(10, Math.min(area.w, area.h) / 2 - 4);
    let angle = -Math.PI / 2;
    const theme = doc.wb.theme.colors as unknown as Record<string, string>;
    const points = childrenNamed(childrenNamed(plot, 'c:ser')[0], 'c:dPt');
    vals.forEach((v, i) => {
      const a2 = angle + (v / total) * Math.PI * 2;
      const large = a2 - angle > Math.PI ? 1 : 0;
      const x1 = cx + rad * Math.cos(angle);
      const y1 = cy + rad * Math.sin(angle);
      const x2 = cx + rad * Math.cos(a2);
      const y2 = cy + rad * Math.sin(a2);
      const dPt = points.find((p) => numAttr(child(p, 'c:idx'), 'val') === i);
      const color = dPt ? seriesColor(dPt, i, theme) : PALETTE[i % PALETTE.length];
      const path = vals.length === 1 ? `M ${cx - rad} ${cy} a ${rad} ${rad} 0 1 0 ${rad * 2} 0 a ${rad} ${rad} 0 1 0 ${-rad * 2} 0` : `M ${cx} ${cy} L ${x1} ${y1} A ${rad} ${rad} 0 ${large} 1 ${x2} ${y2} Z`;
      svg.appendChild(el('path', { d: path, fill: color, stroke: '#ffffff', 'stroke-width': 1 }));
      angle = a2;
    });
    if (type === 'doughnutChart') svg.appendChild(el('circle', { cx, cy, r: rad * 0.5, fill: '#ffffff' }));
    if (legendOn) drawLegend(svg, s.categories.map((c, i) => ({ name: c, color: PALETTE[i % PALETTE.length] })), w, h);
    return { svg, title, type, series };
  }

  const horizontal = type === 'barChart' && attr(child(plot, 'c:barDir'), 'val') === 'bar';
  const grouping = attr(child(plot, 'c:grouping'), 'val') ?? 'clustered';
  const stacked = grouping === 'stacked' || grouping === 'percentStacked';
  const percent = grouping === 'percentStacked';
  const cats = series[0].categories;
  const n = Math.max(...series.map((s) => s.values.length), cats.length);
  let min = 0;
  let max = 0;
  if (type === 'scatterChart') {
    for (const s of series) for (const v of s.values) if (v !== null) {
      min = Math.min(min, v);
      max = Math.max(max, v);
    }
  } else if (stacked) {
    for (let i = 0; i < n; i++) {
      let pos = 0;
      let neg = 0;
      let tot = 0;
      for (const s of series) {
        const v = s.values[i] ?? 0;
        if (v >= 0) pos += v;
        else neg += v;
        tot += Math.abs(v);
      }
      if (percent) {
        max = Math.max(max, tot ? (pos / tot) * 100 : 0);
        min = Math.min(min, tot ? (neg / tot) * 100 : 0);
      } else {
        max = Math.max(max, pos);
        min = Math.min(min, neg);
      }
    }
  } else {
    for (const s of series) for (const v of s.values) if (v !== null) {
      min = Math.min(min, v);
      max = Math.max(max, v);
    }
  }
  const scale = niceScale(min, max);
  const valueLen = horizontal ? area.w : area.h;
  const toPx = (v: number) => ((v - scale.min) / (scale.max - scale.min)) * valueLen;
  // Grid and value axis labels.
  for (let v = scale.min; v <= scale.max + scale.step / 2; v += scale.step) {
    const p = toPx(v);
    if (horizontal) {
      svg.appendChild(el('line', { x1: area.x + p, y1: area.y, x2: area.x + p, y2: area.y + area.h, stroke: '#e6e6e6' }));
      svg.appendChild(el('text', { x: area.x + p, y: area.y + area.h + 14, 'text-anchor': 'middle', 'font-size': 10, fill: '#595959' }, fmt(v) + (percent ? '%' : '')));
    } else {
      svg.appendChild(el('line', { x1: area.x, y1: area.y + area.h - p, x2: area.x + area.w, y2: area.y + area.h - p, stroke: '#e6e6e6' }));
      svg.appendChild(el('text', { x: area.x - 4, y: area.y + area.h - p + 3, 'text-anchor': 'end', 'font-size': 10, fill: '#595959' }, fmt(v) + (percent ? '%' : '')));
    }
  }
  const band = (horizontal ? area.h : area.w) / Math.max(1, n);
  // Category labels.
  if (type !== 'scatterChart') {
    const step = Math.max(1, Math.ceil(n / Math.max(1, Math.floor((horizontal ? area.h : area.w) / 50))));
    for (let i = 0; i < n; i += step) {
      const label = (cats[i] ?? '').slice(0, 14);
      if (horizontal) svg.appendChild(el('text', { x: area.x - 4, y: area.y + band * (i + 0.5) + 3, 'text-anchor': 'end', 'font-size': 10, fill: '#595959' }, label));
      else svg.appendChild(el('text', { x: area.x + band * (i + 0.5), y: area.y + area.h + 14, 'text-anchor': 'middle', 'font-size': 10, fill: '#595959' }, label));
    }
  }
  const zero = toPx(Math.max(scale.min, Math.min(scale.max, 0)));
  if (type === 'barChart' || type === 'bar3DChart') {
    const gap = band * 0.25;
    const inner = band - gap * 2;
    const per = stacked ? inner : inner / series.length;
    for (let i = 0; i < n; i++) {
      let pos = 0;
      let neg = 0;
      const tot = series.reduce((a, s) => a + Math.abs(s.values[i] ?? 0), 0) || 1;
      series.forEach((s, k) => {
        let v = s.values[i] ?? 0;
        if (percent) v = (v / tot) * 100;
        let start: number;
        let end: number;
        if (stacked) {
          start = v >= 0 ? toPx(pos) : toPx(neg);
          if (v >= 0) pos += v;
          else neg += v;
          end = v >= 0 ? toPx(pos) : toPx(neg);
        } else {
          start = zero;
          end = toPx(v);
        }
        const lo = Math.min(start, end);
        const len = Math.abs(end - start);
        const off = gap + (stacked ? 0 : per * k);
        if (horizontal) svg.appendChild(el('rect', { x: area.x + lo, y: area.y + band * i + off, width: len, height: Math.max(1, per - 1), fill: s.color ?? PALETTE[k % PALETTE.length] }));
        else svg.appendChild(el('rect', { x: area.x + band * i + off, y: area.y + area.h - lo - len, width: Math.max(1, per - 1), height: len, fill: s.color ?? PALETTE[k % PALETTE.length] }));
      });
    }
  } else if (type === 'scatterChart') {
    let xmin = Infinity;
    let xmax = -Infinity;
    for (const s of series) for (const x of s.xValues ?? s.values.map((_, i) => i + 1)) if (x !== null) {
      xmin = Math.min(xmin, x);
      xmax = Math.max(xmax, x);
    }
    const xs = niceScale(Number.isFinite(xmin) ? xmin : 0, Number.isFinite(xmax) ? xmax : 1);
    series.forEach((s, k) => {
      const xsVals = s.xValues ?? s.values.map((_, i) => i + 1);
      s.values.forEach((v, i) => {
        const x = xsVals[i];
        if (v === null || x === null) return;
        const px = area.x + ((x - xs.min) / (xs.max - xs.min)) * area.w;
        svg.appendChild(el('circle', { cx: px, cy: area.y + area.h - toPx(v), r: 3, fill: s.color ?? PALETTE[k % PALETTE.length] }));
      });
    });
  } else {
    // Line and area charts.
    const isArea = type === 'areaChart';
    const cum = new Array(n).fill(0);
    series.forEach((s, k) => {
      const pts: [number, number][] = [];
      for (let i = 0; i < n; i++) {
        const raw = s.values[i];
        if (raw === null || raw === undefined) continue;
        let v: number = raw;
        if (stacked) {
          cum[i] += v;
          v = cum[i];
        }
        pts.push([area.x + band * (i + 0.5), area.y + area.h - toPx(v)]);
      }
      if (!pts.length) return;
      const color = s.color ?? PALETTE[k % PALETTE.length];
      const d = pts.map(([x, y], i) => `${i ? 'L' : 'M'} ${x.toFixed(1)} ${y.toFixed(1)}`).join(' ');
      if (isArea) svg.appendChild(el('path', { d: `${d} L ${pts[pts.length - 1][0]} ${area.y + area.h - zero} L ${pts[0][0]} ${area.y + area.h - zero} Z`, fill: color, 'fill-opacity': 0.6 }));
      svg.appendChild(el('path', { d, fill: 'none', stroke: color, 'stroke-width': 2 }));
    });
  }
  svg.appendChild(el('line', horizontal ? { x1: area.x + zero, y1: area.y, x2: area.x + zero, y2: area.y + area.h, stroke: '#bfbfbf' } : { x1: area.x, y1: area.y + area.h - zero, x2: area.x + area.w, y2: area.y + area.h - zero, stroke: '#bfbfbf' }));
  if (legendOn) drawLegend(svg, series.map((s, i) => ({ name: s.name, color: s.color ?? PALETTE[i % PALETTE.length] })), w, h);
  return { svg, title, type, series };
}

function drawLegend(svg: SVGSVGElement, items: { name: string; color: string }[], w: number, h: number): void {
  const widths = items.map((i) => Math.min(120, i.name.length * 6 + 18));
  const total = widths.reduce((a, b) => a + b, 0);
  let x = Math.max(4, (w - total) / 2);
  const y = h - 10;
  items.forEach((item, i) => {
    svg.appendChild(el('rect', { x, y: y - 8, width: 8, height: 8, fill: item.color }));
    svg.appendChild(el('text', { x: x + 11, y, 'font-size': 10, fill: '#595959' }, item.name.slice(0, 18)));
    x += widths[i];
  });
}

export function chartsOf(doc: SpreadsheetDocument, sheet: Sheet): ChartInfo[] {
  const out: ChartInfo[] = [];
  sheet.otherAnchors.forEach((anchor, index) => {
    const frame = findAll(anchor, 'xdr:graphicFrame')[0];
    const rect = anchorRect(sheet, anchor);
    const name = attr(findAll(anchor, 'xdr:cNvPr')[0], 'name') ?? `Object ${index + 1}`;
    const element = document.createElement('div');
    element.className = 'xchart-box';
    if (!frame) {
      element.classList.add('xchart-shape');
      element.title = name;
      out.push({ id: `${sheet.uid}-${index}`, rect, title: name, type: 'shape', element, sourceRanges: {} });
      return;
    }
    const xml = chartXml(doc, sheet, anchor);
    if (!xml) {
      element.textContent = name;
      out.push({ id: `${sheet.uid}-${index}`, rect, title: name, type: 'object', element, sourceRanges: {} });
      return;
    }
    const { svg, title, type, series } = renderChart(xml, Math.max(60, rect.w), Math.max(40, rect.h), doc, sheet);
    element.appendChild(svg);
    element.setAttribute('role', 'img');
    element.setAttribute('aria-label', title || name);
    const sourceRanges: Record<string, string[]> = {};
    for (const s of series) {
      if (s.refs.val) (sourceRanges.DataRange ??= []).push(s.refs.val);
      if (s.refs.cat) (sourceRanges.CategoryRange ??= []).push(s.refs.cat);
      if (s.refs.name) (sourceRanges.NameRange ??= []).push(s.refs.name);
    }
    out.push({ id: `${sheet.uid}-${index}`, rect, title: title || name, type, element, sourceRanges });
  });
  return out;
}

export function chartPlaceholders(doc: SpreadsheetDocument, sheet: Sheet): { rect: Rect; element: HTMLElement }[] {
  if (!sheet.otherAnchors.length || typeof document === 'undefined') return [];
  return chartsOf(doc, sheet);
}
