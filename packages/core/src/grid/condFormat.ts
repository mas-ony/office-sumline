/**
 * Conditional formatting for display: evaluates the common rule types (cell value, expression,
 * text, blanks/errors, duplicates, top/bottom, above/below average, color scales, data bars) and
 * returns the formatting to draw. Rules are kept unchanged in the file.
 */
import type { SpreadsheetDocument } from '../core/document';
import type { Sheet } from '../core/model/sheet';
import type { FontStyle } from '../core/model/style';
import { resolveColor, type ColorSpec } from '../core/model/color';
import { attr, child, childrenNamed, elements, numAttr, textContent, type XmlElement } from '../core/xml/xml';
import { readColor } from '../core/xlsx/styles';
import { translateFormula } from '../core/formula/adjust';
import { ArrayValue, RefValue, compareScalars, toBoolean } from '../core/formula/runtime';
import { CellError, type Scalar } from '../core/values';
import { rangeContains, type Range } from '../core/refs';
import { formatValue } from '../core/format/display';

export interface CfResult {
  fill?: string;
  color?: string;
  font?: Partial<FontStyle>;
  bar?: { ratio: number; color: string };
}

interface Dxf {
  fill?: ColorSpec;
  color?: ColorSpec;
  font?: Partial<FontStyle>;
}

interface Rule {
  type: string;
  priority: number;
  stop: boolean;
  dxf?: Dxf;
  raw: XmlElement;
  ranges: Range[];
}

let version = 0;

/** Invalidates cached rule results (call after data changes). */
export function bumpConditionalFormats(): void {
  version++;
}

interface SheetCache {
  version: number;
  rules: Rule[];
  results: Map<number, CfResult | null>;
  stats: Map<Rule, { values: number[]; min: number; max: number; avg: number; std: number }>;
}

const caches = new WeakMap<Sheet, SheetCache>();
const dxfCache = new WeakMap<object, Dxf[]>();

function readDxfs(doc: SpreadsheetDocument): Dxf[] {
  const list = doc.wb.dxfs;
  if (!list.length) return [];
  const hit = dxfCache.get(list);
  if (hit) return hit;
  const out = list.map((d): Dxf => {
    const font = child(d, 'font');
    const fill = child(child(d, 'fill'), 'patternFill');
    const dxf: Dxf = {};
    if (font) {
      const f: Partial<FontStyle> = {};
      if (child(font, 'b')) f.bold = attr(child(font, 'b'), 'val') !== '0';
      if (child(font, 'i')) f.italic = attr(child(font, 'i'), 'val') !== '0';
      if (child(font, 'strike')) f.strike = attr(child(font, 'strike'), 'val') !== '0';
      if (child(font, 'u')) f.underline = 'single';
      dxf.font = f;
      dxf.color = readColor(child(font, 'color'));
    }
    if (fill) dxf.fill = readColor(child(fill, 'bgColor')) ?? readColor(child(fill, 'fgColor'));
    return dxf;
  });
  dxfCache.set(list, out);
  return out;
}

function rulesOf(doc: SpreadsheetDocument, sheet: Sheet): Rule[] {
  const dxfs = readDxfs(doc);
  const rules: Rule[] = [];
  for (const cf of sheet.condFormats) {
    for (const r of childrenNamed(cf.raw, 'cfRule')) {
      const dxfId = numAttr(r, 'dxfId');
      rules.push({
        type: attr(r, 'type') ?? '',
        priority: numAttr(r, 'priority') ?? 1000,
        stop: attr(r, 'stopIfTrue') === '1',
        dxf: dxfId !== undefined ? dxfs[dxfId] : undefined,
        raw: r,
        ranges: cf.sqref,
      });
    }
  }
  return rules.sort((a, b) => a.priority - b.priority);
}

function valueAt(sheet: Sheet, r: number, c: number): Scalar {
  return sheet.getCell(r, c)?.v ?? null;
}

function evalFormula(doc: SpreadsheetDocument, sheet: Sheet, formula: string, rule: Rule, r: number, c: number): Scalar {
  const origin = rule.ranges[0];
  const f = translateFormula(formula, r - origin.r0, c - origin.c0);
  const v = doc.engine.evaluateText(f, sheet, r, c);
  if (v instanceof RefValue) return doc.engine.cellValue(v.single.sheet, v.single.range.r0, v.single.range.c0);
  if (v instanceof ArrayValue) return v.get(0, 0);
  return v;
}

function numbersInRanges(sheet: Sheet, ranges: Range[]): number[] {
  const out: number[] = [];
  for (const range of ranges) sheet.forEachCellIn(range, (_r, _c, cell) => {
    if (typeof cell.v === 'number') out.push(cell.v);
  });
  return out;
}

function stats(cache: SheetCache, sheet: Sheet, rule: Rule) {
  let s = cache.stats.get(rule);
  if (!s) {
    const values = numbersInRanges(sheet, rule.ranges).sort((a, b) => a - b);
    const n = values.length;
    const avg = n ? values.reduce((a, b) => a + b, 0) / n : 0;
    const std = n > 1 ? Math.sqrt(values.reduce((a, b) => a + (b - avg) ** 2, 0) / (n - 1)) : 0;
    s = { values, min: values[0] ?? 0, max: values[n - 1] ?? 0, avg, std };
    cache.stats.set(rule, s);
  }
  return s;
}

function cfvoValue(doc: SpreadsheetDocument, sheet: Sheet, cfvo: XmlElement | undefined, st: { values: number[]; min: number; max: number }, rule: Rule): number {
  const type = attr(cfvo, 'type');
  const val = attr(cfvo, 'val');
  switch (type) {
    case 'min':
      return st.min;
    case 'max':
      return st.max;
    case 'percent':
      return st.min + ((st.max - st.min) * Number(val ?? 0)) / 100;
    case 'percentile': {
      const p = Number(val ?? 50) / 100;
      const h = (st.values.length - 1) * p;
      const lo = Math.floor(h);
      return st.values.length ? st.values[lo] + (h - lo) * ((st.values[lo + 1] ?? st.values[lo]) - st.values[lo]) : 0;
    }
    case 'formula': {
      const v = evalFormula(doc, sheet, val ?? '0', rule, rule.ranges[0].r0, rule.ranges[0].c0);
      return typeof v === 'number' ? v : 0;
    }
    default:
      return Number(val ?? 0);
  }
}

function mix(a: string, b: string, t: number): string {
  const p = (h: string) => [parseInt(h.slice(0, 2), 16), parseInt(h.slice(2, 4), 16), parseInt(h.slice(4, 6), 16)];
  const [r1, g1, b1] = p(a);
  const [r2, g2, b2] = p(b);
  const m = (x: number, y: number) => Math.round(x + (y - x) * t).toString(16).padStart(2, '0');
  return `#${m(r1, r2)}${m(g1, g2)}${m(b1, b2)}`;
}

function evaluate(doc: SpreadsheetDocument, sheet: Sheet, cache: SheetCache, rule: Rule, r: number, c: number): CfResult | null | 'skip' {
  const v = valueAt(sheet, r, c);
  const theme = doc.wb.theme;
  const apply = (): CfResult => {
    const out: CfResult = {};
    if (rule.dxf?.fill) out.fill = `#${resolveColor(rule.dxf.fill, theme, 'FFFFFF')}`;
    if (rule.dxf?.color) out.color = `#${resolveColor(rule.dxf.color, theme, '000000')}`;
    if (rule.dxf?.font && Object.keys(rule.dxf.font).length) out.font = rule.dxf.font;
    return out;
  };
  const formulas = childrenNamed(rule.raw, 'formula').map((f) => textContent(f));
  switch (rule.type) {
    case 'cellIs': {
      if (v === null) return null;
      const a = evalFormula(doc, sheet, formulas[0] ?? '0', rule, r, c);
      const b = formulas[1] !== undefined ? evalFormula(doc, sheet, formulas[1], rule, r, c) : null;
      const cmp = compareScalars(v, a);
      const op = attr(rule.raw, 'operator') ?? 'equal';
      const ok =
        op === 'lessThan' ? cmp < 0 : op === 'lessThanOrEqual' ? cmp <= 0 : op === 'greaterThan' ? cmp > 0 : op === 'greaterThanOrEqual' ? cmp >= 0
          : op === 'notEqual' ? cmp !== 0 : op === 'between' ? cmp >= 0 && compareScalars(v, b) <= 0 : op === 'notBetween' ? cmp < 0 || compareScalars(v, b) > 0 : cmp === 0;
      return ok ? apply() : null;
    }
    case 'expression': {
      const res = evalFormula(doc, sheet, formulas[0] ?? 'FALSE', rule, r, c);
      const b = res instanceof CellError ? false : toBoolean(res);
      return b === true ? apply() : null;
    }
    case 'containsText':
    case 'notContainsText':
    case 'beginsWith':
    case 'endsWith': {
      const text = (attr(rule.raw, 'text') ?? '').toLowerCase();
      const s = v === null ? '' : formatValue(v, 'General').text.toLowerCase();
      const ok = rule.type === 'containsText' ? s.includes(text) : rule.type === 'notContainsText' ? !s.includes(text) : rule.type === 'beginsWith' ? s.startsWith(text) : s.endsWith(text);
      return ok ? apply() : null;
    }
    case 'containsBlanks':
      return v === null || (typeof v === 'string' && !v.trim()) ? apply() : null;
    case 'notContainsBlanks':
      return v !== null && !(typeof v === 'string' && !v.trim()) ? apply() : null;
    case 'containsErrors':
      return v instanceof CellError ? apply() : null;
    case 'notContainsErrors':
      return !(v instanceof CellError) ? apply() : null;
    case 'duplicateValues':
    case 'uniqueValues': {
      if (v === null) return null;
      let count = 0;
      for (const range of rule.ranges) sheet.forEachCellIn(range, (_r, _c, cell) => {
        if (cell.v !== null && compareScalars(cell.v, v) === 0) count++;
      });
      return (rule.type === 'duplicateValues' ? count > 1 : count === 1) ? apply() : null;
    }
    case 'top10': {
      if (typeof v !== 'number') return null;
      const st = stats(cache, sheet, rule);
      const rank = numAttr(rule.raw, 'rank') ?? 10;
      const n = attr(rule.raw, 'percent') === '1' ? Math.max(1, Math.floor((st.values.length * rank) / 100)) : rank;
      const bottom = attr(rule.raw, 'bottom') === '1';
      const threshold = bottom ? st.values[Math.min(n, st.values.length) - 1] : st.values[Math.max(0, st.values.length - n)];
      return (bottom ? v <= threshold : v >= threshold) ? apply() : null;
    }
    case 'aboveAverage': {
      if (typeof v !== 'number') return null;
      const st = stats(cache, sheet, rule);
      const above = attr(rule.raw, 'aboveAverage') !== '0';
      const eq = attr(rule.raw, 'equalAverage') === '1';
      const sd = numAttr(rule.raw, 'stdDev') ?? 0;
      const limit = st.avg + (above ? 1 : -1) * sd * st.std;
      const ok = above ? (eq ? v >= limit : v > limit) : eq ? v <= limit : v < limit;
      return ok ? apply() : null;
    }
    case 'colorScale': {
      if (typeof v !== 'number') return null;
      const scale = child(rule.raw, 'colorScale');
      const cfvos = childrenNamed(scale, 'cfvo');
      const colors = childrenNamed(scale, 'color').map((x) => resolveColor(readColor(x), theme, 'FFFFFF'));
      const st = stats(cache, sheet, rule);
      const points = cfvos.map((x) => cfvoValue(doc, sheet, x, st, rule));
      if (points.length < 2 || colors.length < 2) return null;
      let fill: string;
      if (points.length === 2) {
        const t = points[1] === points[0] ? 0 : (v - points[0]) / (points[1] - points[0]);
        fill = mix(colors[0], colors[1], Math.max(0, Math.min(1, t)));
      } else if (v <= points[1]) {
        const t = points[1] === points[0] ? 0 : (v - points[0]) / (points[1] - points[0]);
        fill = mix(colors[0], colors[1], Math.max(0, Math.min(1, t)));
      } else {
        const t = points[2] === points[1] ? 1 : (v - points[1]) / (points[2] - points[1]);
        fill = mix(colors[1], colors[2], Math.max(0, Math.min(1, t)));
      }
      return { fill };
    }
    case 'dataBar': {
      if (typeof v !== 'number') return null;
      const bar = child(rule.raw, 'dataBar');
      const cfvos = childrenNamed(bar, 'cfvo');
      const st = stats(cache, sheet, rule);
      const lo = cfvoValue(doc, sheet, cfvos[0], st, rule);
      const hi = cfvoValue(doc, sheet, cfvos[1], st, rule);
      const col = resolveColor(readColor(child(bar, 'color')), theme, '638EC6');
      const ratio = hi === lo ? 1 : (v - Math.min(0, lo)) / (hi - Math.min(0, lo));
      return { bar: { ratio, color: `#${col}` } };
    }
    default:
      return 'skip';
  }
}

/** Returns a function giving the conditional format of a cell, or null when the sheet has none. */
export function condFormatFor(doc: SpreadsheetDocument, sheet: Sheet): ((r: number, c: number) => CfResult | null) | null {
  if (!sheet.condFormats.length) return null;
  let cache = caches.get(sheet);
  if (!cache || cache.version !== version) {
    cache = { version, rules: rulesOf(doc, sheet), results: new Map(), stats: new Map() };
    caches.set(sheet, cache);
  }
  const c0 = cache;
  return (r, c) => {
    const key = r * 16384 + c;
    const hit = c0.results.get(key);
    if (hit !== undefined) return hit;
    let out = null as CfResult | null;
    for (const rule of c0.rules) {
      if (!rule.ranges.some((rg) => rangeContains(rg, r, c))) continue;
      let res: CfResult | null | 'skip';
      try {
        res = evaluate(doc, sheet, c0, rule, r, c);
      } catch {
        res = null;
      }
      if (res === 'skip' || !res) continue;
      const prev: CfResult = out ?? {};
      out = { ...res, ...prev, font: { ...(res.font ?? {}), ...(prev.font ?? {}) } };
      if (rule.stop) break;
    }
    c0.results.set(key, out);
    return out;
  };
}

void elements;
