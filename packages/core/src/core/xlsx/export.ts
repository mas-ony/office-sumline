/**
 * Writes the editor model to an XLSX package. When the workbook came from a file, the original
 * package is patched: unknown parts, relationships and XML elements are kept, and only what the
 * editor owns (cells, styles, shared strings, merges, views, hyperlinks, images, comments, the
 * sheet list and defined names) is regenerated.
 */
import { strToU8 } from 'fflate';
import { writeChartFormulas } from './chartRefs';
import { cellName, colName, quoteSheetName, rangeName, sqrefString, type Range } from '../refs';
import { CellError } from '../values';
import type { AutoFilter, Cell, CommentThread, DataValidation, RichTextRun, Sheet, SheetImage } from '../model/sheet';
import { widthPxToChars } from '../model/sheet';
import type { Workbook } from '../model/workbook';
import type { FontStyle } from '../model/style';
import {
  attr, child, cloneXml, el, elements, escapeAttr, escapeText, namespacesFor, parseXml, serializeNode, serializeXml, stripInvalidXmlChars,
  txt, XML_DECL, type XmlElement,
} from '../xml/xml';
import { MAIN_NS, NS, REL_NS } from '../xml/namespaces';
import { CT, IMAGE_CONTENT_TYPES, OoxmlPackage, REL, extensionForMime, relTypeIs, type Relationship } from './package';
import { colorAttrs, writeStyles } from './styles';
import { DEFAULT_THEME_XML } from './theme';
import { formulaToFile } from './formulaText';
import { needsArrayFormula } from './arrayFormulas';
import { pxToEmu } from './util';
import { CHARTEX_TYPES, type SheetChart, type SheetShape } from '../model/drawing';
import { chartAnchor, chartReference, CT_CHART, CT_CHARTEX, REL_CHART, REL_CHARTEX, repositionAnchor, shapeAnchor, writeChartexPart, writeChartPart } from './drawingObjects';

const THREADED_NS = NS.tc;

export function newGuid(): string {
  const bytes = new Uint8Array(16);
  const c = (globalThis as { crypto?: Crypto }).crypto;
  if (c?.getRandomValues) c.getRandomValues(bytes);
  else for (let i = 0; i < 16; i++) bytes[i] = Math.floor(Math.random() * 256);
  bytes[6] = (bytes[6] & 0x0f) | 0x40;
  bytes[8] = (bytes[8] & 0x3f) | 0x80;
  const h = [...bytes].map((b) => b.toString(16).padStart(2, '0')).join('').toUpperCase();
  return `{${h.slice(0, 8)}-${h.slice(8, 12)}-${h.slice(12, 16)}-${h.slice(16, 20)}-${h.slice(20)}}`;
}

/**
 * Numbers as stored in files: the shortest text that reads back as the same double (like Excel,
 * which keeps full precision in the file and rounds to 15 digits only for display).
 */
export function numberXml(n: number): string {
  return Number.isFinite(n) ? String(n).replace('e', 'E') : '0';
}

/** Threaded comment time stamps the way Excel writes them: UTC without a zone, centiseconds. */
export function excelDateTime(iso: string): string {
  const m = /^(\d{4}-\d{2}-\d{2}T\d{2}:\d{2}:\d{2})(\.\d+)?(Z|[+-]\d{2}:\d{2})?$/.exec(iso);
  if (!m || (m[3] && m[3] !== 'Z')) return iso;
  return `${m[1]}${(m[2] ?? '.00').padEnd(3, '0').slice(0, 3)}`;
}

/** Escapes characters Excel cannot store directly as _xHHHH_. */
function escapeExcelText(s: string): string {
  // eslint-disable-next-line no-control-regex
  return s.replace(/_x[0-9A-Fa-f]{4}_/g, (m) => `_x005F${m}`).replace(/[\u0000-\u0008\u000B\u000C\u000E-\u001F]/g, (ch) => `_x${ch.charCodeAt(0).toString(16).toUpperCase().padStart(4, '0')}_`);
}

function textElementXml(text: string): string {
  return `<t xml:space="preserve">${escapeText(escapeExcelText(text))}</t>`;
}

function runPropsXml(f: Partial<FontStyle> | undefined): string {
  if (!f) return '';
  const parts: string[] = [];
  if (f.bold) parts.push('<b/>');
  if (f.italic) parts.push('<i/>');
  if (f.strike) parts.push('<strike/>');
  if (f.underline) parts.push(f.underline === 'single' ? '<u/>' : `<u val="${f.underline}"/>`);
  if (f.vertAlign) parts.push(`<vertAlign val="${f.vertAlign}"/>`);
  if (f.size) parts.push(`<sz val="${f.size}"/>`);
  if (f.color) parts.push(serializeNode(el('color', colorAttrs(f.color))));
  if (f.name) parts.push(`<rFont val="${escapeAttr(f.name)}"/>`);
  if (f.family !== undefined) parts.push(`<family val="${f.family}"/>`);
  if (f.scheme) parts.push(`<scheme val="${f.scheme}"/>`);
  return parts.length ? `<rPr>${parts.join('')}</rPr>` : '';
}

function richTextXml(runs: RichTextRun[]): string {
  return runs.map((r) => `<r>${runPropsXml(r.font)}${textElementXml(r.text)}</r>`).join('');
}

class SharedStrings {
  private index = new Map<string, number>();
  readonly items: string[] = [];
  count = 0;

  add(text: string, runs?: RichTextRun[]): number {
    this.count++;
    const key = runs ? `\u0001${JSON.stringify(runs)}` : text;
    let i = this.index.get(key);
    if (i === undefined) {
      i = this.items.length;
      this.items.push(runs ? richTextXml(runs) : textElementXml(text));
      this.index.set(key, i);
    }
    return i;
  }

  xml(): string {
    return `${XML_DECL}<sst xmlns="${MAIN_NS}" count="${this.count}" uniqueCount="${this.items.length}">${this.items.map((x) => `<si>${x}</si>`).join('')}</sst>`;
  }
}

interface WriteContext {
  pkg: OoxmlPackage;
  wb: Workbook;
  wbPath: string;
  sst: SharedStrings;
  dynamicCm: number | null;
  needsDynamicMetadata: boolean;
  usesExtension: boolean;
  persons: Map<string, { id: string; displayName: string; userId?: string; providerId?: string }>;
}

function basePackage(): OoxmlPackage {
  const pkg = new OoxmlPackage();
  pkg.contentTypes.defaults.set('rels', CT.rels);
  pkg.contentTypes.defaults.set('xml', CT.xml);
  pkg.setRels('', [
    { id: 'rId1', type: REL.officeDocument, target: 'xl/workbook.xml' },
    { id: 'rId2', type: REL.coreProps, target: 'docProps/core.xml' },
    { id: 'rId3', type: REL.extendedProps, target: 'docProps/app.xml' },
  ]);
  pkg.setText('xl/theme/theme1.xml', DEFAULT_THEME_XML, CT.theme);
  pkg.setRels('xl/workbook.xml', [{ id: 'rIdTheme', type: REL.theme, target: 'theme/theme1.xml' }]);
  return pkg;
}

/** Collects r:id-like attribute values used in kept XML (so their relationships are kept). */
function collectRelIds(node: XmlElement, into: Set<string>): void {
  for (const [k, v] of Object.entries(node.attrs)) if (k.startsWith('r:')) into.add(v);
  for (const c of node.children) if (c.type === 'element') collectRelIds(c, into);
}

// -----------------------------------------------------------------------------------------------
// Worksheets

const WORKSHEET_ORDER = [
  'sheetPr', 'dimension', 'sheetViews', 'sheetFormatPr', 'cols', 'sheetData', 'sheetCalcPr', 'sheetProtection',
  'protectedRanges', 'scenarios', 'autoFilter', 'sortState', 'dataConsolidate', 'customSheetViews', 'mergeCells',
  'phoneticPr', 'conditionalFormatting', 'dataValidations', 'hyperlinks', 'printOptions', 'pageMargins', 'pageSetup',
  'headerFooter', 'rowBreaks', 'colBreaks', 'customProperties', 'cellWatches', 'ignoredErrors', 'smartTags', 'drawing',
  'legacyDrawing', 'legacyDrawingHF', 'drawingHF', 'picture', 'oleObjects', 'controls', 'webPublishItems', 'tableParts',
  'extLst',
];
/** Elements the writer regenerates from the model. */
const OWNED = new Set(['sheetPr', 'dimension', 'sheetViews', 'sheetFormatPr', 'cols', 'sheetData', 'autoFilter', 'mergeCells', 'conditionalFormatting', 'dataValidations', 'hyperlinks', 'drawing', 'legacyDrawing']);

function cellXml(r: number, c: number, cell: Cell, ctx: WriteContext): string {
  let attrs = ` r="${colName(c)}${r + 1}"`;
  if (cell.s) attrs += ` s="${cell.s}"`;
  const v = cell.v;
  let type = '';
  let inner = '';
  if (cell.f !== undefined) {
    const text = escapeText(stripInvalidXmlChars(formulaToFile(cell.f)));
    if (cell.arr && !cell.dyn) inner += `<f t="array" ref="${cell.arr}">${text}</f>`;
    else if (cell.dyn && (cell.arr || needsArrayFormula(cell.f))) {
      ctx.needsDynamicMetadata = true;
      attrs += ` cm="${ctx.dynamicCm ?? 1}"`;
      inner += `<f t="array" ref="${cell.arr ?? `${colName(c)}${r + 1}`}" aca="false">${text}</f>`;
    } else inner += `<f>${text}</f>`;
    if (typeof v === 'string') {
      type = ' t="str"';
      inner += `<v>${escapeText(escapeExcelText(v))}</v>`;
    } else if (typeof v === 'number') inner += `<v>${numberXml(v)}</v>`;
    else if (typeof v === 'boolean') {
      type = ' t="b"';
      inner += `<v>${v ? 1 : 0}</v>`;
    } else if (v instanceof CellError) {
      type = ' t="e"';
      inner += `<v>${escapeText(v.code)}</v>`;
    }
  } else if (typeof v === 'string') {
    type = ' t="s"';
    inner = `<v>${ctx.sst.add(v, cell.rt && cell.rt.length ? cell.rt : undefined)}</v>`;
  } else if (typeof v === 'number') inner = `<v>${numberXml(v)}</v>`;
  else if (typeof v === 'boolean') {
    type = ' t="b"';
    inner = `<v>${v ? 1 : 0}</v>`;
  } else if (v instanceof CellError) {
    type = ' t="e"';
    inner = `<v>${escapeText(v.code)}</v>`;
  }
  return inner ? `<c${attrs}${type}>${inner}</c>` : `<c${attrs}/>`;
}

function sheetDataXml(sheet: Sheet, ctx: WriteContext): string {
  const rows = new Set<number>([...sheet.cells.keys(), ...sheet.rows.keys()]);
  const sorted = [...rows].sort((a, b) => a - b);
  const out: string[] = ['<sheetData>'];
  for (const r of sorted) {
    const info = sheet.rows.get(r);
    let attrs = ` r="${r + 1}"`;
    if (info?.s !== undefined) attrs += ` s="${info.s}" customFormat="1"`;
    if (info?.h !== undefined) attrs += ` ht="${info.h}"${info.customHeight ? ' customHeight="1"' : ''}`;
    if (info?.hidden) attrs += ' hidden="1"';
    if (info?.outline) attrs += ` outlineLevel="${info.outline}"`;
    if (info?.collapsed) attrs += ' collapsed="1"';
    const row = sheet.cells.get(r);
    if (!row || !row.size) {
      out.push(`<row${attrs}/>`);
      continue;
    }
    out.push(`<row${attrs}>`);
    const cols = [...row.keys()].sort((a, b) => a - b);
    for (const c of cols) out.push(cellXml(r, c, row.get(c)!, ctx));
    out.push('</row>');
  }
  out.push('</sheetData>');
  return out.join('');
}

function colsXml(sheet: Sheet): XmlElement | null {
  const keys = [...sheet.cols.keys()].sort((a, b) => a - b);
  if (!keys.length) return null;
  const items: XmlElement[] = [];
  const defaultChars = widthPxToChars(sheet.defaultColWidthPx);
  let start = keys[0];
  let prev = keys[0];
  let sig = JSON.stringify(sheet.cols.get(start));
  const flush = (min: number, max: number) => {
    const info = sheet.cols.get(min)!;
    items.push(el('col', {
      min: min + 1,
      max: max + 1,
      width: info.w ?? defaultChars,
      style: info.s,
      hidden: info.hidden ? 1 : undefined,
      bestFit: info.bestFit ? 1 : undefined,
      customWidth: info.w !== undefined || info.customWidth ? 1 : undefined,
      outlineLevel: info.outline,
      collapsed: info.collapsed ? 1 : undefined,
    }));
  };
  for (const k of keys.slice(1)) {
    const s = JSON.stringify(sheet.cols.get(k));
    if (k === prev + 1 && s === sig) {
      prev = k;
      continue;
    }
    flush(start, prev);
    start = k;
    prev = k;
    sig = s;
  }
  flush(start, prev);
  return el('cols', {}, items);
}

function sheetViewsXml(sheet: Sheet, active: boolean, source: XmlElement | undefined): XmlElement {
  const srcView = child(child(source, 'sheetViews'), 'sheetView');
  const attrs: Record<string, string | number | undefined> = { ...(srcView?.attrs ?? {}) };
  delete attrs.tabSelected;
  delete attrs.zoomScale;
  delete attrs.zoomScaleNormal;
  delete attrs.showGridLines;
  delete attrs.rightToLeft;
  if (active) attrs.tabSelected = 1;
  if (!sheet.showGridLines) attrs.showGridLines = 0;
  if (sheet.rightToLeft) attrs.rightToLeft = 1;
  if (Math.abs(sheet.zoom - 1) > 0.001) {
    attrs.zoomScale = Math.round(sheet.zoom * 100);
    attrs.zoomScaleNormal = Math.round(sheet.zoom * 100);
  }
  attrs.workbookViewId ??= 0;
  const kids: XmlElement[] = [];
  const { rows, cols } = sheet.freeze;
  let pane: string | undefined;
  if (rows > 0 || cols > 0) {
    pane = rows && cols ? 'bottomRight' : rows ? 'bottomLeft' : 'topRight';
    kids.push(el('pane', { xSplit: cols || undefined, ySplit: rows || undefined, topLeftCell: cellName(rows, cols), activePane: pane, state: 'frozen' }));
  }
  const sel = sheet.selection;
  const activeCell = sel ? cellName(sel.active.r, sel.active.c) : 'A1';
  const sqref = sel && sel.ranges.length ? sqrefString(sel.ranges) : activeCell;
  kids.push(el('selection', { pane, activeCell, sqref }));
  for (const k of elements(srcView)) if (!['pane', 'selection'].includes(k.name)) kids.push(cloneXml(k));
  // Ordered as attributes in the source when possible: rebuild with workbookViewId last is fine.
  return el('sheetViews', {}, [el('sheetView', attrs, kids)]);
}

function hyperlinksXml(sheet: Sheet, rels: Relationship[], path: string, pkg: OoxmlPackage): XmlElement | null {
  if (!sheet.hyperlinks.length) return null;
  const items = sheet.hyperlinks.map((h) => {
    let rid: string | undefined;
    if (h.target) {
      rid = nextId(rels);
      rels.push({ id: rid, type: REL.hyperlink, target: h.target, targetMode: 'External' });
    }
    return el('hyperlink', { ref: rangeName(h.ref), 'r:id': rid, location: h.location, tooltip: h.tooltip, display: h.display });
  });
  void path;
  void pkg;
  return el('hyperlinks', {}, items);
}

const VALIDATION_ATTRS = ['type', 'errorStyle', 'operator', 'allowBlank', 'showDropDown', 'showInputMessage', 'showErrorMessage', 'errorTitle', 'error', 'promptTitle', 'prompt', 'sqref'];

/** A `<dataValidation>` from the model; attributes the editor does not model come from the source. */
export function validationXml(v: DataValidation): XmlElement {
  const kept: Record<string, string> = {};
  for (const [k, val] of Object.entries(v.raw?.attrs ?? {})) if (!VALIDATION_ATTRS.includes(k)) kept[k] = val;
  const type = v.type && v.type !== 'none' ? v.type : undefined;
  const usesOperator = type && type !== 'list' && type !== 'custom';
  const kids: XmlElement[] = [];
  if (v.formula1 !== undefined && type) kids.push(el('formula1', {}, [txt(formulaToFile(v.formula1))]));
  if (v.formula2 !== undefined && usesOperator && (v.operator === 'between' || v.operator === 'notBetween' || !v.operator)) kids.push(el('formula2', {}, [txt(formulaToFile(v.formula2))]));
  // Extension children (x14 lists and the like) of the source element are kept.
  for (const k of elements(v.raw)) if (k.name !== 'formula1' && k.name !== 'formula2') kids.push(cloneXml(k));
  return el('dataValidation', {
    type,
    errorStyle: v.errorStyle && v.errorStyle !== 'stop' ? v.errorStyle : undefined,
    operator: usesOperator && v.operator && v.operator !== 'between' ? v.operator : undefined,
    allowBlank: v.allowBlank ? 1 : undefined,
    showDropDown: v.hideDropDown ? 1 : undefined,
    showInputMessage: v.showInputMessage ? 1 : undefined,
    showErrorMessage: v.showErrorMessage ? 1 : undefined,
    errorTitle: v.errorTitle || undefined,
    error: v.error || undefined,
    promptTitle: v.promptTitle || undefined,
    prompt: v.prompt || undefined,
    sqref: sqrefString(v.sqref),
    ...kept,
  }, kids);
}

function escapeFilterValue(s: string): string {
  return s.replace(/[*?~]/g, (m) => `~${m}`);
}

/** `<autoFilter>` with the column criteria of the model (sort state from the source is kept). */
export function autoFilterXml(af: AutoFilter): XmlElement {
  const kids: XmlElement[] = [];
  const cols = [...af.columns.keys()].filter((c) => c >= af.ref.c0 && c <= af.ref.c1).sort((a, b) => a - b);
  for (const c of cols) {
    const f = af.columns.get(c)!;
    const inner: XmlElement[] = [];
    if (f.values) {
      inner.push(el('filters', { blank: f.blanks ? 1 : undefined }, f.values.map((v) => el('filter', { val: v }))));
    } else if (f.conditions?.length) {
      const conds = f.conditions.slice(0, 2).map((cond) => {
        const v = escapeFilterValue(cond.value);
        switch (cond.op) {
          case 'contains': return el('customFilter', { val: `*${v}*` });
          case 'notContains': return el('customFilter', { operator: 'notEqual', val: `*${v}*` });
          case 'beginsWith': return el('customFilter', { val: `${v}*` });
          case 'endsWith': return el('customFilter', { val: `*${v}` });
          case 'equal': return el('customFilter', { val: v });
          default: return el('customFilter', { operator: cond.op, val: cond.value });
        }
      });
      inner.push(el('customFilters', { and: f.and && conds.length > 1 ? 1 : undefined }, conds));
    }
    if (inner.length) kids.push(el('filterColumn', { colId: c - af.ref.c0 }, inner));
  }
  for (const k of elements(af.raw)) if (k.name !== 'filterColumn') kids.push(cloneXml(k));
  const attrs: Record<string, string | number | undefined> = { ...(af.raw?.attrs ?? {}) };
  attrs.ref = rangeName(af.ref);
  return el('autoFilter', attrs, kids);
}

function nextId(rels: Relationship[]): string {
  const used = new Set(rels.map((r) => r.id));
  let n = rels.length + 1;
  while (used.has(`rId${n}`)) n++;
  return `rId${n}`;
}

function writeWorksheet(sheet: Sheet, index: number, path: string, ctx: WriteContext): string {
  const { pkg, wb } = ctx;
  const source = sheet.sourceXml;
  const active = index === wb.activeSheetIndex;
  // Relationships: keep those referenced by kept elements, regenerate the rest.
  const keptRelIds = new Set<string>();
  const kept = new Map<string, XmlElement[]>();
  for (const e of elements(source)) {
    if (OWNED.has(e.name)) continue;
    collectRelIds(e, keptRelIds);
    const list = kept.get(e.name) ?? [];
    list.push(e);
    kept.set(e.name, list);
  }
  const oldRels = source ? pkg.getRels(path) : [];
  const rels: Relationship[] = oldRels.filter((r) => keptRelIds.has(r.id));

  const parts: string[] = [];
  const add = (e: XmlElement | null | undefined) => {
    if (e) parts.push(serializeNode(e));
  };

  for (const name of WORKSHEET_ORDER) {
    switch (name) {
      case 'sheetPr': {
        const src = child(source, 'sheetPr');
        const pr = src ? cloneXml(src) : el('sheetPr');
        pr.children = pr.children.filter((c) => c.type !== 'element' || c.name !== 'tabColor');
        if (sheet.tabColor) pr.children.unshift(el('tabColor', colorAttrs(sheet.tabColor)));
        if (pr.children.length || Object.keys(pr.attrs).length) add(pr);
        break;
      }
      case 'dimension': {
        const used = sheet.usedRange();
        add(el('dimension', { ref: used ? rangeName(used) : 'A1' }));
        break;
      }
      case 'sheetViews':
        add(sheetViewsXml(sheet, active, source));
        break;
      case 'sheetFormatPr': {
        const src = child(source, 'sheetFormatPr');
        const fmt = src ? cloneXml(src) : el('sheetFormatPr');
        fmt.attrs.defaultRowHeight = String(sheet.defaultRowHeightPt);
        if (sheet.defaultColWidthPx !== 64 || fmt.attrs.defaultColWidth) fmt.attrs.defaultColWidth = String(widthPxToChars(sheet.defaultColWidthPx));
        add(fmt);
        break;
      }
      case 'cols':
        add(colsXml(sheet));
        break;
      case 'sheetData':
        parts.push(sheetDataXml(sheet, ctx));
        break;
      case 'autoFilter':
        if (sheet.autoFilter) add(autoFilterXml(sheet.autoFilter));
        break;
      case 'mergeCells':
        if (sheet.merges.length) add(el('mergeCells', { count: sheet.merges.length }, sheet.merges.map((m) => el('mergeCell', { ref: rangeName(m) }))));
        break;
      case 'conditionalFormatting':
        for (const cf of sheet.condFormats) {
          const e = cloneXml(cf.raw);
          e.attrs.sqref = sqrefString(cf.sqref);
          add(e);
        }
        break;
      case 'dataValidations':
        if (sheet.validations.length) {
          const src = child(source, 'dataValidations');
          add(el('dataValidations', { ...(src?.attrs ?? {}), count: sheet.validations.length }, sheet.validations.map(validationXml)));
        }
        break;
      case 'hyperlinks':
        add(hyperlinksXml(sheet, rels, path, pkg));
        break;
      case 'pageMargins':
        if (kept.get('pageMargins')) for (const e of kept.get('pageMargins')!) add(e);
        else add(el('pageMargins', { left: 0.7, right: 0.7, top: 0.75, bottom: 0.75, header: 0.3, footer: 0.3 }));
        break;
      case 'drawing': {
        const rid = writeDrawing(sheet, path, rels, ctx);
        if (rid) add(el('drawing', { 'r:id': rid }));
        break;
      }
      case 'legacyDrawing': {
        const rid = writeComments(sheet, path, rels, ctx);
        if (rid) add(el('legacyDrawing', { 'r:id': rid }));
        break;
      }
      default:
        for (const e of kept.get(name) ?? []) add(e);
    }
  }

  pkg.setRels(path, rels);
  const rootAttrs: Record<string, string> = {};
  for (const [k, v] of Object.entries(source?.attrs ?? {})) rootAttrs[k] = v;
  const probe = el('worksheet', rootAttrs, parts.map((p) => txt(p)));
  // Namespaces: canonical ones used by kept content plus unknown ones from the source.
  const used = new Set<string>(['', 'r']);
  for (const p of parts) for (const m of p.matchAll(/<\/?([A-Za-z][\w.-]*):|\s([A-Za-z][\w.-]*):[\w.-]+=/g)) used.add(m[1] ?? m[2]);
  for (const k of Object.keys(rootAttrs)) if (k.includes(':')) used.add(k.split(':')[0]);
  for (const p of (rootAttrs['mc:Ignorable'] ?? '').split(/\s+/)) if (p) used.add(p);
  if (ctx.usesExtension && parts.some((p) => p.includes('xe:'))) used.add('xe');
  const ns: string[] = [];
  for (const p of used) {
    if (p === 'xml' || p === 'xmlns') continue;
    const uri = NS[p] ?? ctx.wb.source?.extraNamespaces[p];
    if (uri) ns.push(p ? `xmlns:${p}="${uri}"` : `xmlns="${uri}"`);
  }
  if (!ns.some((x) => x.startsWith('xmlns='))) ns.unshift(`xmlns="${MAIN_NS}"`);
  const attrsXml = Object.entries(rootAttrs).map(([k, v]) => ` ${k}="${escapeAttr(v)}"`).join('');
  void probe;
  return `${XML_DECL}<worksheet ${ns.join(' ')}${attrsXml}>${parts.join('')}</worksheet>`;
}

// -----------------------------------------------------------------------------------------------
// Drawings (pictures)

function anchorXml(tag: string, p: { r: number; c: number; dx: number; dy: number }): XmlElement {
  return el(tag, {}, [
    el('xdr:col', {}, [txt(String(p.c))]),
    el('xdr:colOff', {}, [txt(String(Math.round(p.dx)))]),
    el('xdr:row', {}, [txt(String(p.r))]),
    el('xdr:rowOff', {}, [txt(String(Math.round(p.dy)))]),
  ]);
}

function pictureAnchor(img: SheetImage, rid: string, id: number): XmlElement {
  const cx = img.ext?.cx ?? pxToEmu(200);
  const cy = img.ext?.cy ?? pxToEmu(200);
  const pic = el('xdr:pic', {}, [
    el('xdr:nvPicPr', {}, [
      el('xdr:cNvPr', { id, name: img.name, descr: img.descr }),
      el('xdr:cNvPicPr', {}, [el('a:picLocks', { noChangeAspect: 1 })]),
    ]),
    el('xdr:blipFill', {}, [el('a:blip', { 'r:embed': rid }), el('a:stretch', {}, [el('a:fillRect')])]),
    el('xdr:spPr', {}, [
      el('a:xfrm', {}, [el('a:off', { x: 0, y: 0 }), el('a:ext', { cx, cy })]),
      el('a:prstGeom', { prst: 'rect' }, [el('a:avLst')]),
    ]),
  ]);
  if (img.anchorType === 'twoCell' && img.to) {
    return el('xdr:twoCellAnchor', { editAs: img.editAs }, [anchorXml('xdr:from', img.from), anchorXml('xdr:to', img.to), pic, el('xdr:clientData')]);
  }
  if (img.anchorType === 'absolute' && img.pos) {
    return el('xdr:absoluteAnchor', {}, [el('xdr:pos', { x: img.pos.x, y: img.pos.y }), el('xdr:ext', { cx, cy }), pic, el('xdr:clientData')]);
  }
  return el('xdr:oneCellAnchor', {}, [anchorXml('xdr:from', img.from), el('xdr:ext', { cx, cy }), pic, el('xdr:clientData')]);
}

function writeDrawing(sheet: Sheet, sheetPath: string, sheetRels: Relationship[], ctx: WriteContext): string | null {
  const { pkg, wb } = ctx;
  const srcDrawing = child(sheet.sourceXml, 'drawing');
  const srcRel = srcDrawing && sheet.sourcePath ? pkg.getRels(sheet.sourcePath).find((r) => r.id === attr(srcDrawing, 'r:id')) : undefined;
  const srcPath = srcRel ? pkg.resolveTarget(sheetPath, srcRel.target) : undefined;
  if (!sheet.images.length && !sheet.otherAnchors.length && !sheet.charts.length && !sheet.shapes.length) {
    // The drawing and the charts only it used go.
    if (srcPath) removeOrphans(pkg, [srcPath]);
    return null;
  }
  const path = srcPath ?? pkg.uniquePartName('xl/drawings', 'drawing', 'xml');
  const oldRels = srcPath ? pkg.getRels(srcPath) : [];
  const rels: Relationship[] = [];
  const keep = (id: string) => {
    const r = oldRels.find((x) => x.id === id);
    if (r && !rels.some((x) => x.id === id)) rels.push(r);
  };
  const keepable = (anchor: XmlElement) => {
    const ids = new Set<string>();
    collectRelIds(anchor, ids);
    return [...ids].every((id) => oldRels.some((r) => r.id === id)) ? ids : null;
  };
  const taken = new Set(oldRels.map((r) => r.id));
  const newRelId = () => {
    let n = 1;
    while (taken.has(`rId${n}`) || rels.some((r) => r.id === `rId${n}`)) n++;
    taken.add(`rId${n}`);
    return `rId${n}`;
  };
  // Anchors in stacking order (later ones are drawn on top).
  const entries: { z: number; anchor: XmlElement }[] = [];
  let maxId = 1;
  const scanIds = (e: XmlElement) => {
    for (const m of serializeNode(e).matchAll(/<xdr:cNvPr\b[^>]*\bid="(\d+)"/g)) maxId = Math.max(maxId, Number(m[1]));
  };
  sheet.otherAnchors.forEach((a, i) => {
    const ids = new Set<string>();
    collectRelIds(a, ids);
    ids.forEach(keep);
    scanIds(a);
    entries.push({ z: -1e6 + i, anchor: a });
  });
  // First the objects written back as they were (their ids are reserved), then the new ones.
  const newImages: SheetImage[] = [];
  sheet.images.forEach((img, i) => {
    const ids = img.raw ? keepable(img.raw) : null;
    if (img.raw && ids) {
      ids.forEach(keep);
      scanIds(img.raw);
      entries.push({ z: img.z ?? i, anchor: img.raw });
    } else newImages.push(img);
  });
  const newCharts: SheetChart[] = [];
  for (const chart of sheet.charts) {
    const src = chart.source;
    const ref = src && !chart.dirty ? chartReference(src.anchor) : null;
    const old = ref ? oldRels.find((r) => r.id === ref.rid) : undefined;
    const ids = src && old && pkg.resolveTarget(path, old.target) === src.chartPath && pkg.has(src.chartPath) ? keepable(src.anchor) : null;
    if (src && ids) {
      ids.forEach(keep);
      scanIds(src.anchor);
      entries.push({ z: chart.z, anchor: repositionAnchor(src.anchor, chart.from, chart.to) });
    } else newCharts.push(chart);
  }
  const newShapes: SheetShape[] = [];
  for (const shape of sheet.shapes) {
    const ids = shape.source && !shape.dirty ? keepable(shape.source.anchor) : null;
    if (shape.source && ids) {
      ids.forEach(keep);
      scanIds(shape.source.anchor);
      entries.push({ z: shape.z, anchor: repositionAnchor(shape.source.anchor, shape.from, shape.to) });
    } else newShapes.push(shape);
  }
  for (const img of newImages) {
    const ext = extensionForMime(img.mime);
    const mediaPath = pkg.uniquePartName('xl/media', 'image', ext);
    pkg.set(mediaPath, img.data);
    pkg.ensureDefaultContentType(ext, IMAGE_CONTENT_TYPES[ext] ?? img.mime);
    const rid = newRelId();
    rels.push({ id: rid, type: REL.image, target: pkg.relativeTarget(path, mediaPath) });
    entries.push({ z: img.z ?? sheet.images.indexOf(img), anchor: pictureAnchor(img, rid, ++maxId) });
  }
  for (const chart of newCharts) {
    const chartex = CHARTEX_TYPES.has(chart.type);
    const partPath = pkg.uniquePartName('xl/charts', chartex ? 'chartEx' : 'chart', 'xml');
    pkg.setText(partPath, chartex ? writeChartexPart(wb, chart) : writeChartPart(wb, chart), chartex ? CT_CHARTEX : CT_CHART);
    const rid = newRelId();
    rels.push({ id: rid, type: chartex ? REL_CHARTEX : REL_CHART, target: pkg.relativeTarget(path, partPath) });
    entries.push({ z: chart.z, anchor: chartAnchor(sheet, chart, rid, ++maxId, chartex) });
  }
  for (const shape of newShapes) entries.push({ z: shape.z, anchor: shapeAnchor(sheet, shape, ++maxId) });
  entries.sort((a, b) => a.z - b.z);
  const root = el('xdr:wsDr', {}, entries.map((e) => e.anchor));
  const ns = namespacesFor(root, ctx.wb.source?.extraNamespaces ?? {}, false);
  ns.xdr = NS.xdr;
  ns.a = NS.a;
  ns.r = NS.r;
  pkg.setText(path, serializeXml(root, { namespaces: ns }), CT.drawing);
  pkg.setRels(path, rels);
  // Charts of the file that were deleted or rewritten.
  const dropped = oldRels
    .filter((r) => r.targetMode !== 'External' && !rels.some((k) => k.id === r.id) && (relTypeIs(r.type, REL_CHART) || r.type === REL_CHARTEX))
    .map((r) => pkg.resolveTarget(path, r.target))
    .filter((target) => !isReferenced(pkg, target));
  if (dropped.length) removeOrphans(pkg, dropped);
  const existing = sheetRels.find((r) => relTypeIs(r.type, REL.drawing) && pkg.resolveTarget(sheetPath, r.target) === path);
  if (existing) return existing.id;
  const rid = nextId(sheetRels);
  sheetRels.push({ id: rid, type: REL.drawing, target: pkg.relativeTarget(sheetPath, path) });
  return rid;
}

/** Whether any relationship of the package still points at a part. */
function isReferenced(pkg: OoxmlPackage, target: string): boolean {
  for (const file of pkg.files.keys()) {
    if (!file.endsWith('.rels')) continue;
    const owner = file.replace(/_rels\/([^/]+)\.rels$/, '$1');
    if (pkg.getRels(owner === file ? '' : owner).some((x) => x.targetMode !== 'External' && pkg.resolveTarget(owner, x.target) === target)) return true;
  }
  return false;
}

// -----------------------------------------------------------------------------------------------
// Comments (legacy notes, threaded comments and the VML that shows them)

function noteShape(t: CommentThread, shapeId: number, visible: boolean): string {
  const anchor = `${t.c + 1}, 15, ${Math.max(0, t.r - 1)}, 10, ${t.c + 3}, 15, ${t.r + 3}, 4`;
  return `<v:shape id="_x0000_s${shapeId}" type="#_x0000_t202" style="position:absolute;margin-left:59.25pt;margin-top:1.5pt;width:108pt;height:59.25pt;z-index:${shapeId - 1024};visibility:${visible ? 'visible' : 'hidden'}" fillcolor="#ffffe1" o:insetmode="auto"><v:fill color2="#ffffe1"/><v:shadow on="t" color="black" obscured="t"/><v:path o:connecttype="none"/><v:textbox style="mso-direction-alt:auto"><div style="text-align:left"></div></v:textbox><x:ClientData ObjectType="Note"><x:MoveWithCells/><x:SizeWithCells/><x:Anchor>${anchor}</x:Anchor><x:AutoFill>False</x:AutoFill><x:Row>${t.r}</x:Row><x:Column>${t.c}</x:Column></x:ClientData></v:shape>`;
}

function personFor(ctx: WriteContext, author: string, personId?: string): string {
  if (personId && ctx.persons.has(personId)) return personId;
  for (const p of ctx.persons.values()) if (p.displayName === author) return p.id;
  const id = newGuid();
  ctx.persons.set(id, { id, displayName: author || 'Guest', userId: author || 'Guest', providerId: 'None' });
  return id;
}

function ensureGuid(id: string): string {
  return /^\{[0-9A-F-]{36}\}$/i.test(id) ? id.toUpperCase() : newGuid();
}

function writeComments(sheet: Sheet, sheetPath: string, sheetRels: Relationship[], ctx: WriteContext): string | null {
  const { pkg } = ctx;
  // Remove the parts we regenerate.
  if (sheet.sourcePath) {
    for (const r of pkg.getRels(sheet.sourcePath)) {
      if (relTypeIs(r.type, REL.comments) || relTypeIs(r.type, REL.threadedComment)) pkg.delete(pkg.resolveTarget(sheetPath, r.target));
    }
    const legacy = child(sheet.sourceXml, 'legacyDrawing');
    const vmlRel = legacy ? pkg.getRels(sheet.sourcePath).find((r) => r.id === attr(legacy, 'r:id')) : undefined;
    if (vmlRel) pkg.delete(pkg.resolveTarget(sheetPath, vmlRel.target));
  }
  const threads = sheet.comments.filter((t) => t.comments.length);
  if (!threads.length) return null;

  const authors: string[] = [];
  const authorIndex = (a: string) => {
    let i = authors.indexOf(a);
    if (i < 0) {
      i = authors.length;
      authors.push(a);
    }
    return i;
  };
  const commentEls: string[] = [];
  const threadedEls: string[] = [];
  const shapes: string[] = [];
  let shapeId = 1025;
  for (const t of threads) {
    const ref = cellName(t.r, t.c);
    if (t.kind === 'threaded') {
      const rootId = ensureGuid(t.comments[0].id);
      t.comments[0].id = rootId;
      t.id = rootId;
      t.comments.forEach((cm, i) => {
        const id = i === 0 ? rootId : ensureGuid(cm.id);
        cm.id = id;
        const personId = personFor(ctx, cm.author, cm.personId);
        cm.personId = personId;
        const done = i === 0 && t.resolved ? ' done="1"' : '';
        const parent = i === 0 ? '' : ` parentId="${rootId}"`;
        const dT = excelDateTime(cm.created ?? new Date().toISOString());
        threadedEls.push(`<threadedComment ref="${ref}" dT="${escapeAttr(dT)}" personId="${personId}" id="${id}"${parent}${done}><text>${escapeText(stripInvalidXmlChars(cm.text))}</text></threadedComment>`);
      });
      const body = ['[Threaded comment]', '', 'Your version of Excel allows you to read this threaded comment; however, any edits to it will get removed if the file is opened in a newer version of Excel. Learn more: https://go.microsoft.com/fwlink/?linkid=870924', '', `Comment:\n    ${t.comments[0].text}`, ...t.comments.slice(1).map((c) => `Reply:\n    ${c.text}`)].join('\n');
      // A single run (rather than a bare <t>) is read by more consumers; both are valid.
      commentEls.push(`<comment ref="${ref}" authorId="${authorIndex(`tc=${rootId}`)}" shapeId="0" xr:uid="${rootId}"><text><r>${textElementXml(body)}</r></text></comment>`);
    } else {
      const c = t.comments[0];
      const runs = c.author
        ? `<r><rPr><b/><sz val="9"/><color indexed="81"/><rFont val="Tahoma"/><family val="2"/></rPr>${textElementXml(`${c.author}:`)}</r><r><rPr><sz val="9"/><color indexed="81"/><rFont val="Tahoma"/><family val="2"/></rPr>${textElementXml(`\n${c.text}`)}</r>`
        : `<r>${textElementXml(c.text)}</r>`;
      commentEls.push(`<comment ref="${ref}" authorId="${authorIndex(c.author)}" shapeId="0"><text>${runs}</text></comment>`);
    }
    const id = shapeId++;
    if (t.vmlShape && t.kind === 'note') {
      shapes.push(t.vmlShape.replace(/id="_x0000_s\d+"/, `id="_x0000_s${id}"`).replace(/<x:Row>\d+<\/x:Row>/, `<x:Row>${t.r}</x:Row>`).replace(/<x:Column>\d+<\/x:Column>/, `<x:Column>${t.c}</x:Column>`));
    } else shapes.push(noteShape(t, id, false));
  }

  const commentsPath = pkg.uniquePartName('xl', 'comments', 'xml');
  pkg.setText(
    commentsPath,
    `${XML_DECL}<comments xmlns="${MAIN_NS}" xmlns:mc="${NS.mc}" mc:Ignorable="xr" xmlns:xr="${NS.xr}"><authors>${authors.map((a) => `<author>${escapeText(a)}</author>`).join('')}</authors><commentList>${commentEls.join('')}</commentList></comments>`,
    CT.comments,
  );
  sheetRels.push({ id: nextId(sheetRels), type: REL.comments, target: pkg.relativeTarget(sheetPath, commentsPath) });

  if (threadedEls.length) {
    const tcPath = pkg.uniquePartName('xl/threadedComments', 'threadedComment', 'xml');
    pkg.setText(tcPath, `${XML_DECL}<ThreadedComments xmlns="${THREADED_NS}" xmlns:x="${MAIN_NS}">${threadedEls.join('')}</ThreadedComments>`, CT.threadedComments);
    sheetRels.push({ id: nextId(sheetRels), type: REL.threadedComment, target: pkg.relativeTarget(sheetPath, tcPath) });
  }

  const vmlPath = pkg.uniquePartName('xl/drawings', 'vmlDrawing', 'vml');
  const vml = `<xml xmlns:v="urn:schemas-microsoft-com:vml" xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:x="urn:schemas-microsoft-com:office:excel"><o:shapelayout v:ext="edit"><o:idmap v:ext="edit" data="${ctx.wb.sheets.indexOf(sheet) + 1}"/></o:shapelayout><v:shapetype id="_x0000_t202" coordsize="21600,21600" o:spt="202" path="m,l,21600r21600,l21600,xe"><v:stroke joinstyle="miter"/><v:path gradientshapeok="t" o:connecttype="rect"/></v:shapetype>${shapes.join('')}</xml>`;
  pkg.set(vmlPath, strToU8(vml));
  pkg.ensureDefaultContentType('vml', CT.vml);
  const rid = nextId(sheetRels);
  sheetRels.push({ id: rid, type: REL.vmlDrawing, target: pkg.relativeTarget(sheetPath, vmlPath) });
  return rid;
}

// -----------------------------------------------------------------------------------------------
// Workbook

const WORKBOOK_ORDER = [
  'fileVersion', 'fileSharing', 'workbookPr', 'mc:AlternateContent', 'xr:revisionPtr', 'workbookProtection', 'bookViews',
  'sheets', 'functionGroups', 'externalReferences', 'definedNames', 'calcPr', 'oleSize', 'customWorkbookViews',
  'pivotCaches', 'smartTagPr', 'smartTagTypes', 'webPublishing', 'fileRecoveryPr', 'webPublishObjects', 'extLst',
];

function metadataXml(): string {
  return `${XML_DECL}<metadata xmlns="${MAIN_NS}" xmlns:xda="${NS.xda}"><metadataTypes count="1"><metadataType name="XLDAPR" minSupportedVersion="120000" copy="1" pasteAll="1" pasteValues="1" merge="1" splitFirst="1" rowColShift="1" clearFormats="1" clearComments="1" assign="1" coerce="1" cellMeta="1"/></metadataTypes><futureMetadata name="XLDAPR" count="1"><bk><extLst><ext uri="{bdbb8cdc-fa1e-496e-a857-3c3f30c029c3}"><xda:dynamicArrayProperties fDynamic="1" fCollapsed="0"/></ext></extLst></bk></futureMetadata><cellMetadata count="1"><bk><rc t="1" v="0"/></bk></cellMetadata></metadata>`;
}

/** cm index of the dynamic array metadata in an existing metadata part, or null. */
function existingDynamicCm(pkg: OoxmlPackage, wbPath: string): { cm: number; path: string } | null {
  const rel = pkg.findRel(wbPath, REL.sheetMetadata);
  if (!rel) return null;
  const path = pkg.resolveTarget(wbPath, rel.target);
  const text = pkg.getText(path);
  if (!text) return null;
  const root = parseXml(text).root;
  const types = elements(child(root, 'metadataTypes')).map((t) => attr(t, 'name'));
  const blocks = elements(child(root, 'cellMetadata'));
  for (let i = 0; i < blocks.length; i++) {
    const rc = child(blocks[i], 'rc');
    const t = Number(attr(rc, 't'));
    if (types[t - 1] === 'XLDAPR') return { cm: i + 1, path };
  }
  return null;
}

function workbookXml(ctx: WriteContext, sheetRelIds: string[]): string {
  const { wb } = ctx;
  const src = wb.source?.workbookXml;
  const children = new Map<string, XmlElement[]>();
  for (const e of elements(src)) {
    const list = children.get(e.name) ?? [];
    list.push(e);
    children.set(e.name, list);
  }
  const out: XmlElement[] = [];
  for (const name of WORKBOOK_ORDER) {
    switch (name) {
      case 'workbookPr': {
        const pr = children.get('workbookPr')?.[0] ? cloneXml(children.get('workbookPr')![0]) : el('workbookPr', { defaultThemeVersion: 166925 });
        if (wb.date1904) pr.attrs.date1904 = '1';
        else delete pr.attrs.date1904;
        out.push(pr);
        break;
      }
      case 'bookViews': {
        const views = children.get('bookViews')?.[0] ? cloneXml(children.get('bookViews')![0]) : el('bookViews', {}, [el('workbookView', { xWindow: 0, yWindow: 0, windowWidth: 28800, windowHeight: 12300 })]);
        const first = elements(views)[0];
        if (first) {
          if (wb.activeSheetIndex) first.attrs.activeTab = String(wb.activeSheetIndex);
          else delete first.attrs.activeTab;
          const firstVisible = wb.sheets.findIndex((s) => s.state === 'visible');
          if (firstVisible > 0) first.attrs.firstSheet = String(firstVisible);
          else delete first.attrs.firstSheet;
        }
        out.push(views);
        break;
      }
      case 'sheets':
        out.push(el('sheets', {}, wb.sheets.map((s, i) => el('sheet', { name: s.name, sheetId: s.sheetId, state: s.state === 'visible' ? undefined : s.state, 'r:id': sheetRelIds[i] }))));
        break;
      case 'definedNames': {
        const names = wb.names.filter((n) => n.name);
        if (names.length) {
          out.push(el('definedNames', {}, names.map((n) => el('definedName', { name: n.name, comment: n.comment, ...(n.extra ?? {}), localSheetId: n.localSheet, hidden: n.hidden ? 1 : undefined }, [txt(formulaToFile(n.formula))]))));
        }
        break;
      }
      case 'calcPr': {
        const pr = children.get('calcPr')?.[0] ? cloneXml(children.get('calcPr')![0]) : el('calcPr', { calcId: 191029 });
        pr.attrs.fullCalcOnLoad = '1';
        out.push(pr);
        break;
      }
      default:
        for (const e of children.get(name) ?? []) out.push(e);
    }
  }
  // Keep unknown elements (not in the schema order list) at the end before extLst.
  for (const [name, list] of children) if (!WORKBOOK_ORDER.includes(name)) out.splice(Math.max(0, out.length - (children.has('extLst') ? 1 : 0)), 0, ...list);
  const root = el('workbook', { ...(src?.attrs ?? {}) }, out);
  const ns = namespacesFor(root, wb.source?.extraNamespaces ?? {});
  ns.r = REL_NS;
  return serializeXml(root, { namespaces: ns });
}

/** Parts reachable only from removed sheets are deleted (charts, drawings, comments...). */
function removeOrphans(pkg: OoxmlPackage, removed: string[]): void {
  const queue = [...removed];
  while (queue.length) {
    const path = queue.pop()!;
    if (!pkg.has(path)) continue;
    const rels = pkg.getRels(path).filter((r) => r.targetMode !== 'External');
    pkg.delete(path);
    for (const r of rels) {
      const target = pkg.resolveTarget(path, r.target);
      // Still referenced elsewhere?
      let referenced = false;
      for (const file of pkg.files.keys()) {
        if (!file.endsWith('.rels')) continue;
        const owner = file.replace(/_rels\/([^/]+)\.rels$/, '$1');
        if (pkg.getRels(owner === file ? '' : owner).some((x) => x.targetMode !== 'External' && pkg.resolveTarget(owner, x.target) === target)) {
          referenced = true;
          break;
        }
      }
      if (!referenced) queue.push(target);
    }
  }
}

export interface WriteOptions {
  /** false drops VBA macros (saving an .xlsm workbook as .xlsx). */
  macros?: boolean;
}

const MACRO_WORKBOOK_CT = 'application/vnd.ms-excel.sheet.macroEnabled.main+xml';

export function writeXlsx(wb: Workbook, opts: WriteOptions = {}): Uint8Array {
  const pkg = wb.source ? wb.source.pkg.clone() : basePackage();
  if (opts.macros === false && wb.source) {
    const wbPathSrc = wb.source.workbookPath;
    if (pkg.contentTypes.typeOf(wbPathSrc) === MACRO_WORKBOOK_CT || /macroEnabled/i.test(pkg.contentTypes.typeOf(wbPathSrc) ?? '')) {
      pkg.contentTypes.overrides.set(wbPathSrc, CT.workbook);
      const rels = pkg.getRels(wbPathSrc);
      const vba = rels.filter((r) => /vbaProject/i.test(r.type));
      for (const r of vba) pkg.delete(pkg.resolveTarget(wbPathSrc, r.target));
      if (vba.length) pkg.setRels(wbPathSrc, rels.filter((r) => !vba.includes(r)));
    }
  }
  // Chart references follow renamed sheets and inserted or deleted rows and columns.
  for (const [path, refs] of wb.chartRefs) {
    const xml = pkg.has(path) ? pkg.getText(path) : undefined;
    if (!xml) continue;
    const updated = writeChartFormulas(xml, refs);
    if (updated !== xml) pkg.setText(path, updated);
  }
  const wbPath = wb.source?.workbookPath ?? 'xl/workbook.xml';
  const ctx: WriteContext = {
    pkg,
    wb,
    wbPath,
    sst: new SharedStrings(),
    dynamicCm: null,
    needsDynamicMetadata: false,
    usesExtension: false,
    persons: new Map(wb.persons.map((p) => [p.id, { ...p }])),
  };

  // calcChain is rebuilt by Excel; it would be stale.
  const calc = pkg.findRel(wbPath, REL.calcChain);
  if (calc) pkg.delete(pkg.resolveTarget(wbPath, calc.target));
  const meta = existingDynamicCm(pkg, wbPath);
  if (meta) ctx.dynamicCm = meta.cm;

  // Styles first (cell style indexes are kept as they are).
  const styles = writeStyles(wb.styles, wb.source?.stylesXml, wb.dxfs);
  ctx.usesExtension = styles.usesExtension;
  const stylesRel = pkg.findRel(wbPath, REL.styles);
  const stylesPath = stylesRel ? pkg.resolveTarget(wbPath, stylesRel.target) : 'xl/styles.xml';
  const stylesNs = namespacesFor(styles.root, wb.source?.extraNamespaces ?? {});
  if (styles.usesExtension) {
    stylesNs.xe = NS.xe;
    stylesNs.mc = NS.mc;
    const ign = new Set((styles.root.attrs['mc:Ignorable'] ?? '').split(/\s+/).filter(Boolean));
    ign.add('xe');
    styles.root.attrs['mc:Ignorable'] = [...ign].join(' ');
  }
  pkg.setText(stylesPath, serializeXml(styles.root, { namespaces: stylesNs }), CT.styles);

  // Worksheets.
  const oldRels = pkg.getRels(wbPath);
  const sheetRelIds: string[] = [];
  const newRels: Relationship[] = oldRels.filter((r) => !relTypeIs(r.type, REL.worksheet) && !relTypeIs(r.type, REL.chartsheet) && !relTypeIs(r.type, REL.calcChain));
  const keptPaths = new Set<string>();
  const usedPaths = new Set<string>();
  wb.sheets.forEach((sheet) => {
    if (sheet.sourcePath) usedPaths.add(sheet.sourcePath);
  });
  wb.sheets.forEach((sheet, i) => {
    let path = sheet.sourcePath;
    if (!path || (keptPaths.has(path))) {
      let n = 1;
      while (pkg.has(`xl/worksheets/sheet${n}.xml`) || usedPaths.has(`xl/worksheets/sheet${n}.xml`) || keptPaths.has(`xl/worksheets/sheet${n}.xml`)) n++;
      path = `xl/worksheets/sheet${n}.xml`;
    }
    keptPaths.add(path);
    const old = oldRels.find((r) => pkg.resolveTarget(wbPath, r.target) === path);
    const id = old?.id && !newRels.some((r) => r.id === old.id) ? old.id : nextId([...newRels, ...oldRels]);
    newRels.push({ id, type: sheet.chartsheet ? REL.chartsheet : REL.worksheet, target: pkg.relativeTarget(wbPath, path) });
    sheetRelIds.push(id);
    if (!sheet.chartsheet) pkg.setText(path, writeWorksheet(sheet, i, path, ctx), CT.worksheet);
  });
  const removed = oldRels
    .filter((r) => relTypeIs(r.type, REL.worksheet) || relTypeIs(r.type, REL.chartsheet))
    .map((r) => pkg.resolveTarget(wbPath, r.target))
    .filter((p) => !keptPaths.has(p));

  // Shared strings.
  const sstRel = newRels.find((r) => relTypeIs(r.type, REL.sharedStrings));
  const sstPath = sstRel ? pkg.resolveTarget(wbPath, sstRel.target) : 'xl/sharedStrings.xml';
  pkg.setText(sstPath, ctx.sst.xml(), CT.sharedStrings);
  if (!sstRel) newRels.push({ id: nextId(newRels), type: REL.sharedStrings, target: pkg.relativeTarget(wbPath, sstPath) });
  if (!newRels.some((r) => relTypeIs(r.type, REL.styles))) newRels.push({ id: nextId(newRels), type: REL.styles, target: pkg.relativeTarget(wbPath, stylesPath) });

  // Dynamic array metadata.
  if (ctx.needsDynamicMetadata && !meta) {
    const path = 'xl/metadata.xml';
    pkg.setText(path, metadataXml(), CT.metadata);
    newRels.push({ id: nextId(newRels), type: REL.sheetMetadata, target: pkg.relativeTarget(wbPath, path) });
  }

  // People of threaded comments.
  const personRel = newRels.find((r) => relTypeIs(r.type, REL.person));
  if (ctx.persons.size) {
    const path = personRel ? pkg.resolveTarget(wbPath, personRel.target) : 'xl/persons/person.xml';
    const people = [...ctx.persons.values()].map((p) => `<person displayName="${escapeAttr(p.displayName)}" id="${p.id}" userId="${escapeAttr(p.userId ?? p.displayName)}" providerId="${escapeAttr(p.providerId ?? 'None')}"/>`);
    pkg.setText(path, `${XML_DECL}<personList xmlns="${THREADED_NS}" xmlns:x="${MAIN_NS}">${people.join('')}</personList>`, CT.person);
    if (!personRel) newRels.push({ id: nextId(newRels), type: REL.person, target: pkg.relativeTarget(wbPath, path) });
    wb.persons = [...ctx.persons.values()];
  }

  pkg.setText(wbPath, workbookXml(ctx, sheetRelIds), wb.source ? pkg.contentTypes.typeOf(wbPath) ?? CT.workbook : CT.workbook);
  pkg.setRels(wbPath, newRels);
  removeOrphans(pkg, removed);

  // Theme content type (new packages) and document properties.
  if (pkg.has('xl/theme/theme1.xml')) pkg.contentTypes.overrides.set('xl/theme/theme1.xml', CT.theme);
  writeDocProps(pkg, wb);
  pkg.ensureDefaultContentType('rels', CT.rels);
  pkg.ensureDefaultContentType('xml', CT.xml);
  return pkg.toBytes();
}

function writeDocProps(pkg: OoxmlPackage, wb: Workbook): void {
  const now = new Date().toISOString().replace(/\.\d+Z$/, 'Z');
  const coreRel = pkg.findRel('', REL.coreProps);
  const corePath = coreRel ? pkg.resolveTarget('', coreRel.target) : 'docProps/core.xml';
  const existing = pkg.getText(corePath);
  if (existing) {
    const updated = /<dcterms:modified\b/.test(existing)
      ? existing.replace(/(<dcterms:modified\b[^>]*>)[^<]*(<\/dcterms:modified>)/, `$1${now}$2`)
      : existing.replace(/<\/cp:coreProperties>/, `<dcterms:modified xsi:type="dcterms:W3CDTF">${now}</dcterms:modified></cp:coreProperties>`);
    pkg.setText(corePath, updated, CT.coreProps);
  } else {
    pkg.setText(
      corePath,
      `${XML_DECL}<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" xmlns:dcmitype="http://purl.org/dc/dcmitype/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">${wb.props.title ? `<dc:title>${escapeText(wb.props.title)}</dc:title>` : ''}<dc:creator>${escapeText(wb.props.creator ?? 'XLSX Editor')}</dc:creator><dcterms:created xsi:type="dcterms:W3CDTF">${now}</dcterms:created><dcterms:modified xsi:type="dcterms:W3CDTF">${now}</dcterms:modified></cp:coreProperties>`,
      CT.coreProps,
    );
    if (!coreRel) pkg.addRel('', REL.coreProps, corePath);
  }
  const appRel = pkg.findRel('', REL.extendedProps);
  if (!appRel || !pkg.has(pkg.resolveTarget('', appRel.target))) {
    const path = 'docProps/app.xml';
    pkg.setText(path, `${XML_DECL}<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties" xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes"><Application>Microsoft Excel</Application><DocSecurity>0</DocSecurity><ScaleCrop>false</ScaleCrop><AppVersion>16.0300</AppVersion></Properties>`, CT.extendedProps);
    if (!appRel) pkg.addRel('', REL.extendedProps, path);
  }
}

export { quoteSheetName };
export type { Range };
