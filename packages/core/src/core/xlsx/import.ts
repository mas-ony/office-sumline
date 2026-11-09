/** Reads an XLSX package into the editor model, keeping what it does not understand. */
import { strFromU8 } from 'fflate';
import { readPackageChartRefs } from './chartRefs';
import { defaultTheme, THEME_SLOTS, type Theme } from '../model/color';
import {
  Sheet, widthCharsToPx, type ColumnFilter, type CommentThread, type ConditionalFormat, type DataValidation, type FilterCondition,
  type FilterOperator, type Hyperlink, type RichTextRun, type SheetImage,
} from '../model/sheet';
import type { FontStyle, UnderlineType } from '../model/style';
import { Workbook, type DefinedName, type Person } from '../model/workbook';
import { parseCellRef, parseRange, parseSqref } from '../refs';
import {
  attr, boolAttr, child, childVal, childrenNamed, elements, findAll, numAttr, parseXml, textContent, type XmlElement,
} from '../xml/xml';
import { CT, OoxmlPackage, REL, relTypeIs } from './package';
import { parseSheetData, type SharedString } from './sheetData';
import { readColor, readStyles } from './styles';
import { formulaFromFile } from './formulaText';
import { IMAGE_CONTENT_TYPES, type Relationship } from './package';
import { anchorContent, anchorPoints, chartReference, readChartexPart, readChartPart, readShape } from './drawingObjects';
import { newDrawingId } from '../model/drawing';

export interface ImportOptions {
  /** File name, used by CELL("filename"). */
  fileName?: string;
}

function parse(pkg: OoxmlPackage, path: string): ReturnType<typeof parseXml> | undefined {
  const text = pkg.getText(path);
  if (text === undefined) return undefined;
  return parseXml(text);
}

// -----------------------------------------------------------------------------------------------
// Theme

function readTheme(root: XmlElement | undefined): Theme {
  const theme = defaultTheme();
  if (!root) return theme;
  const scheme = findAll(root, 'a:clrScheme')[0];
  if (scheme) {
    for (const slot of THEME_SLOTS) {
      const e = child(scheme, `a:${slot}`);
      if (!e) continue;
      const srgb = child(e, 'a:srgbClr');
      const sys = child(e, 'a:sysClr');
      const val = attr(srgb, 'val') ?? attr(sys, 'lastClr');
      if (val) theme.colors[slot] = val.toUpperCase();
    }
  }
  const fonts = findAll(root, 'a:fontScheme')[0];
  const major = attr(child(child(fonts, 'a:majorFont'), 'a:latin'), 'typeface');
  const minor = attr(child(child(fonts, 'a:minorFont'), 'a:latin'), 'typeface');
  if (major) theme.majorFont = major;
  if (minor) theme.minorFont = minor;
  return theme;
}

// -----------------------------------------------------------------------------------------------
// Shared strings and rich text

function runFont(rPr: XmlElement | undefined): Partial<FontStyle> | undefined {
  if (!rPr) return undefined;
  const f: Partial<FontStyle> = {};
  const name = childVal(rPr, 'rFont') ?? childVal(rPr, 'name');
  if (name) f.name = name;
  const sz = childVal(rPr, 'sz');
  if (sz) f.size = Number(sz);
  const flag = (n: string) => {
    const c = child(rPr, n);
    if (!c) return undefined;
    const v = attr(c, 'val');
    return v === undefined || v === '1' || v === 'true';
  };
  const b = flag('b');
  if (b !== undefined) f.bold = b;
  const i = flag('i');
  if (i !== undefined) f.italic = i;
  const s = flag('strike');
  if (s !== undefined) f.strike = s;
  const u = child(rPr, 'u');
  if (u) {
    const v = attr(u, 'val') ?? 'single';
    if (v !== 'none') f.underline = v as UnderlineType;
  }
  const color = readColor(child(rPr, 'color'));
  if (color) f.color = color;
  const va = childVal(rPr, 'vertAlign');
  if (va === 'superscript' || va === 'subscript') f.vertAlign = va;
  return Object.keys(f).length ? f : undefined;
}

const NS_DECL = 'xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"';

/** Parses the content of <si> or <is>. */
export function parseRichText(inner: string): SharedString {
  const clean = inner.replace(/<(\w+:)?rPh\b[\s\S]*?<\/(\w+:)?rPh>/g, '').replace(/<(\w+:)?phoneticPr\b[^>]*\/>/g, '');
  if (!/<(\w+:)?r[\s>]/.test(clean)) {
    const m = /<(?:\w+:)?t(?:\s[^>]*)?>([\s\S]*?)<\/(?:\w+:)?t>/.exec(clean);
    if (!m) return { text: '' };
    return { text: decodeText(m[1]) };
  }
  const root = parseXml(`<si ${NS_DECL}>${clean}</si>`).root;
  const runs: RichTextRun[] = [];
  let text = '';
  for (const e of elements(root)) {
    if (e.name === 't') {
      const t = textContent(e);
      text += t;
      runs.push({ text: t });
    } else if (e.name === 'r') {
      const t = textContent(child(e, 't') ?? e);
      text += t;
      const font = runFont(child(e, 'rPr'));
      runs.push(font ? { text: t, font } : { text: t });
    }
  }
  return { text, runs: runs.length > 1 || runs.some((r) => r.font) ? runs : undefined };
}

function decodeText(s: string): string {
  return s
    .replace(/&(#x[0-9a-fA-F]+|#[0-9]+|lt|gt|amp|quot|apos);/g, (_m, ent: string) => {
      switch (ent) {
        case 'lt': return '<';
        case 'gt': return '>';
        case 'amp': return '&';
        case 'quot': return '"';
        case 'apos': return "'";
        default: return String.fromCodePoint(ent[1] === 'x' ? parseInt(ent.slice(2), 16) : parseInt(ent.slice(1), 10));
      }
    })
    .replace(/_x([0-9A-Fa-f]{4})_/g, (_m, hex: string) => String.fromCharCode(parseInt(hex, 16)));
}

function readSharedStrings(text: string | undefined): SharedString[] {
  if (!text) return [];
  const out: SharedString[] = [];
  const re = /<(?:\w+:)?si>([\s\S]*?)<\/(?:\w+:)?si>|<(?:\w+:)?si\/>/g;
  let m: RegExpExecArray | null;
  while ((m = re.exec(text))) out.push(m[1] === undefined ? { text: '' } : parseRichText(m[1]));
  return out;
}

// -----------------------------------------------------------------------------------------------
// Worksheets

function splitSheetXml(xml: string): { rest: string; data: string } {
  const open = /<((?:\w+:)?)sheetData\b[^>]*?(\/?)>/.exec(xml);
  if (!open) return { rest: xml, data: '' };
  if (open[2] === '/') return { rest: xml, data: '' };
  const closeTag = `</${open[1]}sheetData>`;
  const close = xml.indexOf(closeTag, open.index);
  if (close < 0) return { rest: xml, data: '' };
  const data = xml.slice(open.index + open[0].length, close);
  const rest = `${xml.slice(0, open.index)}<${open[1]}sheetData/>${xml.slice(close + closeTag.length)}`;
  return { rest, data };
}

interface SheetContext {
  pkg: OoxmlPackage;
  shared: SharedString[];
  dynamicCm: Set<number>;
  persons: Person[];
  wb: Workbook;
}

function readWorksheet(sheet: Sheet, path: string, ctx: SheetContext): void {
  const text = ctx.pkg.getText(path);
  if (text === undefined) return;
  const { rest, data } = splitSheetXml(text);
  const doc = parseXml(rest);
  const root = doc.root;
  sheet.sourceXml = root;
  sheet.sourcePath = path;
  parseSheetData(data, sheet, { shared: ctx.shared, dynamicCm: ctx.dynamicCm, parseRuns: parseRichText });

  // Sheet properties.
  const sheetPr = child(root, 'sheetPr');
  const tab = readColor(child(sheetPr, 'tabColor'));
  if (tab) sheet.tabColor = tab;

  // Default sizes.
  const fmt = child(root, 'sheetFormatPr');
  const defH = numAttr(fmt, 'defaultRowHeight');
  if (defH) sheet.defaultRowHeightPt = defH;
  const defW = numAttr(fmt, 'defaultColWidth');
  const baseW = numAttr(fmt, 'baseColWidth');
  if (defW) sheet.defaultColWidthPx = widthCharsToPx(defW);
  else if (baseW !== undefined && baseW !== 8) sheet.defaultColWidthPx = widthCharsToPx(baseW + 5 / 7) + 3;

  // Columns.
  for (const col of childrenNamed(child(root, 'cols'), 'col')) {
    const min = (numAttr(col, 'min') ?? 1) - 1;
    const max = Math.min((numAttr(col, 'max') ?? min + 1) - 1, 16383);
    const width = numAttr(col, 'width');
    const info = {
      w: width,
      customWidth: boolAttr(col, 'customWidth') || undefined,
      hidden: boolAttr(col, 'hidden') || undefined,
      s: numAttr(col, 'style') || undefined,
      outline: numAttr(col, 'outlineLevel') || undefined,
      collapsed: boolAttr(col, 'collapsed') || undefined,
      bestFit: boolAttr(col, 'bestFit') || undefined,
    };
    for (let c = min; c <= max; c++) sheet.setColInfo(c, { ...info });
  }

  // View: zoom, gridlines, frozen panes, selection.
  const view = child(child(root, 'sheetViews'), 'sheetView');
  if (view) {
    const zoom = numAttr(view, 'zoomScale');
    if (zoom) sheet.zoom = zoom / 100;
    if (attr(view, 'showGridLines') === '0' || attr(view, 'showGridLines') === 'false') sheet.showGridLines = false;
    if (boolAttr(view, 'rightToLeft')) sheet.rightToLeft = true;
    const pane = child(view, 'pane');
    const state = attr(pane, 'state');
    if (pane && (state === 'frozen' || state === 'frozenSplit')) {
      sheet.freeze = { rows: Math.round(numAttr(pane, 'ySplit') ?? 0), cols: Math.round(numAttr(pane, 'xSplit') ?? 0) };
    }
    const sels = childrenNamed(view, 'selection');
    const sel = sels[sels.length - 1];
    const active = parseCellRef(attr(sel, 'activeCell') ?? '');
    if (active) sheet.selection = { active: { r: active.r, c: active.c }, ranges: parseSqref(attr(sel, 'sqref') ?? '') };
  }

  // Merged cells.
  sheet.setMerges(childrenNamed(child(root, 'mergeCells'), 'mergeCell').map((m) => parseRange(attr(m, 'ref') ?? '')).filter((r): r is NonNullable<typeof r> => !!r));

  // Hyperlinks.
  const rels = ctx.pkg.getRels(path);
  const links: Hyperlink[] = [];
  for (const h of childrenNamed(child(root, 'hyperlinks'), 'hyperlink')) {
    const ref = parseRange(attr(h, 'ref') ?? '');
    if (!ref) continue;
    const rid = attr(h, 'r:id');
    const rel = rid ? rels.find((r) => r.id === rid) : undefined;
    links.push({ ref, target: rel?.target, location: attr(h, 'location'), tooltip: attr(h, 'tooltip'), display: attr(h, 'display') });
  }
  sheet.setHyperlinks(links);

  // Data validations (with their messages) and conditional formats (kept as XML).
  sheet.validations = childrenNamed(child(root, 'dataValidations'), 'dataValidation').map((dv): DataValidation => {
    const errorStyle = attr(dv, 'errorStyle');
    return {
      sqref: parseSqref(attr(dv, 'sqref') ?? ''),
      type: attr(dv, 'type') ?? 'none',
      operator: attr(dv, 'operator'),
      formula1: child(dv, 'formula1') ? formulaFromFile(textContent(child(dv, 'formula1')!)) : undefined,
      formula2: child(dv, 'formula2') ? formulaFromFile(textContent(child(dv, 'formula2')!)) : undefined,
      allowBlank: boolAttr(dv, 'allowBlank'),
      hideDropDown: boolAttr(dv, 'showDropDown'),
      showErrorMessage: boolAttr(dv, 'showErrorMessage'),
      errorStyle: errorStyle === 'warning' || errorStyle === 'information' ? errorStyle : 'stop',
      errorTitle: attr(dv, 'errorTitle'),
      error: attr(dv, 'error'),
      showInputMessage: boolAttr(dv, 'showInputMessage'),
      promptTitle: attr(dv, 'promptTitle'),
      prompt: attr(dv, 'prompt'),
      raw: dv,
    };
  });
  sheet.condFormats = childrenNamed(root, 'conditionalFormatting').map((cf): ConditionalFormat => ({ sqref: parseSqref(attr(cf, 'sqref') ?? ''), raw: cf }));
  const af = child(root, 'autoFilter');
  const afRef = parseRange(attr(af, 'ref') ?? '');
  if (af && afRef) {
    sheet.autoFilter = { ref: afRef, columns: readFilterColumns(af, afRef.c0), raw: af };
    // Hidden rows inside a filtered area were hidden by the filter.
    if (sheet.autoFilter.columns.size) {
      for (let r = afRef.r0 + 1; r <= afRef.r1; r++) if (sheet.rows.get(r)?.hidden) sheet.filteredRows.add(r);
    }
  }

  // Drawings (pictures; other objects are kept).
  const drawingEl = child(root, 'drawing');
  const drawingRel = drawingEl ? rels.find((r) => r.id === attr(drawingEl, 'r:id')) : undefined;
  if (drawingRel) readDrawing(sheet, ctx.pkg.resolveTarget(path, drawingRel.target), ctx.pkg, ctx.wb.theme);

  // Comments.
  readComments(sheet, path, ctx);
}

const CUSTOM_OPS: Record<string, FilterOperator> = {
  equal: 'equal', notEqual: 'notEqual', greaterThan: 'greaterThan', lessThan: 'lessThan',
  greaterThanOrEqual: 'greaterThanOrEqual', lessThanOrEqual: 'lessThanOrEqual',
};

/** Reads `<filterColumn>` criteria (value lists and custom conditions with wildcards). */
export function readFilterColumns(af: XmlElement, firstCol: number): Map<number, ColumnFilter> {
  const out = new Map<number, ColumnFilter>();
  for (const fc of childrenNamed(af, 'filterColumn')) {
    const col = firstCol + (numAttr(fc, 'colId') ?? 0);
    const filters = child(fc, 'filters');
    const custom = child(fc, 'customFilters');
    const f: ColumnFilter = {};
    if (filters) {
      f.values = childrenNamed(filters, 'filter').map((x) => attr(x, 'val') ?? '');
      if (boolAttr(filters, 'blank')) f.blanks = true;
    }
    if (custom) {
      f.and = boolAttr(custom, 'and');
      f.conditions = childrenNamed(custom, 'customFilter').map((cf) => {
        const op = CUSTOM_OPS[attr(cf, 'operator') ?? 'equal'] ?? 'equal';
        const val = attr(cf, 'val') ?? '';
        const star = (s: string) => s.replace(/~\*/g, '\u0001');
        const v = star(val);
        const unstar = (s: string) => s.replace(/\u0001/g, '*');
        if ((op === 'equal' || op === 'notEqual') && v.length > 1 && v.startsWith('*') && v.endsWith('*')) return { op: op === 'equal' ? 'contains' : 'notContains', value: unstar(v.slice(1, -1)) } as FilterCondition;
        if (op === 'equal' && v.endsWith('*') && !v.startsWith('*')) return { op: 'beginsWith', value: unstar(v.slice(0, -1)) } as FilterCondition;
        if (op === 'equal' && v.startsWith('*') && !v.endsWith('*')) return { op: 'endsWith', value: unstar(v.slice(1)) } as FilterCondition;
        return { op, value: unstar(v) };
      });
    }
    if (f.values || f.conditions?.length) out.set(col, f);
  }
  return out;
}

function anchorPoint(e: XmlElement | undefined) {
  return {
    c: Number(textContent(child(e, 'xdr:col') ?? { type: 'text', text: '0' })) || 0,
    dx: Number(textContent(child(e, 'xdr:colOff') ?? { type: 'text', text: '0' })) || 0,
    r: Number(textContent(child(e, 'xdr:row') ?? { type: 'text', text: '0' })) || 0,
    dy: Number(textContent(child(e, 'xdr:rowOff') ?? { type: 'text', text: '0' })) || 0,
  };
}

let imageCounter = 0;

const ANCHORS = ['xdr:twoCellAnchor', 'xdr:oneCellAnchor', 'xdr:absoluteAnchor'];

/** Charts and preset shapes become editable objects; returns false for anything else. */
function readDrawingObject(sheet: Sheet, anchor: XmlElement, z: number, path: string, rels: Relationship[], pkg: OoxmlPackage, theme: Theme): boolean {
  if (!ANCHORS.includes(anchor.name)) return false;
  const ref = chartReference(anchor);
  const name = attr(findAll(anchor, 'xdr:cNvPr')[0], 'name');
  if (ref) {
    const rel = rels.find((r) => r.id === ref.rid);
    if (!rel || rel.targetMode === 'External') return false;
    const chartPath = pkg.resolveTarget(path, rel.target);
    const root = parse(pkg, chartPath)?.root;
    if (!root) return false;
    const chartex = ref.chartex || /chartex/i.test(rel.type);
    const def = chartex ? readChartexPart(root, theme, name ?? 'Chart') : readChartPart(root, theme, name ?? 'Chart');
    if (!def) return false;
    const { from, to } = anchorPoints(sheet, anchor);
    sheet.charts.push({ ...def, id: newDrawingId('chart'), z, from, to, source: { anchor, chartPath, chartex } });
    return true;
  }
  const content = anchorContent(anchor);
  const shape = content ? readShape(content, theme) : null;
  if (!shape) return false;
  const { from, to } = anchorPoints(sheet, anchor);
  sheet.shapes.push({ ...shape, id: newDrawingId('shape'), z, from, to, source: { anchor } });
  return true;
}

function readDrawing(sheet: Sheet, path: string, pkg: OoxmlPackage, theme: Theme): void {
  const doc = parse(pkg, path);
  if (!doc) return;
  const rels = pkg.getRels(path);
  elements(doc.root).forEach((anchor, z) => {
    const kind = anchor.name;
    const pic = child(anchor, 'xdr:pic');
    const blip = pic ? findAll(pic, 'a:blip')[0] : undefined;
    const embed = attr(blip, 'r:embed');
    const rel = embed ? rels.find((r) => r.id === embed) : undefined;
    const mediaPath = rel && rel.targetMode !== 'External' ? pkg.resolveTarget(path, rel.target) : undefined;
    const bytes = mediaPath ? pkg.get(mediaPath) : undefined;
    if (!pic || !bytes || !ANCHORS.includes(kind)) {
      if (!readDrawingObject(sheet, anchor, z, path, rels, pkg, theme)) sheet.otherAnchors.push(anchor);
      return;
    }
    const ext = mediaPath!.slice(mediaPath!.lastIndexOf('.') + 1).toLowerCase();
    const cNvPr = findAll(pic, 'xdr:cNvPr')[0];
    const img: SheetImage = {
      id: `img${++imageCounter}`,
      name: attr(cNvPr, 'name') ?? `Picture ${imageCounter}`,
      descr: attr(cNvPr, 'descr'),
      mime: IMAGE_CONTENT_TYPES[ext] ?? 'image/png',
      data: bytes,
      anchorType: kind === 'xdr:twoCellAnchor' ? 'twoCell' : kind === 'xdr:oneCellAnchor' ? 'oneCell' : 'absolute',
      from: anchorPoint(child(anchor, 'xdr:from')),
      raw: anchor,
      z,
    };
    if (kind === 'xdr:twoCellAnchor') {
      img.to = anchorPoint(child(anchor, 'xdr:to'));
      const editAs = attr(anchor, 'editAs');
      if (editAs === 'oneCell' || editAs === 'absolute' || editAs === 'twoCell') img.editAs = editAs;
    } else {
      const extEl = child(anchor, 'xdr:ext');
      img.ext = { cx: numAttr(extEl, 'cx') ?? 0, cy: numAttr(extEl, 'cy') ?? 0 };
      if (kind === 'xdr:absoluteAnchor') {
        const pos = child(anchor, 'xdr:pos');
        img.pos = { x: numAttr(pos, 'x') ?? 0, y: numAttr(pos, 'y') ?? 0 };
      }
    }
    sheet.images.push(img);
  });
}

function readComments(sheet: Sheet, path: string, ctx: SheetContext): void {
  const rels = ctx.pkg.getRels(path);
  const threads: CommentThread[] = [];
  const threadedRel = rels.find((r) => relTypeIs(r.type, REL.threadedComment));
  if (threadedRel) {
    const doc = parse(ctx.pkg, ctx.pkg.resolveTarget(path, threadedRel.target));
    const byId = new Map<string, CommentThread>();
    for (const tc of childrenNamed(doc?.root, 'tc:threadedComment')) {
      const ref = parseCellRef(attr(tc, 'ref') ?? '');
      if (!ref) continue;
      const personId = attr(tc, 'personId');
      const person = ctx.persons.find((p) => p.id === personId);
      const entry = {
        id: attr(tc, 'id') ?? `c${threads.length}`,
        author: person?.displayName ?? 'Unknown',
        personId,
        text: textContent(child(tc, 'tc:text') ?? tc),
        created: attr(tc, 'dT'),
      };
      const parentId = attr(tc, 'parentId');
      const parent = parentId ? byId.get(parentId) : undefined;
      if (parent) {
        parent.comments.push(entry);
        if (boolAttr(tc, 'done')) parent.resolved = true;
        continue;
      }
      const thread: CommentThread = { id: entry.id, r: ref.r, c: ref.c, kind: 'threaded', comments: [entry], resolved: boolAttr(tc, 'done') };
      byId.set(entry.id, thread);
      threads.push(thread);
    }
  }
  const legacyRel = rels.find((r) => relTypeIs(r.type, REL.comments));
  if (legacyRel) {
    const doc = parse(ctx.pkg, ctx.pkg.resolveTarget(path, legacyRel.target));
    const authors = childrenNamed(child(doc?.root, 'authors'), 'author').map((a) => textContent(a));
    const vml = readVmlShapes(sheet, path, ctx.pkg);
    for (const c of childrenNamed(child(doc?.root, 'commentList'), 'comment')) {
      const ref = parseCellRef(attr(c, 'ref') ?? '');
      if (!ref) continue;
      if (threads.some((t) => t.r === ref.r && t.c === ref.c)) continue;
      const author = authors[numAttr(c, 'authorId') ?? 0] ?? '';
      let text = textContent(child(c, 'text') ?? c);
      if (author && text.startsWith(`${author}:\n`)) text = text.slice(author.length + 2);
      threads.push({
        id: attr(c, 'xr:uid') ?? `note-${ref.r}-${ref.c}`,
        r: ref.r,
        c: ref.c,
        kind: 'note',
        comments: [{ id: attr(c, 'xr:uid') ?? `note-${ref.r}-${ref.c}`, author, text }],
        resolved: false,
        vmlShape: vml.get(`${ref.r},${ref.c}`),
      });
    }
  }
  sheet.comments = threads;
}

function readVmlShapes(sheet: Sheet, path: string, pkg: OoxmlPackage): Map<string, string> {
  const out = new Map<string, string>();
  const root = sheet.sourceXml;
  const legacy = child(root, 'legacyDrawing');
  const rel = legacy ? pkg.getRels(path).find((r) => r.id === attr(legacy, 'r:id')) : undefined;
  if (!rel) return out;
  const text = pkg.getText(pkg.resolveTarget(path, rel.target));
  if (!text) return out;
  const re = /<v:shape\b[\s\S]*?<\/v:shape>/g;
  let m: RegExpExecArray | null;
  while ((m = re.exec(text))) {
    const row = /<x:Row>(\d+)<\/x:Row>/.exec(m[0]);
    const col = /<x:Column>(\d+)<\/x:Column>/.exec(m[0]);
    if (row && col) out.set(`${row[1]},${col[1]}`, m[0]);
  }
  return out;
}

// -----------------------------------------------------------------------------------------------
// Workbook

function dynamicArrayMetadata(pkg: OoxmlPackage, wbPath: string): Set<number> {
  const out = new Set<number>();
  const rel = pkg.findRel(wbPath, REL.sheetMetadata);
  if (!rel) return out;
  const doc = parse(pkg, pkg.resolveTarget(wbPath, rel.target));
  if (!doc) return out;
  const types = childrenNamed(child(doc.root, 'metadataTypes'), 'metadataType').map((t) => attr(t, 'name'));
  const blocks = childrenNamed(child(doc.root, 'cellMetadata'), 'bk');
  blocks.forEach((bk, i) => {
    const rc = child(bk, 'rc');
    const t = numAttr(rc, 't');
    if (t !== undefined && types[t - 1] === 'XLDAPR') out.add(i + 1);
  });
  return out;
}

export function readXlsx(bytes: Uint8Array, opts: ImportOptions = {}): Workbook {
  const pkg = OoxmlPackage.fromBytes(bytes);
  const wbPath = pkg.workbookPath();
  const wbDoc = parse(pkg, wbPath);
  if (!wbDoc) throw new Error('The workbook part is missing.');
  const wbRoot = wbDoc.root;

  const stylesRel = pkg.findRel(wbPath, REL.styles);
  const stylesDoc = stylesRel ? parse(pkg, pkg.resolveTarget(wbPath, stylesRel.target)) : undefined;
  const parsedStyles = readStyles(stylesDoc?.root);

  const wb = new Workbook(parsedStyles.styles[0]);
  for (const s of parsedStyles.styles.slice(1)) wb.styles.push(s);
  wb.dxfs = childrenNamed(child(stylesDoc?.root, 'dxfs'), 'dxf');
  if (opts.fileName) wb.props.fileName = opts.fileName;

  const themeRel = pkg.findRel(wbPath, REL.theme);
  wb.theme = readTheme(themeRel ? parse(pkg, pkg.resolveTarget(wbPath, themeRel.target))?.root : undefined);
  if (parsedStyles.indexedColors) wb.theme.indexed = parsedStyles.indexedColors;

  const wbPr = child(wbRoot, 'workbookPr');
  wb.date1904 = boolAttr(wbPr, 'date1904');

  const sstRel = pkg.findRel(wbPath, REL.sharedStrings);
  const shared = readSharedStrings(sstRel ? pkg.getText(pkg.resolveTarget(wbPath, sstRel.target)) : undefined);

  const personRel = pkg.findRel(wbPath, REL.person);
  const persons: Person[] = [];
  if (personRel) {
    const doc = parse(pkg, pkg.resolveTarget(wbPath, personRel.target));
    for (const p of childrenNamed(doc?.root, 'tc:person')) {
      persons.push({ id: attr(p, 'id') ?? '', displayName: attr(p, 'displayName') ?? '', userId: attr(p, 'userId'), providerId: attr(p, 'providerId') });
    }
  }
  wb.persons = persons;

  const dynamicCm = dynamicArrayMetadata(pkg, wbPath);
  const ctx: SheetContext = { pkg, shared, dynamicCm, persons, wb };
  const wbRels = pkg.getRels(wbPath);
  for (const s of childrenNamed(child(wbRoot, 'sheets'), 'sheet')) {
    const name = attr(s, 'name') ?? `Sheet${wb.sheets.length + 1}`;
    const sheet = new Sheet(name, numAttr(s, 'sheetId') ?? wb.nextSheetId());
    const state = attr(s, 'state');
    if (state === 'hidden' || state === 'veryHidden') sheet.state = state;
    wb.sheets.push(sheet);
    const rel = wbRels.find((r) => r.id === attr(s, 'r:id'));
    if (!rel) continue;
    const path = pkg.resolveTarget(wbPath, rel.target);
    if (relTypeIs(rel.type, REL.chartsheet)) {
      sheet.sourcePath = path;
      sheet.chartsheet = true;
      continue;
    }
    readWorksheet(sheet, path, ctx);
  }
  if (!wb.sheets.length) wb.addSheet('Sheet1');

  const names: DefinedName[] = [];
  for (const dn of childrenNamed(child(wbRoot, 'definedNames'), 'definedName')) {
    const extra: Record<string, string> = {};
    for (const [k, v] of Object.entries(dn.attrs)) if (!['name', 'localSheetId', 'hidden', 'comment'].includes(k)) extra[k] = v;
    names.push({
      name: attr(dn, 'name') ?? '',
      formula: formulaFromFile(textContent(dn)),
      localSheet: numAttr(dn, 'localSheetId'),
      hidden: boolAttr(dn, 'hidden') || undefined,
      comment: attr(dn, 'comment'),
      extra: Object.keys(extra).length ? extra : undefined,
    });
  }
  wb.names = names;

  const view = child(child(wbRoot, 'bookViews'), 'workbookView');
  const active = numAttr(view, 'activeTab') ?? 0;
  wb.activeSheetIndex = Math.min(Math.max(0, active), wb.sheets.length - 1);
  if (wb.sheets[wb.activeSheetIndex]?.state !== 'visible') {
    const firstVisible = wb.sheets.findIndex((s) => s.state === 'visible');
    wb.activeSheetIndex = Math.max(0, firstVisible);
  }

  // Document properties.
  const coreRel = pkg.findRel('', REL.coreProps);
  const core = coreRel ? parse(pkg, pkg.resolveTarget('', coreRel.target)) : undefined;
  if (core) {
    for (const e of elements(core.root)) {
      const local = e.name.replace(/^.*:/, '');
      const v = textContent(e);
      if (v && ['title', 'subject', 'creator', 'keywords', 'description', 'lastModifiedBy', 'created', 'modified', 'category'].includes(local)) wb.props[local] = v;
    }
  }

  wb.source = {
    pkg,
    workbookPath: wbPath,
    workbookXml: wbRoot,
    stylesXml: stylesDoc?.root,
    extraNamespaces: wbDoc.extraNamespaces,
  };
  wb.chartRefs = readPackageChartRefs(pkg);
  void CT;
  void strFromU8;
  return wb;
}
