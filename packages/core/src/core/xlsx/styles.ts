/** styles.xml: reading cell formats into CellStyle objects and writing them back. */
import type { ColorSpec } from '../model/color';
import {
  BUILTIN_NUMFMTS, DEFAULT_FONT, borderKey, fillKey, fontKey, type AlignmentStyle, type BorderLine, type BorderStyle,
  type CellStyle, type FillStyle, type FontStyle, type StyleTable, type UnderlineType,
} from '../model/style';
import {
  attr, boolAttr, child, childVal, childrenNamed, el, elements, numAttr, parseFragment, serializeNode, type XmlElement,
} from '../xml/xml';

export function readColor(e: XmlElement | undefined): ColorSpec | undefined {
  if (!e) return undefined;
  const tint = numAttr(e, 'tint');
  const withTint = <T extends object>(c: T): T => (tint ? { ...c, tint } : c);
  if (attr(e, 'auto') === '1' || attr(e, 'auto') === 'true') return { auto: true };
  const rgb = attr(e, 'rgb');
  if (rgb) return withTint({ rgb: rgb.toUpperCase() });
  const theme = numAttr(e, 'theme');
  if (theme !== undefined) return withTint({ theme });
  const indexed = numAttr(e, 'indexed');
  if (indexed !== undefined) return withTint({ indexed });
  return undefined;
}

export function colorAttrs(c: ColorSpec): Record<string, string | number | undefined> {
  if ('auto' in c) return { auto: 1 };
  const tint = 'tint' in c && c.tint ? c.tint : undefined;
  if ('rgb' in c) return { rgb: c.rgb.length === 6 ? `FF${c.rgb}` : c.rgb, tint };
  if ('theme' in c) return { theme: c.theme, tint };
  return { indexed: c.indexed, tint };
}

export function readFont(e: XmlElement): FontStyle {
  const font: FontStyle = { name: childVal(e, 'name') ?? childVal(e, 'rFont') ?? DEFAULT_FONT.name, size: Number(childVal(e, 'sz') ?? DEFAULT_FONT.size) };
  const flag = (name: string) => {
    const c = child(e, name);
    if (!c) return false;
    const v = attr(c, 'val');
    return v === undefined || v === '1' || v === 'true';
  };
  if (flag('b')) font.bold = true;
  if (flag('i')) font.italic = true;
  if (flag('strike')) font.strike = true;
  const u = child(e, 'u');
  if (u) {
    const v = attr(u, 'val') ?? 'single';
    if (v !== 'none') font.underline = v as UnderlineType;
  }
  const color = readColor(child(e, 'color'));
  if (color) font.color = color;
  const va = childVal(e, 'vertAlign');
  if (va === 'superscript' || va === 'subscript') font.vertAlign = va;
  const family = childVal(e, 'family');
  if (family) font.family = Number(family);
  const scheme = childVal(e, 'scheme');
  if (scheme === 'major' || scheme === 'minor') font.scheme = scheme;
  const charset = childVal(e, 'charset');
  if (charset) font.charset = Number(charset);
  return font;
}

function readFill(e: XmlElement): FillStyle | undefined {
  const pattern = child(e, 'patternFill');
  if (pattern) {
    const type = attr(pattern, 'patternType') ?? (child(pattern, 'fgColor') ? 'solid' : 'none');
    if (type === 'none') return undefined;
    return { pattern: type, fg: readColor(child(pattern, 'fgColor')), bg: readColor(child(pattern, 'bgColor')) };
  }
  const gradient = child(e, 'gradientFill');
  if (gradient) {
    const stop = child(gradient, 'stop');
    return { gradientXml: serializeNode(gradient), color: readColor(child(stop, 'color')) };
  }
  return undefined;
}

function readBorder(e: XmlElement): BorderStyle | undefined {
  const b: BorderStyle = {};
  for (const side of ['left', 'right', 'top', 'bottom', 'diagonal'] as const) {
    const s = child(e, side) ?? (side === 'left' ? child(e, 'start') : side === 'right' ? child(e, 'end') : undefined);
    const style = attr(s, 'style');
    if (s && style && style !== 'none') b[side] = { style: style as BorderLine['style'], color: readColor(child(s, 'color')) };
  }
  if (boolAttr(e, 'diagonalUp')) b.diagonalUp = true;
  if (boolAttr(e, 'diagonalDown')) b.diagonalDown = true;
  return Object.keys(b).length ? b : undefined;
}

function readAlignment(e: XmlElement | undefined): AlignmentStyle | undefined {
  if (!e) return undefined;
  const a: AlignmentStyle = {};
  const h = attr(e, 'horizontal');
  if (h) a.horizontal = h as AlignmentStyle['horizontal'];
  const v = attr(e, 'vertical');
  if (v) a.vertical = v as AlignmentStyle['vertical'];
  if (boolAttr(e, 'wrapText')) a.wrap = true;
  if (boolAttr(e, 'shrinkToFit')) a.shrink = true;
  const indent = numAttr(e, 'indent');
  if (indent) a.indent = indent;
  const rot = numAttr(e, 'textRotation');
  if (rot) a.rotation = rot;
  const ro = numAttr(e, 'readingOrder');
  if (ro) a.readingOrder = ro;
  if (attr(e, 'xe:clip') === '1') a.clip = true;
  return Object.keys(a).length ? a : undefined;
}

export interface ParsedStyles {
  /** One style per cellXfs entry (index = cell s attribute). */
  styles: CellStyle[];
  /** Custom indexed color palette, if the file defines one. */
  indexedColors?: string[];
  root: XmlElement;
}

export function readStyles(root: XmlElement | undefined): ParsedStyles {
  if (!root) return { styles: [{ font: { ...DEFAULT_FONT }, numFmt: 'General' }], root: el('styleSheet') };
  const numFmts = new Map<number, string>();
  for (const nf of childrenNamed(child(root, 'numFmts'), 'numFmt')) {
    const id = numAttr(nf, 'numFmtId');
    const code = attr(nf, 'formatCode');
    if (id !== undefined && code !== undefined) numFmts.set(id, code);
  }
  const fonts = childrenNamed(child(root, 'fonts'), 'font').map(readFont);
  const fills = childrenNamed(child(root, 'fills'), 'fill').map(readFill);
  const borders = childrenNamed(child(root, 'borders'), 'border').map(readBorder);
  const styles: CellStyle[] = [];
  for (const xf of childrenNamed(child(root, 'cellXfs'), 'xf')) {
    const numFmtId = numAttr(xf, 'numFmtId') ?? 0;
    const style: CellStyle = {
      font: { ...(fonts[numAttr(xf, 'fontId') ?? 0] ?? DEFAULT_FONT) },
      fill: fills[numAttr(xf, 'fillId') ?? 0],
      border: borders[numAttr(xf, 'borderId') ?? 0],
      numFmt: numFmts.get(numFmtId) ?? BUILTIN_NUMFMTS[numFmtId] ?? 'General',
      align: readAlignment(child(xf, 'alignment')),
    };
    const prot = child(xf, 'protection');
    if (prot) {
      if (attr(prot, 'locked') === '0' || attr(prot, 'locked') === 'false') style.locked = false;
      if (boolAttr(prot, 'hidden')) style.hidden = true;
    }
    if (boolAttr(xf, 'quotePrefix')) style.quotePrefix = true;
    const xfId = numAttr(xf, 'xfId');
    if (xfId) style.xfId = xfId;
    styles.push(style);
  }
  if (!styles.length) styles.push({ font: { ...(fonts[0] ?? DEFAULT_FONT) }, numFmt: 'General' });
  const indexed = childrenNamed(child(child(root, 'colors'), 'indexedColors'), 'rgbColor').map((c) => (attr(c, 'rgb') ?? 'FF000000').slice(-6).toUpperCase());
  return { styles, indexedColors: indexed.length ? indexed : undefined, root };
}

// -----------------------------------------------------------------------------------------------
// Writing

function fontElement(f: FontStyle): XmlElement {
  const kids: XmlElement[] = [];
  if (f.bold) kids.push(el('b'));
  if (f.italic) kids.push(el('i'));
  if (f.strike) kids.push(el('strike'));
  if (f.underline) kids.push(el('u', f.underline === 'single' ? {} : { val: f.underline }));
  if (f.vertAlign) kids.push(el('vertAlign', { val: f.vertAlign }));
  kids.push(el('sz', { val: f.size }));
  if (f.color) kids.push(el('color', colorAttrs(f.color)));
  kids.push(el('name', { val: f.name }));
  if (f.family !== undefined) kids.push(el('family', { val: f.family }));
  if (f.charset !== undefined) kids.push(el('charset', { val: f.charset }));
  if (f.scheme) kids.push(el('scheme', { val: f.scheme }));
  return el('font', {}, kids);
}

function fillElement(f: FillStyle | undefined): XmlElement {
  if (!f) return el('fill', {}, [el('patternFill', { patternType: 'none' })]);
  if ('gradientXml' in f) return el('fill', {}, [parseFragment(f.gradientXml)]);
  const kids: XmlElement[] = [];
  if (f.fg) kids.push(el('fgColor', colorAttrs(f.fg)));
  if (f.bg) kids.push(el('bgColor', colorAttrs(f.bg)));
  else if (f.pattern === 'solid') kids.push(el('bgColor', { indexed: 64 }));
  return el('fill', {}, [el('patternFill', { patternType: f.pattern }, kids)]);
}

function borderElement(b: BorderStyle | undefined): XmlElement {
  const side = (name: string, line: BorderLine | undefined) =>
    line ? el(name, { style: line.style }, [line.color ? el('color', colorAttrs(line.color)) : el('color', { auto: 1 })]) : el(name);
  return el('border', { diagonalUp: b?.diagonalUp ? 1 : undefined, diagonalDown: b?.diagonalDown ? 1 : undefined }, [
    side('left', b?.left),
    side('right', b?.right),
    side('top', b?.top),
    side('bottom', b?.bottom),
    side('diagonal', b?.diagonal),
  ]);
}

function alignmentElement(a: AlignmentStyle | undefined): XmlElement | null {
  if (!a) return null;
  const attrs: Record<string, string | number | undefined> = {
    horizontal: a.horizontal,
    vertical: a.vertical,
    textRotation: a.rotation,
    wrapText: a.wrap ? 1 : undefined,
    indent: a.indent,
    shrinkToFit: a.shrink ? 1 : undefined,
    readingOrder: a.readingOrder,
    'xe:clip': a.clip ? 1 : undefined,
  };
  if (Object.values(attrs).every((v) => v === undefined)) return null;
  return el('alignment', attrs);
}

const BUILTIN_IDS = new Map<string, number>(Object.entries(BUILTIN_NUMFMTS).map(([id, code]) => [code, Number(id)]));

export interface StylesOutput {
  root: XmlElement;
  /** True when the extension namespace (xe) is used (clip alignment). */
  usesExtension: boolean;
}

/** Builds styles.xml for a style table, keeping cell styles, dxfs and other parts of the source. */
export function writeStyles(table: StyleTable, source: XmlElement | undefined, dxfs?: XmlElement[]): StylesOutput {
  const numFmts = new Map<string, number>();
  const sourceNumFmts = new Map<string, number>();
  for (const nf of childrenNamed(child(source, 'numFmts'), 'numFmt')) {
    const id = numAttr(nf, 'numFmtId');
    const code = attr(nf, 'formatCode');
    if (id !== undefined && code !== undefined && !sourceNumFmts.has(code)) sourceNumFmts.set(code, id);
  }
  let nextNumFmt = Math.max(163, ...sourceNumFmts.values()) + 1;
  const fontList: XmlElement[] = [];
  const fontIndex = new Map<string, number>();
  const fillList: XmlElement[] = [fillElement(undefined), el('fill', {}, [el('patternFill', { patternType: 'gray125' })])];
  const fillIndex = new Map<string, number>([[fillKey(undefined), 0], [fillKey({ pattern: 'gray125' }), 1]]);
  const borderList: XmlElement[] = [borderElement(undefined)];
  const borderIndex = new Map<string, number>([[borderKey(undefined), 0]]);
  let usesExtension = false;

  const fontId = (f: FontStyle) => {
    const k = fontKey(f);
    let i = fontIndex.get(k);
    if (i === undefined) {
      i = fontList.length;
      fontList.push(fontElement(f));
      fontIndex.set(k, i);
    }
    return i;
  };
  const fillId = (f: FillStyle | undefined) => {
    const k = fillKey(f);
    let i = fillIndex.get(k);
    if (i === undefined) {
      i = fillList.length;
      fillList.push(fillElement(f));
      fillIndex.set(k, i);
    }
    return i;
  };
  const borderId = (b: BorderStyle | undefined) => {
    const k = borderKey(b);
    let i = borderIndex.get(k);
    if (i === undefined) {
      i = borderList.length;
      borderList.push(borderElement(b));
      borderIndex.set(k, i);
    }
    return i;
  };
  const numFmtId = (code: string) => {
    if (!code || code === 'General') return 0;
    const builtin = BUILTIN_IDS.get(code);
    if (builtin !== undefined) return builtin;
    let id = numFmts.get(code) ?? sourceNumFmts.get(code);
    if (id === undefined) id = nextNumFmt++;
    numFmts.set(code, id);
    return id;
  };

  const all = table.all();
  fontId(all[0].font);
  const xfs: XmlElement[] = all.map((s) => {
    const align = alignmentElement(s.align);
    if (s.align?.clip) usesExtension = true;
    const prot = s.locked === false || s.hidden ? el('protection', { locked: s.locked === false ? 0 : undefined, hidden: s.hidden ? 1 : undefined }) : null;
    const nf = numFmtId(s.numFmt);
    const fo = fontId(s.font);
    const fi = fillId(s.fill);
    const bo = borderId(s.border);
    return el(
      'xf',
      {
        numFmtId: nf,
        fontId: fo,
        fillId: fi,
        borderId: bo,
        xfId: s.xfId ?? 0,
        quotePrefix: s.quotePrefix ? 1 : undefined,
        applyNumberFormat: nf ? 1 : undefined,
        applyFont: fo ? 1 : undefined,
        applyFill: fi ? 1 : undefined,
        applyBorder: bo ? 1 : undefined,
        applyAlignment: align ? 1 : undefined,
        applyProtection: prot ? 1 : undefined,
      },
      [align, prot],
    );
  });

  const numFmtEls = [...numFmts.entries()].filter(([, id]) => id >= 164).sort((a, b) => a[1] - b[1]).map(([code, id]) => el('numFmt', { numFmtId: id, formatCode: code }));
  const kept = (name: string) => child(source, name);
  const cellStyleXfs = kept('cellStyleXfs') ?? el('cellStyleXfs', { count: 1 }, [el('xf', { numFmtId: 0, fontId: 0, fillId: 0, borderId: 0 })]);
  // Font, fill and border ids referenced by cellStyleXfs must still exist.
  for (const xf of elements(cellStyleXfs)) {
    const fo = numAttr(xf, 'fontId') ?? 0;
    const fi = numAttr(xf, 'fillId') ?? 0;
    const bo = numAttr(xf, 'borderId') ?? 0;
    if (fo >= fontList.length) xf.attrs.fontId = '0';
    if (fi >= fillList.length) xf.attrs.fillId = '0';
    if (bo >= borderList.length) xf.attrs.borderId = '0';
  }
  const children: (XmlElement | null | undefined)[] = [
    numFmtEls.length ? el('numFmts', { count: numFmtEls.length }, numFmtEls) : null,
    el('fonts', { count: fontList.length }, fontList),
    el('fills', { count: fillList.length }, fillList),
    el('borders', { count: borderList.length }, borderList),
    cellStyleXfs,
    el('cellXfs', { count: xfs.length }, xfs),
    kept('cellStyles') ?? el('cellStyles', { count: 1 }, [el('cellStyle', { name: 'Normal', xfId: 0, builtinId: 0 })]),
    dxfs ? el('dxfs', { count: dxfs.length }, dxfs) : kept('dxfs') ?? el('dxfs', { count: 0 }),
    kept('tableStyles') ?? el('tableStyles', { count: 0, defaultTableStyle: 'TableStyleMedium2', defaultPivotStyle: 'PivotStyleLight16' }),
    kept('colors'),
    kept('extLst'),
  ];
  const attrs: Record<string, string> = {};
  for (const [k, v] of Object.entries(source?.attrs ?? {})) attrs[k] = v;
  const root = el('styleSheet', attrs, children);
  return { root, usesExtension };
}
