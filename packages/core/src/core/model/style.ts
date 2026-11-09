import type { ColorSpec } from './color';

export type UnderlineType = 'single' | 'double' | 'singleAccounting' | 'doubleAccounting';

export interface FontStyle {
  name: string;
  /** Points. */
  size: number;
  bold?: boolean;
  italic?: boolean;
  underline?: UnderlineType;
  strike?: boolean;
  color?: ColorSpec;
  vertAlign?: 'superscript' | 'subscript';
  family?: number;
  scheme?: 'major' | 'minor';
  charset?: number;
}

export const BORDER_LINE_STYLES = [
  'thin', 'medium', 'thick', 'dashed', 'dotted', 'double', 'hair', 'mediumDashed', 'dashDot', 'mediumDashDot',
  'dashDotDot', 'mediumDashDotDot', 'slantDashDot',
] as const;

export type BorderLineStyle = (typeof BORDER_LINE_STYLES)[number];

export interface BorderLine {
  style: BorderLineStyle;
  color?: ColorSpec;
}

export type BorderSide = 'left' | 'right' | 'top' | 'bottom';

export interface BorderStyle {
  left?: BorderLine;
  right?: BorderLine;
  top?: BorderLine;
  bottom?: BorderLine;
  diagonal?: BorderLine;
  diagonalUp?: boolean;
  diagonalDown?: boolean;
}

export interface PatternFill {
  pattern: string;
  fg?: ColorSpec;
  bg?: ColorSpec;
}

export interface GradientFill {
  /** Serialized `<gradientFill>` element, kept for round trips. */
  gradientXml: string;
  /** First stop color, used to paint the cell. */
  color?: ColorSpec;
}

export type FillStyle = PatternFill | GradientFill;

export type HAlign = 'general' | 'left' | 'center' | 'right' | 'fill' | 'justify' | 'centerContinuous' | 'distributed';
export type VAlign = 'top' | 'center' | 'bottom' | 'justify' | 'distributed';

export interface AlignmentStyle {
  horizontal?: HAlign;
  vertical?: VAlign;
  wrap?: boolean;
  /** Clip text at the cell edge instead of overflowing (an editor extension, see export). */
  clip?: boolean;
  shrink?: boolean;
  indent?: number;
  /** 0..180 degrees, 255 = stacked vertical text. */
  rotation?: number;
  readingOrder?: number;
}

export interface CellStyle {
  font: FontStyle;
  fill?: FillStyle;
  border?: BorderStyle;
  numFmt: string;
  align?: AlignmentStyle;
  /** Protection: undefined means locked (the Excel default). */
  locked?: boolean;
  hidden?: boolean;
  quotePrefix?: boolean;
  /** Index of the parent cell style (cellStyleXfs) in the source file. */
  xfId?: number;
}

export const DEFAULT_FONT: FontStyle = { name: 'Calibri', size: 11, color: { theme: 1 }, family: 2, scheme: 'minor' };

export function defaultStyle(font: FontStyle = DEFAULT_FONT): CellStyle {
  return { font: { ...font }, numFmt: 'General' };
}

function colorKey(c: ColorSpec | undefined): string {
  if (!c) return '';
  if ('rgb' in c) return `r${c.rgb.toUpperCase()}${c.tint ? `~${c.tint}` : ''}`;
  if ('theme' in c) return `t${c.theme}${c.tint ? `~${c.tint}` : ''}`;
  if ('indexed' in c) return `i${c.indexed}${c.tint ? `~${c.tint}` : ''}`;
  return 'a';
}

function lineKey(l: BorderLine | undefined): string {
  return l ? `${l.style}/${colorKey(l.color)}` : '';
}

export function fontKey(f: FontStyle): string {
  return [
    f.name, f.size, f.bold ? 1 : 0, f.italic ? 1 : 0, f.underline ?? '', f.strike ? 1 : 0, colorKey(f.color),
    f.vertAlign ?? '', f.family ?? '', f.scheme ?? '', f.charset ?? '',
  ].join('|');
}

export function fillKey(f: FillStyle | undefined): string {
  if (!f) return '';
  if ('gradientXml' in f) return `g${f.gradientXml}`;
  return `p${f.pattern}/${colorKey(f.fg)}/${colorKey(f.bg)}`;
}

export function borderKey(b: BorderStyle | undefined): string {
  if (!b) return '';
  return [lineKey(b.left), lineKey(b.right), lineKey(b.top), lineKey(b.bottom), lineKey(b.diagonal), b.diagonalUp ? 1 : 0, b.diagonalDown ? 1 : 0].join(',');
}

export function alignKey(a: AlignmentStyle | undefined): string {
  if (!a) return '';
  return [a.horizontal ?? '', a.vertical ?? '', a.wrap ? 1 : 0, a.clip ? 1 : 0, a.shrink ? 1 : 0, a.indent ?? '', a.rotation ?? '', a.readingOrder ?? ''].join(',');
}

export function styleKey(s: CellStyle): string {
  return [
    fontKey(s.font), fillKey(s.fill), borderKey(s.border), s.numFmt, alignKey(s.align),
    s.locked === false ? 'u' : '', s.hidden ? 'h' : '', s.quotePrefix ? 'q' : '', s.xfId ?? '',
  ].join('#');
}

/** Removes empty sub-objects so equal styles produce equal keys. */
export function normalizeStyle(s: CellStyle): CellStyle {
  const out: CellStyle = { font: { ...s.font }, numFmt: s.numFmt || 'General' };
  for (const k of Object.keys(out.font) as (keyof FontStyle)[]) {
    const v = out.font[k];
    if (v === undefined || v === false) delete out.font[k];
  }
  if (s.fill && ('gradientXml' in s.fill || s.fill.pattern !== 'none')) out.fill = s.fill;
  if (s.border) {
    const b: BorderStyle = {};
    for (const side of ['left', 'right', 'top', 'bottom', 'diagonal'] as const) if (s.border[side]) b[side] = s.border[side];
    if (s.border.diagonalUp) b.diagonalUp = true;
    if (s.border.diagonalDown) b.diagonalDown = true;
    if (Object.keys(b).length) out.border = b;
  }
  if (s.align) {
    const a: AlignmentStyle = {};
    if (s.align.horizontal && s.align.horizontal !== 'general') a.horizontal = s.align.horizontal;
    if (s.align.vertical && s.align.vertical !== 'bottom') a.vertical = s.align.vertical;
    if (s.align.wrap) a.wrap = true;
    if (s.align.clip && !s.align.wrap) a.clip = true;
    if (s.align.shrink) a.shrink = true;
    if (s.align.indent) a.indent = s.align.indent;
    if (s.align.rotation) a.rotation = s.align.rotation;
    if (s.align.readingOrder) a.readingOrder = s.align.readingOrder;
    if (Object.keys(a).length) out.align = a;
  }
  if (s.locked === false) out.locked = false;
  if (s.hidden) out.hidden = true;
  if (s.quotePrefix) out.quotePrefix = true;
  if (s.xfId !== undefined && s.xfId !== 0) out.xfId = s.xfId;
  return out;
}

/** A partial update applied to a style by formatting commands. `null` clears a property. */
export interface StylePatch {
  font?: Partial<{ [K in keyof FontStyle]: FontStyle[K] | null }>;
  fill?: FillStyle | null;
  border?: Partial<{ [K in keyof BorderStyle]: BorderStyle[K] | null }> | null;
  numFmt?: string;
  align?: Partial<{ [K in keyof AlignmentStyle]: AlignmentStyle[K] | null }>;
  locked?: boolean;
  quotePrefix?: boolean | null;
}

export function applyPatch(base: CellStyle, patch: StylePatch): CellStyle {
  const s: CellStyle = {
    font: { ...base.font },
    fill: base.fill,
    border: base.border ? { ...base.border } : undefined,
    numFmt: base.numFmt,
    align: base.align ? { ...base.align } : undefined,
    locked: base.locked,
    hidden: base.hidden,
    quotePrefix: base.quotePrefix,
    xfId: base.xfId,
  };
  if (patch.font) {
    for (const [k, v] of Object.entries(patch.font)) {
      if (v === null || v === undefined) delete (s.font as unknown as Record<string, unknown>)[k];
      else (s.font as unknown as Record<string, unknown>)[k] = v;
    }
    if (!s.font.name) s.font.name = base.font.name;
    if (!s.font.size) s.font.size = base.font.size;
  }
  if (patch.fill !== undefined) s.fill = patch.fill ?? undefined;
  if (patch.border !== undefined) {
    if (patch.border === null) s.border = undefined;
    else {
      const b: BorderStyle = { ...(s.border ?? {}) };
      for (const [k, v] of Object.entries(patch.border)) {
        if (v === null || v === undefined) delete (b as Record<string, unknown>)[k];
        else (b as Record<string, unknown>)[k] = v;
      }
      s.border = b;
    }
  }
  if (patch.numFmt !== undefined) s.numFmt = patch.numFmt;
  if (patch.align) {
    const a: AlignmentStyle = { ...(s.align ?? {}) };
    for (const [k, v] of Object.entries(patch.align)) {
      if (v === null || v === undefined) delete (a as Record<string, unknown>)[k];
      else (a as Record<string, unknown>)[k] = v;
    }
    if (patch.align.wrap) delete a.clip;
    if (patch.align.clip) delete a.wrap;
    s.align = a;
  }
  if (patch.locked !== undefined) s.locked = patch.locked;
  if (patch.quotePrefix !== undefined) s.quotePrefix = patch.quotePrefix ?? undefined;
  return normalizeStyle(s);
}

/** Interned cell styles. Cells store an index into this table; index 0 is the workbook default. */
export class StyleTable {
  private readonly list: CellStyle[] = [];
  private readonly index = new Map<string, number>();

  constructor(base: CellStyle = defaultStyle()) {
    this.intern(base);
  }

  get size(): number {
    return this.list.length;
  }

  get(id: number | undefined): CellStyle {
    return this.list[id ?? 0] ?? this.list[0];
  }

  get defaultStyle(): CellStyle {
    return this.list[0];
  }

  intern(style: CellStyle): number {
    const s = normalizeStyle(style);
    const key = styleKey(s);
    const existing = this.index.get(key);
    if (existing !== undefined) return existing;
    const id = this.list.length;
    this.list.push(s);
    this.index.set(key, id);
    return id;
  }

  /** Appends a style even if an identical one exists (keeps file xf indexes stable on import). */
  push(style: CellStyle): number {
    const s = normalizeStyle(style);
    const id = this.list.length;
    this.list.push(s);
    const key = styleKey(s);
    if (!this.index.has(key)) this.index.set(key, id);
    return id;
  }

  patch(id: number | undefined, patch: StylePatch): number {
    return this.intern(applyPatch(this.get(id), patch));
  }

  all(): readonly CellStyle[] {
    return this.list;
  }
}

/** Excel's built-in number formats (numFmtId < 164). */
export const BUILTIN_NUMFMTS: Record<number, string> = {
  0: 'General',
  1: '0',
  2: '0.00',
  3: '#,##0',
  4: '#,##0.00',
  9: '0%',
  10: '0.00%',
  11: '0.00E+00',
  12: '# ?/?',
  13: '# ??/??',
  14: 'm/d/yyyy',
  15: 'd-mmm-yy',
  16: 'd-mmm',
  17: 'mmm-yy',
  18: 'h:mm AM/PM',
  19: 'h:mm:ss AM/PM',
  20: 'h:mm',
  21: 'h:mm:ss',
  22: 'm/d/yyyy h:mm',
  37: '#,##0 ;(#,##0)',
  38: '#,##0 ;[Red](#,##0)',
  39: '#,##0.00;(#,##0.00)',
  40: '#,##0.00;[Red](#,##0.00)',
  41: '_(* #,##0_);_(* (#,##0);_(* "-"_);_(@_)',
  42: '_("$"* #,##0_);_("$"* (#,##0);_("$"* "-"_);_(@_)',
  43: '_(* #,##0.00_);_(* (#,##0.00);_(* "-"??_);_(@_)',
  44: '_("$"* #,##0.00_);_("$"* (#,##0.00);_("$"* "-"??_);_(@_)',
  45: 'mm:ss',
  46: '[h]:mm:ss',
  47: 'mm:ss.0',
  48: '##0.0E+0',
  49: '@',
};
