/** SpreadsheetML colors (rgb, theme + tint, indexed, auto) and their resolution to CSS. */

export type ColorSpec =
  | { rgb: string; tint?: number }
  | { theme: number; tint?: number }
  | { indexed: number; tint?: number }
  | { auto: true };

/** Excel's default indexed palette (64 entries; 64 and 65 are system foreground/background). */
export const INDEXED_COLORS = [
  '000000', 'FFFFFF', 'FF0000', '00FF00', '0000FF', 'FFFF00', 'FF00FF', '00FFFF',
  '000000', 'FFFFFF', 'FF0000', '00FF00', '0000FF', 'FFFF00', 'FF00FF', '00FFFF',
  '800000', '008000', '000080', '808000', '800080', '008080', 'C0C0C0', '808080',
  '9999FF', '993366', 'FFFFCC', 'CCFFFF', '660066', 'FF8080', '0066CC', 'CCCCFF',
  '000080', 'FF00FF', 'FFFF00', '00FFFF', '800080', '800000', '008080', '0000FF',
  '00CCFF', 'CCFFFF', 'CCFFCC', 'FFFF99', '99CCFF', 'FF99CC', 'CC99FF', 'FFCC99',
  '3366FF', '33CCCC', '99CC00', 'FFCC00', 'FF9900', 'FF6600', '666699', '969696',
  '003366', '339966', '003300', '333300', '993300', '993366', '333399', '333333',
  '000000', 'FFFFFF',
];

/**
 * Theme color slots in SpreadsheetML index order. Excel swaps the first two pairs relative to
 * the DrawingML scheme: 0 = lt1 (background 1), 1 = dk1 (text 1), 2 = lt2, 3 = dk2.
 */
export const THEME_SLOTS = ['lt1', 'dk1', 'lt2', 'dk2', 'accent1', 'accent2', 'accent3', 'accent4', 'accent5', 'accent6', 'hlink', 'folHlink'] as const;

export type ThemeSlot = (typeof THEME_SLOTS)[number];

/** The Office theme (2013-2022) colors, used when a file has no theme part. */
export const DEFAULT_THEME_COLORS: Record<ThemeSlot, string> = {
  dk1: '000000',
  lt1: 'FFFFFF',
  dk2: '44546A',
  lt2: 'E7E6E6',
  accent1: '4472C4',
  accent2: 'ED7D31',
  accent3: 'A5A5A5',
  accent4: 'FFC000',
  accent5: '5B9BD5',
  accent6: '70AD47',
  hlink: '0563C1',
  folHlink: '954F72',
};

export interface Theme {
  colors: Record<ThemeSlot, string>;
  majorFont: string;
  minorFont: string;
  /** Custom indexed palette from styles.xml (replaces the default 64 colors). */
  indexed?: string[];
}

export function defaultTheme(): Theme {
  return { colors: { ...DEFAULT_THEME_COLORS }, majorFont: 'Calibri Light', minorFont: 'Calibri' };
}

function rgbToHls(r: number, g: number, b: number): [number, number, number] {
  r /= 255;
  g /= 255;
  b /= 255;
  const max = Math.max(r, g, b);
  const min = Math.min(r, g, b);
  const l = (max + min) / 2;
  if (max === min) return [0, l, 0];
  const d = max - min;
  const s = l > 0.5 ? d / (2 - max - min) : d / (max + min);
  let h: number;
  if (max === r) h = (g - b) / d + (g < b ? 6 : 0);
  else if (max === g) h = (b - r) / d + 2;
  else h = (r - g) / d + 4;
  return [h / 6, l, s];
}

function hue(p: number, q: number, t: number): number {
  if (t < 0) t += 1;
  if (t > 1) t -= 1;
  if (t < 1 / 6) return p + (q - p) * 6 * t;
  if (t < 1 / 2) return q;
  if (t < 2 / 3) return p + (q - p) * (2 / 3 - t) * 6;
  return p;
}

function hlsToRgb(h: number, l: number, s: number): [number, number, number] {
  if (s === 0) return [l * 255, l * 255, l * 255];
  const q = l < 0.5 ? l * (1 + s) : l + s - l * s;
  const p = 2 * l - q;
  return [hue(p, q, h + 1 / 3) * 255, hue(p, q, h) * 255, hue(p, q, h - 1 / 3) * 255];
}

/** Applies an OOXML tint (-1..1) to an RRGGBB color. */
export function applyTint(hex: string, tint: number): string {
  if (!tint) return hex;
  const r = parseInt(hex.slice(0, 2), 16);
  const g = parseInt(hex.slice(2, 4), 16);
  const b = parseInt(hex.slice(4, 6), 16);
  const [h, l, s] = rgbToHls(r, g, b);
  const l2 = tint < 0 ? l * (1 + tint) : l * (1 - tint) + tint;
  const [r2, g2, b2] = hlsToRgb(h, Math.max(0, Math.min(1, l2)), s);
  const to = (v: number) => Math.round(Math.max(0, Math.min(255, v))).toString(16).padStart(2, '0');
  return `${to(r2)}${to(g2)}${to(b2)}`.toUpperCase();
}

/** Normalises 'FF112233', '#112233', '112233' or 'rgb(1,2,3)' to 'RRGGBB'. Returns null if invalid. */
export function normalizeHex(input: string): string | null {
  let s = input.trim();
  const rgb = /^rgba?\(\s*(\d+)\s*,\s*(\d+)\s*,\s*(\d+)/i.exec(s);
  if (rgb) {
    return [rgb[1], rgb[2], rgb[3]].map((v) => Math.min(255, Number(v)).toString(16).padStart(2, '0')).join('').toUpperCase();
  }
  if (s.startsWith('#')) s = s.slice(1);
  if (/^[0-9a-f]{3}$/i.test(s)) s = s.split('').map((ch) => ch + ch).join('');
  if (/^[0-9a-f]{8}$/i.test(s)) return s.slice(2).toUpperCase();
  if (/^[0-9a-f]{6}$/i.test(s)) return s.toUpperCase();
  return null;
}

/** Resolves a color spec to 'RRGGBB'; `auto` and unknown colors resolve to `fallback`. */
export function resolveColor(spec: ColorSpec | undefined, theme: Theme, fallback: string): string {
  if (!spec) return fallback;
  let hex: string | null = null;
  if ('rgb' in spec) hex = normalizeHex(spec.rgb);
  else if ('theme' in spec) hex = theme.colors[THEME_SLOTS[spec.theme] ?? 'dk1'] ?? null;
  else if ('indexed' in spec) {
    if (spec.indexed === 64) hex = fallback;
    else if (spec.indexed === 65) hex = 'FFFFFF';
    else hex = theme.indexed?.[spec.indexed] ?? INDEXED_COLORS[spec.indexed] ?? null;
  } else hex = fallback;
  if (!hex) return fallback;
  const tint = 'tint' in spec ? spec.tint : undefined;
  return tint ? applyTint(hex, tint) : hex;
}

export function cssColor(hex: string): string {
  return `#${hex}`;
}

export function colorEquals(a: ColorSpec | undefined, b: ColorSpec | undefined): boolean {
  return JSON.stringify(a ?? null) === JSON.stringify(b ?? null);
}

export function rgbSpec(cssOrHex: string): ColorSpec | undefined {
  const hex = normalizeHex(cssOrHex);
  return hex ? { rgb: `FF${hex}` } : undefined;
}
