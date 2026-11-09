/** Font stacks for spreadsheet fonts and cached text measurement. */
import type { FontStyle } from '../core/model/style';

/** Metric-compatible open fonts are bundled for the common Office fonts. */
const STACKS: Record<string, string> = {
  calibri: '"Calibri", "Carlito", "Liberation Sans", Arial, sans-serif',
  'calibri light': '"Calibri Light", "Carlito", "Liberation Sans", Arial, sans-serif',
  aptos: '"Aptos", "Carlito", "Liberation Sans", Arial, sans-serif',
  'aptos narrow': '"Aptos Narrow", "Arial Narrow", "Carlito", sans-serif',
  arial: 'Arial, "Arimo", "Liberation Sans", Helvetica, sans-serif',
  helvetica: 'Helvetica, Arial, "Arimo", sans-serif',
  'times new roman': '"Times New Roman", "Tinos", "Liberation Serif", Times, serif',
  'courier new': '"Courier New", "Cousine", "Liberation Mono", monospace',
  courier: '"Courier New", "Cousine", monospace',
  cambria: 'Cambria, "Caladea", Georgia, serif',
  georgia: 'Georgia, "Tinos", serif',
  verdana: 'Verdana, "DejaVu Sans", sans-serif',
  tahoma: 'Tahoma, "DejaVu Sans", sans-serif',
  'segoe ui': '"Segoe UI", "Carlito", sans-serif',
};

export function fontStack(name: string): string {
  const key = name.trim().toLowerCase();
  return STACKS[key] ?? `"${name.replace(/"/g, '')}", "Carlito", Arial, sans-serif`;
}

/** CSS font shorthand for a cell font at a pixel scale (zoom). */
export function cssFont(font: Pick<FontStyle, 'name' | 'size' | 'bold' | 'italic'>, scale = 1): string {
  const px = (font.size * 96) / 72;
  return `${font.italic ? 'italic ' : ''}${font.bold ? 'bold ' : ''}${Math.max(1, px * scale).toFixed(2)}px ${fontStack(font.name)}`;
}

let measureCtx: CanvasRenderingContext2D | null = null;
const cache = new Map<string, number>();

function ctx(): CanvasRenderingContext2D | null {
  if (!measureCtx && typeof document !== 'undefined') {
    measureCtx = document.createElement('canvas').getContext('2d');
  }
  return measureCtx;
}

export function textWidth(text: string, font: string): number {
  const key = `${font}\u0000${text}`;
  let w = cache.get(key);
  if (w !== undefined) return w;
  const c = ctx();
  if (!c) return text.length * 7;
  c.font = font;
  w = c.measureText(text).width;
  if (cache.size > 20000) cache.clear();
  cache.set(key, w);
  return w;
}

/** Line height in pixels for a font size in points. */
export function lineHeightPx(sizePt: number, scale = 1): number {
  return Math.ceil(((sizePt * 96) / 72) * 1.22 * scale);
}

/** Breaks text into lines that fit `width` (word wrap, long words broken by character). */
export function wrapText(text: string, font: string, width: number): string[] {
  const out: string[] = [];
  for (const para of text.split('\n')) {
    if (!para) {
      out.push('');
      continue;
    }
    const words = para.split(/(\s+)/);
    let line = '';
    for (const w of words) {
      const candidate = line + w;
      if (textWidth(candidate, font) <= width || !line) {
        if (!line && textWidth(w, font) > width) {
          // Break a long word.
          let chunk = '';
          for (const ch of w) {
            if (textWidth(chunk + ch, font) > width && chunk) {
              out.push(chunk);
              chunk = ch;
            } else chunk += ch;
          }
          line = chunk;
        } else line = candidate;
      } else {
        out.push(line.replace(/\s+$/, ''));
        line = w.trimStart();
      }
    }
    out.push(line.replace(/\s+$/, ''));
  }
  return out;
}

export function clearMeasureCache(): void {
  cache.clear();
}
