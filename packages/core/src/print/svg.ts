/** Printed pages as SVG: the print preview and the pages sent to the browser's print dialog. */
import type { FontSpec, ImageSrc, Op, Page } from './pages';

const esc = (s: string) => s.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
const n = (v: number) => String(Math.round(v * 100) / 100);

/** Fonts metric-compatible with the PDF standard fonts (bundled with the app). */
export const FAMILY_CSS: Record<FontSpec['family'], string> = {
  sans: 'Arimo, Arial, Helvetica, sans-serif',
  serif: 'Tinos, "Times New Roman", Times, serif',
  mono: 'Cousine, "Courier New", Courier, monospace',
};

function base64(bytes: Uint8Array): string {
  let s = '';
  for (let i = 0; i < bytes.length; i += 0x8000) s += String.fromCharCode(...bytes.subarray(i, i + 0x8000));
  return btoa(s);
}

/** An SVG document (chart, shape) placed into a box as a nested SVG element. */
function nestedSvg(src: Extract<ImageSrc, { kind: 'svg' }>, x: number, y: number, w: number, h: number): string {
  return src.markup.replace(/^<svg\b([^>]*)>/, (_m, attrs: string) => {
    const kept = attrs.replace(/\s(width|height|x|y)="[^"]*"/g, '');
    return `<svg${kept} x="${n(x)}" y="${n(y)}" width="${n(w)}" height="${n(h)}" preserveAspectRatio="none">`;
  });
}

function opSvg(op: Op): string {
  switch (op.k) {
    case 'rect':
      return `<rect x="${n(op.x)}" y="${n(op.y)}" width="${n(op.w)}" height="${n(op.h)}" fill="${op.fill}"/>`;
    case 'line':
      return `<line x1="${n(op.x1)}" y1="${n(op.y1)}" x2="${n(op.x2)}" y2="${n(op.y2)}" stroke="${op.color}" stroke-width="${n(op.width)}"${op.dash ? ` stroke-dasharray="${op.dash.map(n).join(' ')}"` : ''}/>`;
    case 'text': {
      const f = op.font;
      const deco = [op.underline ? 'underline' : '', op.strike ? 'line-through' : ''].filter(Boolean).join(' ');
      const t = (x: number, y: number) =>
        `<text x="${n(x)}" y="${n(y)}" font-family="${esc(FAMILY_CSS[f.family])}" font-size="${n(f.size)}"${f.bold ? ' font-weight="bold"' : ''}${f.italic ? ' font-style="italic"' : ''}${deco ? ` text-decoration="${deco}"` : ''} fill="${op.color}" xml:space="preserve">${esc(op.text)}</text>`;
      if (!op.clip) return t(op.x, op.y);
      const c = op.clip;
      return `<svg x="${n(c.x)}" y="${n(c.y)}" width="${n(c.w)}" height="${n(c.h)}" overflow="hidden">${t(op.x - c.x, op.y - c.y)}</svg>`;
    }
    case 'image': {
      const inner = op.src.kind === 'svg'
        ? nestedSvg(op.src, op.clip ? op.x - op.clip.x : op.x, op.clip ? op.y - op.clip.y : op.y, op.w, op.h)
        : `<image x="${n(op.clip ? op.x - op.clip.x : op.x)}" y="${n(op.clip ? op.y - op.clip.y : op.y)}" width="${n(op.w)}" height="${n(op.h)}" preserveAspectRatio="none" href="data:${op.src.mime};base64,${base64(op.src.data)}"/>`;
      if (!op.clip) return inner;
      const c = op.clip;
      return `<svg x="${n(c.x)}" y="${n(c.y)}" width="${n(c.w)}" height="${n(c.h)}" overflow="hidden">${inner}</svg>`;
    }
  }
}

/** One page as an SVG document sized in points. */
export function pageSvg(page: Page, opts: { width?: string; height?: string; className?: string } = {}): string {
  return `<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 ${n(page.width)} ${n(page.height)}" width="${opts.width ?? `${n(page.width)}pt`}" height="${opts.height ?? `${n(page.height)}pt`}"${opts.className ? ` class="${opts.className}"` : ''}><rect width="100%" height="100%" fill="#ffffff"/>${page.ops.map(opSvg).join('')}</svg>`;
}
