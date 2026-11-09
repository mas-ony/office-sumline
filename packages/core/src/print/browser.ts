/**
 * Browser side of printing: text measurement with the bundled fonts, drawing images and charts to
 * pixels for the PDF, saving the PDF and sending the pages to the system print dialog.
 */
import type { FontSpec, ImageSrc, Measure, Page } from './pages';
import type { Raster } from './pdf';
import { FAMILY_CSS, pageSvg } from './svg';

/** Waits for the fonts used by printed pages (their metrics match the PDF standard fonts). */
export async function loadPrintFonts(): Promise<void> {
  const fonts = (document as Document & { fonts?: FontFaceSet }).fonts;
  if (!fonts?.load) return;
  const faces = ['Arimo', 'Tinos', 'Cousine'].flatMap((f) => [`12px ${f}`, `bold 12px ${f}`, `italic 12px ${f}`, `italic bold 12px ${f}`]);
  try {
    await Promise.race([Promise.all(faces.map((f) => fonts.load(f))), new Promise((r) => setTimeout(r, 2500))]);
  } catch {
    // Measurement falls back to the system fonts.
  }
}

/** Text width in points (the font size is in points). */
export function canvasMeasure(): Measure {
  const ctx = document.createElement('canvas').getContext('2d');
  const cache = new Map<string, number>();
  return (text: string, font: FontSpec) => {
    if (!ctx) return text.length * font.size * 0.52;
    const key = `${font.family}${font.bold ? 'b' : ''}${font.italic ? 'i' : ''}\u0000${text}`;
    let w = cache.get(key);
    if (w === undefined) {
      ctx.font = `${font.italic ? 'italic ' : ''}${font.bold ? 'bold ' : ''}100px ${FAMILY_CSS[font.family]}`;
      w = ctx.measureText(text).width / 100;
      if (cache.size > 20000) cache.clear();
      cache.set(key, w);
    }
    return w * font.size;
  };
}

function loadImage(url: string): Promise<HTMLImageElement> {
  return new Promise((resolve, reject) => {
    const img = new Image();
    img.onload = () => resolve(img);
    img.onerror = () => reject(new Error('image'));
    img.src = url;
  });
}

/** Pixels of a picture, chart or shape at a pixel size. */
export async function rasterize(src: ImageSrc, width: number, height: number): Promise<Raster | null> {
  const canvas = document.createElement('canvas');
  canvas.width = width;
  canvas.height = height;
  const ctx = canvas.getContext('2d');
  if (!ctx) return null;
  const blob = src.kind === 'svg' ? new Blob([src.markup], { type: 'image/svg+xml' }) : new Blob([src.data as BlobPart], { type: src.mime });
  const url = URL.createObjectURL(blob);
  try {
    const img = await loadImage(url);
    ctx.drawImage(img, 0, 0, width, height);
    return { width, height, rgba: ctx.getImageData(0, 0, width, height).data };
  } catch {
    return null;
  } finally {
    URL.revokeObjectURL(url);
  }
}

/** Saves bytes as a file (works without secure-context APIs, also on Android). */
export function saveFile(bytes: Uint8Array, name: string, type: string): void {
  const url = URL.createObjectURL(new Blob([bytes as BlobPart], { type }));
  const a = document.createElement('a');
  a.href = url;
  a.download = name;
  a.rel = 'noopener';
  document.body.appendChild(a);
  a.click();
  a.remove();
  setTimeout(() => URL.revokeObjectURL(url), 60_000);
}

/** Sends the pages to the browser's print dialog (vector pages, the paper size set by @page). */
export function printPages(pages: Page[]): void {
  document.getElementById('sl-print-root')?.remove();
  document.getElementById('sl-print-style')?.remove();
  const first = pages[0];
  const root = document.createElement('div');
  root.id = 'sl-print-root';
  root.innerHTML = pages.map((p) => `<div class="sl-printed-page" style="width:${p.width}pt;height:${p.height}pt">${pageSvg(p)}</div>`).join('');
  const style = document.createElement('style');
  style.id = 'sl-print-style';
  style.textContent = `
@page { size: ${first.width}pt ${first.height}pt; margin: 0; }
@media screen { #sl-print-root { display: none; } }
@media print {
  html, body { margin: 0 !important; padding: 0 !important; background: #fff !important; height: auto !important; overflow: visible !important; }
  body > *:not(#sl-print-root) { display: none !important; }
  #sl-print-root { display: block !important; }
  .sl-printed-page { overflow: hidden; break-after: page; page-break-after: always; }
  .sl-printed-page:last-child { break-after: auto; page-break-after: auto; }
  .sl-printed-page > svg { display: block; width: 100%; height: 100%; }
}`;
  document.head.appendChild(style);
  document.body.appendChild(root);
  const cleanup = () => {
    window.removeEventListener('afterprint', cleanup);
    root.remove();
    style.remove();
  };
  window.addEventListener('afterprint', cleanup);
  // Give the browser a frame to lay out the pages (images inside them decode asynchronously).
  setTimeout(() => window.print(), 120);
}
