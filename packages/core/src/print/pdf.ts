/**
 * A small PDF writer for printed pages: vector rectangles, lines and text in the standard PDF
 * fonts (Helvetica, Times and Courier; the bundled Arimo, Tinos and Cousine have the same
 * metrics, so the layout measured in the browser matches), JPEG pictures as they are, and other
 * images (charts, shapes, PNG pictures) as pixels with transparency. Text that the standard fonts
 * cannot show (outside Windows-1252, such as CJK) is drawn as an image.
 */
import { zlibSync } from 'fflate';
import type { FontSpec, ImageSrc, Measure, Op, Page } from './pages';
import { FAMILY_CSS } from './svg';

export interface Raster {
  width: number;
  height: number;
  rgba: Uint8ClampedArray | Uint8Array;
}

export interface PdfEnv {
  /** Pixels of an image drawn at a size in pixels (null when it cannot be drawn). */
  rasterize(src: ImageSrc, width: number, height: number): Promise<Raster | null>;
  measure: Measure;
}

const WIN_ANSI: Record<string, number> = {
  '€': 0x80, '‚': 0x82, 'ƒ': 0x83, '„': 0x84, '…': 0x85, '†': 0x86, '‡': 0x87, 'ˆ': 0x88, '‰': 0x89,
  'Š': 0x8a, '‹': 0x8b, 'Œ': 0x8c, 'Ž': 0x8e, '‘': 0x91, '’': 0x92, '“': 0x93, '”': 0x94, '•': 0x95,
  '–': 0x96, '—': 0x97, '˜': 0x98, '™': 0x99, 'š': 0x9a, '›': 0x9b, 'œ': 0x9c, 'ž': 0x9e, 'Ÿ': 0x9f,
};

/** Windows-1252 bytes of a text, or null when a character is not in it. */
export function winAnsi(s: string): number[] | null {
  const out: number[] = [];
  for (const ch of s) {
    const code = ch.codePointAt(0)!;
    if ((code >= 0x20 && code <= 0x7e) || (code >= 0xa0 && code <= 0xff)) out.push(code);
    else if (WIN_ANSI[ch] !== undefined) out.push(WIN_ANSI[ch]);
    else if (code === 9) out.push(0x20);
    else return null;
  }
  return out;
}

function pdfString(bytes: number[]): string {
  let s = '(';
  for (const b of bytes) {
    if (b === 0x28 || b === 0x29 || b === 0x5c) s += `\\${String.fromCharCode(b)}`;
    else if (b < 0x20 || b > 0x7e) s += `\\${b.toString(8).padStart(3, '0')}`;
    else s += String.fromCharCode(b);
  }
  return `${s})`;
}

/** A text string for the document information dictionary (UTF-16 when needed). */
function infoString(s: string): string {
  const bytes = winAnsi(s);
  if (bytes) return pdfString(bytes);
  let hex = 'FEFF';
  for (let i = 0; i < s.length; i++) hex += s.charCodeAt(i).toString(16).padStart(4, '0');
  return `<${hex}>`;
}

const FONT_NAMES: Record<FontSpec['family'], string[]> = {
  sans: ['Helvetica', 'Helvetica-Bold', 'Helvetica-Oblique', 'Helvetica-BoldOblique'],
  serif: ['Times-Roman', 'Times-Bold', 'Times-Italic', 'Times-BoldItalic'],
  mono: ['Courier', 'Courier-Bold', 'Courier-Oblique', 'Courier-BoldOblique'],
};

const num = (v: number) => {
  const r = Math.round(v * 1000) / 1000;
  return Object.is(r, -0) ? '0' : String(r);
};

function rgb(hex: string): string {
  const m = /^#?([0-9a-f]{6})$/i.exec(hex.trim());
  const v = m ? parseInt(m[1], 16) : 0;
  return `${num(((v >> 16) & 255) / 255)} ${num(((v >> 8) & 255) / 255)} ${num((v & 255) / 255)}`;
}

/** Width and height of a baseline or progressive JPEG (SOFn marker). */
export function jpegSize(data: Uint8Array): { width: number; height: number; components: number } | null {
  if (data[0] !== 0xff || data[1] !== 0xd8) return null;
  let i = 2;
  while (i + 9 < data.length) {
    if (data[i] !== 0xff) {
      i++;
      continue;
    }
    const marker = data[i + 1];
    const len = (data[i + 2] << 8) | data[i + 3];
    if (marker >= 0xc0 && marker <= 0xcf && marker !== 0xc4 && marker !== 0xc8 && marker !== 0xcc) {
      return { height: (data[i + 5] << 8) | data[i + 6], width: (data[i + 7] << 8) | data[i + 8], components: data[i + 9] };
    }
    i += 2 + len;
  }
  return null;
}

class PdfFile {
  private chunks: Uint8Array[] = [];
  private size = 0;
  private offsets: number[] = [0];
  private enc = new TextEncoder();

  constructor() {
    this.raw('%PDF-1.4\n%âãÏÓ\n');
  }

  private raw(s: string | Uint8Array): void {
    const bytes = typeof s === 'string' ? this.latin1(s) : s;
    this.chunks.push(bytes);
    this.size += bytes.length;
  }

  private latin1(s: string): Uint8Array {
    // Object bodies are ASCII except the binary comment of the header.
    if (/^[\x00-\x7f]*$/.test(s)) return this.enc.encode(s);
    const out = new Uint8Array(s.length);
    for (let i = 0; i < s.length; i++) out[i] = s.charCodeAt(i) & 0xff;
    return out;
  }

  reserve(): number {
    this.offsets.push(-1);
    return this.offsets.length - 1;
  }

  object(id: number, body: string): void {
    this.offsets[id] = this.size;
    this.raw(`${id} 0 obj\n${body}\nendobj\n`);
  }

  stream(id: number, dict: string, data: Uint8Array): void {
    this.offsets[id] = this.size;
    this.raw(`${id} 0 obj\n<< ${dict} /Length ${data.length} >>\nstream\n`);
    this.raw(data);
    this.raw('\nendstream\nendobj\n');
  }

  finish(root: number, info: number): Uint8Array {
    const xref = this.size;
    const lines = [`xref\n0 ${this.offsets.length}\n`, '0000000000 65535 f \n'];
    for (let i = 1; i < this.offsets.length; i++) lines.push(`${String(Math.max(0, this.offsets[i])).padStart(10, '0')} 00000 n \n`);
    this.raw(lines.join(''));
    this.raw(`trailer\n<< /Size ${this.offsets.length} /Root ${root} 0 R /Info ${info} 0 R >>\nstartxref\n${xref}\n%%EOF\n`);
    const out = new Uint8Array(this.size);
    let at = 0;
    for (const c of this.chunks) {
      out.set(c, at);
      at += c.length;
    }
    return out;
  }
}

function textSvg(op: Extract<Op, { k: 'text' }>, w: number, h: number): string {
  const esc = (s: string) => s.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
  const f = op.font;
  return `<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 ${num(w)} ${num(h)}" width="${num(w)}" height="${num(h)}"><text x="1" y="${num(f.size * 0.95)}" font-family='${FAMILY_CSS[f.family]}' font-size="${num(f.size)}"${f.bold ? ' font-weight="bold"' : ''}${f.italic ? ' font-style="italic"' : ''} fill="${op.color}" xml:space="preserve">${esc(op.text)}</text></svg>`;
}

export interface PdfInfo {
  title?: string;
}

/** Writes the pages as a PDF document. */
export async function pagesToPdf(pages: Page[], env: PdfEnv, info: PdfInfo = {}): Promise<Uint8Array> {
  const pdf = new PdfFile();
  const catalog = pdf.reserve();
  const pagesId = pdf.reserve();
  const infoId = pdf.reserve();
  const fontIds = new Map<string, number>();
  const fontRef = (font: FontSpec): string => {
    const name = FONT_NAMES[font.family][(font.bold ? 1 : 0) + (font.italic ? 2 : 0)];
    if (!fontIds.has(name)) fontIds.set(name, pdf.reserve());
    return `F${[...fontIds.keys()].indexOf(name) + 1}`;
  };
  const pageIds: number[] = [];
  const pageBodies = new Map<number, string>();
  const jpegCache = new Map<Uint8Array, { id: number; width: number; height: number }>();
  for (const page of pages) {
    const H = page.height;
    const content: string[] = [];
    const xobjects: string[] = [];
    let imageCount = 0;
    const clipOps = (c: { x: number; y: number; w: number; h: number } | undefined) => (c ? `${num(c.x)} ${num(H - c.y - c.h)} ${num(c.w)} ${num(c.h)} re W n ` : '');
    const addImage = async (x: number, y: number, w: number, h: number, src: ImageSrc, clip?: { x: number; y: number; w: number; h: number }) => {
      let id: number | null = null;
      if (src.kind === 'bitmap' && /jpe?g/i.test(src.mime)) {
        let hit = jpegCache.get(src.data);
        const size = hit ? null : jpegSize(src.data);
        if (!hit && size && (size.components === 3 || size.components === 1)) {
          const jid = pdf.reserve();
          pdf.stream(jid, `/Type /XObject /Subtype /Image /Width ${size.width} /Height ${size.height} /ColorSpace /${size.components === 1 ? 'DeviceGray' : 'DeviceRGB'} /BitsPerComponent 8 /Filter /DCTDecode`, src.data);
          hit = { id: jid, width: size.width, height: size.height };
          jpegCache.set(src.data, hit);
        }
        if (hit) id = hit.id;
      }
      if (id === null) {
        const scale = Math.min(3, Math.max(1, 1600 / Math.max(w, h)));
        const pw = Math.max(1, Math.min(3000, Math.round(w * scale * 1.5)));
        const ph = Math.max(1, Math.min(3000, Math.round(h * scale * 1.5)));
        const raster = await env.rasterize(src, pw, ph);
        if (!raster) return;
        const count = raster.width * raster.height;
        const color = new Uint8Array(count * 3);
        const alpha = new Uint8Array(count);
        let opaque = true;
        for (let i = 0; i < count; i++) {
          color[i * 3] = raster.rgba[i * 4];
          color[i * 3 + 1] = raster.rgba[i * 4 + 1];
          color[i * 3 + 2] = raster.rgba[i * 4 + 2];
          alpha[i] = raster.rgba[i * 4 + 3];
          if (alpha[i] !== 255) opaque = false;
        }
        let smask = '';
        if (!opaque) {
          const sid = pdf.reserve();
          pdf.stream(sid, `/Type /XObject /Subtype /Image /Width ${raster.width} /Height ${raster.height} /ColorSpace /DeviceGray /BitsPerComponent 8 /Filter /FlateDecode`, zlibSync(alpha));
          smask = ` /SMask ${sid} 0 R`;
        }
        id = pdf.reserve();
        pdf.stream(id, `/Type /XObject /Subtype /Image /Width ${raster.width} /Height ${raster.height} /ColorSpace /DeviceRGB /BitsPerComponent 8 /Filter /FlateDecode${smask}`, zlibSync(color));
      }
      const name = `Im${++imageCount}`;
      xobjects.push(`/${name} ${id} 0 R`);
      content.push(`q ${clipOps(clip)}${num(w)} 0 0 ${num(h)} ${num(x)} ${num(H - y - h)} cm /${name} Do Q`);
    };
    for (const op of page.ops) {
      switch (op.k) {
        case 'rect':
          content.push(`${rgb(op.fill)} rg ${num(op.x)} ${num(H - op.y - op.h)} ${num(op.w)} ${num(op.h)} re f`);
          break;
        case 'line':
          content.push(`q ${rgb(op.color)} RG ${num(op.width)} w [${(op.dash ?? []).map(num).join(' ')}] 0 d ${num(op.x1)} ${num(H - op.y1)} m ${num(op.x2)} ${num(H - op.y2)} l S Q`);
          break;
        case 'text': {
          const bytes = winAnsi(op.text);
          const width = env.measure(op.text, op.font);
          if (!bytes) {
            const h = op.font.size * 1.3;
            const w = width + 2;
            await addImage(op.x - 1, op.y - op.font.size * 0.95, w, h, { kind: 'svg', markup: textSvg(op, w, h), w, h }, op.clip);
            break;
          }
          const parts = [`q ${clipOps(op.clip)}BT /${fontRef(op.font)} ${num(op.font.size)} Tf ${rgb(op.color)} rg ${num(op.x)} ${num(H - op.y)} Td ${pdfString(bytes)} Tj ET`];
          const lw = Math.max(0.3, op.font.size * 0.06);
          if (op.underline) parts.push(`${rgb(op.color)} RG ${num(lw)} w ${num(op.x)} ${num(H - op.y - op.font.size * 0.12)} m ${num(op.x + width)} ${num(H - op.y - op.font.size * 0.12)} l S`);
          if (op.strike) parts.push(`${rgb(op.color)} RG ${num(lw)} w ${num(op.x)} ${num(H - op.y + op.font.size * 0.3)} m ${num(op.x + width)} ${num(H - op.y + op.font.size * 0.3)} l S`);
          content.push(`${parts.join(' ')} Q`);
          break;
        }
        case 'image':
          await addImage(op.x, op.y, op.w, op.h, op.src, op.clip);
          break;
      }
    }
    const contentId = pdf.reserve();
    pdf.stream(contentId, '/Filter /FlateDecode', zlibSync(new TextEncoder().encode(content.join('\n'))));
    const pageId = pdf.reserve();
    pageIds.push(pageId);
    // Fonts are listed once for the whole document (every page may use any of them).
    const fontsPlaceholder = '__FONTS__';
    pageBodies.set(pageId, `<< /Type /Page /Parent ${pagesId} 0 R /MediaBox [0 0 ${num(page.width)} ${num(page.height)}] /Resources << /Font ${fontsPlaceholder} /XObject << ${xobjects.join(' ')} >> >> /Contents ${contentId} 0 R >>`);
  }
  // Fonts.
  const fontDict = `<< ${[...fontIds.keys()].map((name, i) => `/F${i + 1} ${fontIds.get(name)} 0 R`).join(' ')} >>`;
  for (const [name, id] of fontIds) pdf.object(id, `<< /Type /Font /Subtype /Type1 /BaseFont /${name} /Encoding /WinAnsiEncoding >>`);
  for (const id of pageIds) pdf.object(id, pageBodies.get(id)!.replace('__FONTS__', fontDict));
  pdf.object(pagesId, `<< /Type /Pages /Kids [${pageIds.map((id) => `${id} 0 R`).join(' ')}] /Count ${pageIds.length} >>`);
  pdf.object(catalog, `<< /Type /Catalog /Pages ${pagesId} 0 R >>`);
  pdf.object(infoId, `<< /Producer (Sumline)${info.title ? ` /Title ${infoString(info.title)}` : ''} >>`);
  return pdf.finish(catalog, infoId);
}

