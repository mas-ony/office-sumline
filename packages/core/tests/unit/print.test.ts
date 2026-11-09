import { describe, expect, it } from 'vitest';
import { unzlibSync, strFromU8 } from 'fflate';
import { SpreadsheetDocument } from '../../src/core/document';
import { approxMeasure, buildPages, DEFAULT_PRINT_OPTIONS, type PrintOptions } from '../../src/print/pages';
import { jpegSize, pagesToPdf, winAnsi } from '../../src/print/pdf';
import { pageSvg } from '../../src/print/svg';
import { parseXml } from '../../src/core/xml/xml';
import { makeRange } from '../../src/core/refs';
import { rgbSpec } from '../../src/core/model/color';

function bigDoc(rows: number, cols: number): SpreadsheetDocument {
  const d = SpreadsheetDocument.blank('Report.xlsx');
  const data = Array.from({ length: rows }, (_, r) => Array.from({ length: cols }, (_, c) => (r === 0 ? `Head ${c + 1}` : r * 100 + c)));
  d.setValues(0, 0, data);
  return d;
}

const opts = (o: Partial<PrintOptions> = {}): PrintOptions => ({ ...DEFAULT_PRINT_OPTIONS, ...o });
const texts = (page: { ops: { k: string; text?: string }[] }) => page.ops.filter((o) => o.k === 'text').map((o) => o.text);

describe('print layout', () => {
  it('prints the used range on one page when it fits', () => {
    const d = bigDoc(5, 4);
    const pages = buildPages(d, opts(), approxMeasure, { title: 'Report.xlsx' });
    expect(pages).toHaveLength(1);
    expect(pages[0].width).toBeGreaterThan(pages[0].height);
    expect(texts(pages[0])).toContain('Head 1');
    expect(texts(pages[0])).toContain('403');
  });

  it('splits long and wide sheets into pages in the chosen order', () => {
    const d = bigDoc(120, 20);
    const across = buildPages(d, opts({ orientation: 'portrait' }), approxMeasure, { title: 'x' });
    expect(across.length).toBeGreaterThan(2);
    // Left to right: the second page continues the first rows to the right.
    expect(texts(across[0])).not.toContain('Head 20');
    expect(texts(across[1])).toContain('Head 20');
    const down = buildPages(d, opts({ orientation: 'portrait', order: 'topToBottom' }), approxMeasure, { title: 'x' });
    expect(down).toHaveLength(across.length);
    expect(texts(down[1])).not.toContain('Head 1');
  });

  it('repeats frozen rows on every page', () => {
    const d = bigDoc(150, 3);
    d.setFreeze(1, 0);
    const pages = buildPages(d, opts({ orientation: 'portrait' }), approxMeasure, { title: 'x' });
    expect(pages.length).toBeGreaterThan(1);
    for (const p of pages) expect(texts(p)).toContain('Head 1');
    const plain = buildPages(d, opts({ orientation: 'portrait', repeatRows: false }), approxMeasure, { title: 'x' });
    expect(texts(plain[1])).not.toContain('Head 1');
  });

  it('fits the sheet on one page when asked', () => {
    const d = bigDoc(150, 30);
    expect(buildPages(d, opts({ scale: 'fitPage' }), approxMeasure, { title: 'x' })).toHaveLength(1);
    const fitWidth = buildPages(d, opts({ scale: 'fitWidth' }), approxMeasure, { title: 'x' });
    expect(fitWidth.length).toBeGreaterThan(1);
    expect(texts(fitWidth[0])).toContain('Head 30');
  });

  it('adds headings, titles, page numbers and a notes page', () => {
    const d = bigDoc(4, 3);
    d.setNote(1, 1, 'Check this value');
    const pages = buildPages(d, opts({ headings: true, pageNumbers: true, spreadsheetTitle: true, worksheetTitle: true, notes: true }), approxMeasure, { title: 'Report.xlsx' });
    expect(pages).toHaveLength(2);
    const first = texts(pages[0]);
    expect(first).toEqual(expect.arrayContaining(['A', 'C', '1', '4', 'Report.xlsx - Sheet1', 'Page 1 of 2']));
    expect(texts(pages[1])).toEqual(expect.arrayContaining(['B2', 'Check this value', 'Page 2 of 2']));
  });

  it('prints only the selection when asked', () => {
    const d = bigDoc(10, 5);
    d.goTo('B2:C3');
    const pages = buildPages(d, opts({ scope: 'selection' }), approxMeasure, { title: 'x' });
    expect(texts(pages[0]).sort()).toEqual(['101', '102', '201', '202']);
  });

  it('draws pages as valid SVG', () => {
    const d = bigDoc(6, 4);
    d.setInput(6, 0, 'a < b & "c"');
    const [page] = buildPages(d, opts({ headings: true }), approxMeasure, { title: 'x' });
    expect(() => parseXml(pageSvg(page))).not.toThrow();
  });
});

describe('PDF output', () => {
  const env = {
    measure: approxMeasure,
    rasterize: async (_src: unknown, width: number, height: number) => ({ width, height, rgba: new Uint8Array(width * height * 4).fill(200) }),
  };

  it('writes a well-formed PDF with a correct cross-reference table', async () => {
    const d = bigDoc(80, 12);
    d.setInput(1, 1, 'Ünïcödé €100');
    d.setInput(2, 1, '中文');
    d.format({ fill: { pattern: 'solid', fg: rgbSpec('#ffff00')! } }, [makeRange(3, 1, 4, 2)]);
    const pages = buildPages(d, opts({ orientation: 'portrait' }), approxMeasure, { title: 'x' });
    const bytes = await pagesToPdf(pages, env, { title: 'Report' });
    const text = strFromU8(bytes, true);
    expect(text.startsWith('%PDF-1.4')).toBe(true);
    expect(text.trimEnd().endsWith('%%EOF')).toBe(true);
    const startxref = Number(/startxref\n(\d+)/.exec(text)![1]);
    expect(text.slice(startxref, startxref + 4)).toBe('xref');
    const entries = [...text.slice(startxref).matchAll(/^(\d{10}) 00000 n $/gm)].map((m) => Number(m[1]));
    expect(entries.length).toBeGreaterThan(pages.length * 2);
    entries.forEach((off, i) => expect(text.slice(off, off + 12)).toMatch(new RegExp(`^${i + 1} 0 obj`)));
    expect(text).toContain(`/Count ${pages.length}`);
    expect(text).toContain('/BaseFont /Helvetica');
    // Content streams decompress to drawing operators; CJK text became an image.
    const streams = [...text.matchAll(/<< ([^>]*?) \/Length (\d+) >>\nstream\n/g)].map((m) => ({ dict: m[1], start: m.index! + m[0].length, length: Number(m[2]) }));
    const contents = streams.filter((st) => !st.dict.includes('/Subtype /Image'));
    expect(contents).toHaveLength(pages.length);
    const content = strFromU8(unzlibSync(bytes.subarray(contents[0].start, contents[0].start + contents[0].length)));
    expect(content).toContain('Tf');
    expect(content).toContain('re f');
    expect(content).toContain('Do Q');
    expect(content).toContain('(\\334n\\357c\\366d\\351 \\200100) Tj');
  });

  it('encodes Windows-1252 text and reads JPEG sizes', () => {
    expect(winAnsi('Café €5 – “ok”')).toEqual([67, 97, 102, 233, 32, 128, 53, 32, 150, 32, 147, 111, 107, 148]);
    expect(winAnsi('日本')).toBeNull();
    const jpeg = new Uint8Array([0xff, 0xd8, 0xff, 0xe0, 0, 4, 0, 0, 0xff, 0xc0, 0, 17, 8, 0, 20, 0, 30, 3, 0, 0]);
    expect(jpegSize(jpeg)).toEqual({ width: 30, height: 20, components: 3 });
  });
});
