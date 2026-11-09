import { type Page } from '@playwright/test';

/** Row/column indexes (0-based) of an A1 reference. */
export function rc(ref: string): { r: number; c: number } {
  const m = /^([A-Z]+)(\d+)$/.exec(ref);
  if (!m) throw new Error(`Bad reference ${ref}`);
  let c = 0;
  for (const ch of m[1]) c = c * 26 + ch.charCodeAt(0) - 64;
  return { r: Number(m[2]) - 1, c: c - 1 };
}

type Api = {
  grid: { transform(): { cellRect(r: number, c: number): { x: number; y: number; w: number; h: number } }; root: HTMLElement } | null;
  doc: { fileName: string; displayText(r: number, c: number): string; editText(r: number, c: number): string };
};

/**
 * Opens the editor: a new workbook, or a file from the dev server (`src`). Collects uncaught page
 * errors.
 */
export async function openApp(page: Page, src?: string): Promise<string[]> {
  const errors: string[] = [];
  page.on('pageerror', (e) => errors.push(`${e.message}\n${e.stack ?? ''}`));
  await page.goto(src ? `/?src=${encodeURIComponent(src)}` : '/?new=1');
  await page.waitForFunction(() => !!(window as unknown as { spreadsheetEditor?: Api }).spreadsheetEditor?.grid);
  if (src) {
    const name = src.split('/').pop()!.replace(/\.(xls|xlsb)$/i, '.xlsx');
    await page.waitForFunction((n) => (window as unknown as { spreadsheetEditor: Api }).spreadsheetEditor.doc.fileName === n, name);
  }
  return errors;
}

/** Client coordinates of a cell (center). */
export async function cellPoint(page: Page, ref: string): Promise<{ x: number; y: number }> {
  return page.evaluate(({ r, c }) => {
    const g = (window as unknown as { spreadsheetEditor: Api }).spreadsheetEditor.grid!;
    const rect = g.transform().cellRect(r, c);
    const root = g.root.getBoundingClientRect();
    return { x: root.left + rect.x + rect.w / 2, y: root.top + rect.y + rect.h / 2 };
  }, rc(ref));
}

export async function clickCell(page: Page, ref: string, opts: { button?: 'left' | 'right'; count?: number } = {}): Promise<void> {
  const p = await cellPoint(page, ref);
  await page.mouse.click(p.x, p.y, { button: opts.button, clickCount: opts.count });
}

export async function tapCell(page: Page, ref: string): Promise<void> {
  const p = await cellPoint(page, ref);
  await page.touchscreen.tap(p.x, p.y);
}

/** Types into a cell and commits with Enter. */
export async function typeIn(page: Page, ref: string, text: string): Promise<void> {
  await clickCell(page, ref);
  await page.keyboard.type(text);
  await page.keyboard.press('Enter');
}

/** Sets values through the page API (fast setup). */
export async function setData(page: Page, rows: (string | number | boolean | null)[][], at = 'A1'): Promise<void> {
  await page.evaluate(({ rows, at }) => (window as unknown as { spreadsheetEditor: { setData(d: unknown, a: string): void } }).spreadsheetEditor.setData(rows, at), { rows, at });
}

export async function select(page: Page, ref: string): Promise<void> {
  await page.evaluate((ref) => (window as unknown as { spreadsheetEditor: { select(r: string): void } }).spreadsheetEditor.select(ref), ref);
}

/** Displayed text of a cell of the active sheet. */
export async function shown(page: Page, ref: string): Promise<string> {
  return page.evaluate(({ r, c }) => (window as unknown as { spreadsheetEditor: Api }).spreadsheetEditor.doc.displayText(r, c), rc(ref));
}

/** What the formula bar shows for a cell. */
export async function input(page: Page, ref: string): Promise<string> {
  return page.evaluate(({ r, c }) => (window as unknown as { spreadsheetEditor: Api }).spreadsheetEditor.doc.editText(r, c), rc(ref));
}

/** Evaluates a function of the document in the page. */
export async function docEval<T>(page: Page, fn: string): Promise<T> {
  return page.evaluate((src) => {
    const d = (window as unknown as { spreadsheetEditor: { doc: unknown } }).spreadsheetEditor.doc;
    return new Function('doc', `return (${src})(doc);`)(d) as T;
  }, fn);
}

export const SALES = [
  ['Region', 'Q1', 'Q2', 'Q3'],
  ['North', 120, 150, 90],
  ['South', 90, 110, 130],
  ['East', 60, 80, 100],
];
