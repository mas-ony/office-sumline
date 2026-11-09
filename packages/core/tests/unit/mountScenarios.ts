/**
 * Scenarios for the mount() of the React and Vue packages (run by their own tests): the same
 * library behaviour from both interfaces.
 */
import { afterEach, expect, it, vi } from 'vitest';
import { version } from '../../package.json';
import type { SpreadsheetApi } from '../../src/api/api';
import type { SumlineOptions } from '../../src/ui/controller';
import { readFixture } from './helpers';

interface Mounted {
  ready: Promise<SpreadsheetApi>;
  update(options: SumlineOptions): void;
  destroy(): void;
}

const wait = { timeout: 3000 };
/** Opening a workbook in jsdom takes a moment; the first test also loads the modules. */
const slow = 20_000;

/** The element once it exists. */
function find<T extends Element = HTMLElement>(selector: string, root: ParentNode = document): Promise<T> {
  return vi.waitFor(() => {
    const el = root.querySelector<T>(selector);
    if (!el) throw new Error(`${selector} not found`);
    return el;
  }, wait);
}

/** Type into a text box the way the browser does (React and Vue both see the input event). */
function type(input: HTMLInputElement, text: string) {
  Object.getOwnPropertyDescriptor(HTMLInputElement.prototype, 'value')!.set!.call(input, text);
  input.dispatchEvent(new Event('input', { bubbles: true }));
}

/** The name shown in the editor's title row. */
const shownName = (el: HTMLElement) => el.querySelector('[data-testid="file-name"]')?.textContent;
/** Whether the editor marks its workbook as view-only. */
const viewOnly = (el: HTMLElement) => [...el.querySelectorAll('.sl-hint')].some((h) => h.textContent === '(view only)');

export function mountScenarios(mount: (element: HTMLElement, options?: SumlineOptions) => Mounted) {
  const open: Mounted[] = [];
  const place = () => document.body.appendChild(document.createElement('div'));
  const start = (el: HTMLElement, options: SumlineOptions) => {
    const m = mount(el, options);
    open.push(m);
    return m;
  };

  afterEach(() => {
    for (const m of open.splice(0)) m.destroy();
    document.body.innerHTML = '';
    localStorage.clear();
  });

  it('opens the workbook, keeps its classes off <html> and removes everything again', async () => {
    const el = place();
    const m = start(el, { document: { data: readFixture('kitchen-sink.xlsx'), name: 'Quarterly report.xlsx' } });
    const api = await m.ready;
    expect(api.doc.fileName).toBe('Quarterly report.xlsx');
    expect(api.version).toBe(version);
    expect(api.getValue('Sales!A1')).toBe('Quarterly Sales Report');
    await vi.waitFor(() => expect(el.contains(api.grid!.root)).toBe(true), wait);
    await vi.waitFor(() => expect(shownName(el)).toBe('Quarterly report.xlsx'), wait);
    expect(document.querySelectorAll('.sl-layer')).toHaveLength(1);
    expect(document.documentElement.className).toBe('');
    expect(document.body.className).toBe('');

    m.update({ document: { data: readFixture('modern.xlsx'), name: 'Stock.xlsx' } });
    await vi.waitFor(() => expect(api.worksheets()).toEqual(['Stock']), wait);
    expect(api.getValue('A1')).toBe('Item');
    await vi.waitFor(() => expect(shownName(el)).toBe('Stock.xlsx'), wait);
    expect(document.querySelectorAll('.sl-layer')).toHaveLength(1);

    m.destroy();
    open.length = 0;
    await vi.waitFor(() => expect(document.querySelector('.sl-layer')).toBeNull(), wait);
    expect(el.childElementCount).toBe(0);
  }, slow);

  it('shows menus and dialogs in its layer and reports a new name', async () => {
    const el = place();
    const m = start(el, { document: { data: readFixture('modern.xlsx'), name: 'Stock.xlsx' } });
    const api = await m.ready;
    const onFile = vi.fn();
    api.on('file', onFile);

    (await find('[data-testid="menu-file"]', el)).click();
    const rename = await find('[data-testid="menu-rename"]');
    expect(rename.closest('.sl-layer')).not.toBeNull();
    rename.click();
    const dialog = await find('[data-testid="prompt-dialog"]');
    expect(dialog.closest('.sl-layer')).not.toBeNull();
    expect(el.contains(dialog)).toBe(false);
    const input = await find<HTMLInputElement>('[data-testid="prompt-input"]', dialog);
    expect(input.value).toBe('Stock');
    type(input, 'Stock 2026');
    await vi.waitFor(() => expect(input.value).toBe('Stock 2026'), wait);
    (await find('[data-testid="prompt-ok"]', dialog)).click();

    await vi.waitFor(() => expect(onFile).toHaveBeenCalledWith({ name: 'Stock 2026.xlsx' }), wait);
    expect(api.doc.fileName).toBe('Stock 2026.xlsx');
    await vi.waitFor(() => expect(shownName(el)).toBe('Stock 2026.xlsx'), wait);
    await vi.waitFor(() => expect(document.querySelector('[data-testid="prompt-dialog"]')).toBeNull(), wait);
  }, slow);

  it('keeps two editors on one page apart', async () => {
    const a = place();
    const b = place();
    const [ra, rb] = await Promise.all([
      start(a, { document: { data: readFixture('kitchen-sink.xlsx'), name: 'Report.xlsx' } }).ready,
      start(b, { document: { data: readFixture('modern.xlsx'), name: 'Stock.xlsx' }, readOnly: true }).ready,
    ]);
    expect(ra).not.toBe(rb);
    expect(ra.worksheets()).toContain('Sales');
    expect(rb.worksheets()).toEqual(['Stock']);
    await vi.waitFor(() => {
      expect(a.contains(ra.grid!.root)).toBe(true);
      expect(b.contains(rb.grid!.root)).toBe(true);
    }, wait);
    expect(document.querySelectorAll('.sl-layer')).toHaveLength(2);

    ra.setValue('Sales!A1', 'Only in A');
    expect(ra.getValue('Sales!A1')).toBe('Only in A');
    expect(rb.getValue('A1')).toBe('Item');
    ra.select('C3');
    expect(ra.getSelected()).toBe('C3');
    expect(rb.getSelected()).not.toBe('C3');

    await vi.waitFor(() => expect(viewOnly(b)).toBe(true), wait);
    expect(viewOnly(a)).toBe(false);
  }, slow);
}
