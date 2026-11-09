import { readFileSync } from 'node:fs';
import { expect, test } from '@playwright/test';
import { strFromU8, unzipSync } from 'fflate';
import { SALES, cellPoint, docEval, input, openApp, select, setData, shown, tapCell } from './helpers';
import { buildXls } from '../../packages/core/tests/unit/binaryWorkbooks';

test.describe('phone layout', () => {
  test('fits the screen, taps select and double taps edit with the key bar', async ({ page }) => {
    const errors = await openApp(page);
    await expect(page.getByTestId('tb-more')).toBeVisible();
    expect(await page.evaluate(() => document.documentElement.scrollWidth <= window.innerWidth + 1)).toBe(true);
    await tapCell(page, 'B3');
    await expect(page.getByTestId('name-box')).toHaveValue('B3');
    // Double tap opens the editor with the on-screen keyboard and the key bar.
    const p = await cellPoint(page, 'B3');
    await page.touchscreen.tap(p.x, p.y);
    await page.waitForTimeout(60);
    await page.touchscreen.tap(p.x, p.y);
    await expect(page.getByTestId('cell-editor')).toBeFocused();
    await expect(page.getByTestId('keybar')).toBeVisible();
    await page.keyboard.type('42');
    await page.keyboard.press('Enter');
    expect(await shown(page, 'B3')).toBe('42');
    await expect(page.getByTestId('keybar')).toHaveCount(0);
    expect(errors).toEqual([]);
  });

  test('the key bar types formula characters and functions; Done commits and Cancel discards', async ({ page }) => {
    const errors = await openApp(page);
    await setData(page, [[1], [2], [3]]);
    const bar = page.getByTestId('keybar');
    const editor = page.getByTestId('cell-editor');
    const doubleTap = async (ref: string) => {
      const p = await cellPoint(page, ref);
      await page.touchscreen.tap(p.x, p.y);
      await page.waitForTimeout(60);
      await page.touchscreen.tap(p.x, p.y);
      await expect(bar).toBeVisible();
    };
    const dialog = page.getByTestId('functions-dialog');
    const pickFunction = async (name: string) => {
      await bar.getByTestId('key-functions').tap();
      await dialog.getByTestId('functions-search').fill(name);
      await dialog.locator('.sl-suggest-item').filter({ hasText: new RegExp(`^${name}\\(`) }).tap();
      await expect(dialog).toHaveCount(0);
    };
    // Characters and functions from the bar go in at the caret; Done commits and moves down.
    await doubleTap('B1');
    await bar.getByRole('button', { name: 'Insert =' }).tap();
    await page.keyboard.type('A1');
    await bar.getByRole('button', { name: 'Insert +' }).tap();
    await pickFunction('SUM');
    await expect(editor).toHaveValue('=A1+SUM(');
    await expect(editor).toBeFocused();
    await page.keyboard.type('A2');
    await bar.getByRole('button', { name: 'Insert :' }).tap();
    await page.keyboard.type('A3');
    await bar.getByRole('button', { name: 'Insert )' }).tap();
    await expect(editor).toHaveValue('=A1+SUM(A2:A3)');
    await bar.getByTestId('key-done').tap();
    expect(await input(page, 'B1')).toBe('=A1+SUM(A2:A3)');
    expect(await shown(page, 'B1')).toBe('6');
    await expect(page.getByTestId('name-box')).toHaveValue('B2');
    await expect(bar).toHaveCount(0);
    // In a cell without a formula, Functions starts one; closing the list keeps the formula.
    await doubleTap('B2');
    await pickFunction('SUM');
    await expect(editor).toHaveValue('=SUM(');
    await page.keyboard.type('A1');
    await bar.getByRole('button', { name: 'Insert :' }).tap();
    await page.keyboard.type('A2');
    await bar.getByTestId('key-functions').tap();
    await dialog.getByTestId('modal-close').tap();
    await expect(dialog).toHaveCount(0);
    await expect(editor).toBeFocused();
    await expect(editor).toHaveValue('=SUM(A1:A2');
    await bar.getByRole('button', { name: 'Insert )' }).tap();
    await bar.getByTestId('key-done').tap();
    expect(await input(page, 'B2')).toBe('=SUM(A1:A2)');
    expect(await shown(page, 'B2')).toBe('3');
    // Cancel leaves the cell as it was.
    await doubleTap('A1');
    await page.keyboard.type('99');
    await expect(editor).toHaveValue('199');
    await bar.getByTestId('key-cancel').tap();
    await expect(bar).toHaveCount(0);
    expect(await input(page, 'A1')).toBe('1');
    expect(errors).toEqual([]);
  });

  test('long press opens the cell menu as a bottom sheet', async ({ page }) => {
    await openApp(page);
    const p = await cellPoint(page, 'C4');
    const cdp = await page.context().newCDPSession(page);
    await cdp.send('Input.dispatchTouchEvent', { type: 'touchStart', touchPoints: [{ x: p.x, y: p.y }] });
    await page.waitForTimeout(700);
    await cdp.send('Input.dispatchTouchEvent', { type: 'touchEnd', touchPoints: [] });
    const sheet = page.locator('.sl-bottom-sheet');
    await expect(sheet.getByTestId('ctx-chart')).toBeVisible();
    await sheet.getByTestId('ctx-format').tap();
    await expect(page.getByTestId('format-dialog')).toBeVisible();
  });

  test('opens an Excel 97-2003 file', async ({ page }) => {
    await page.goto('/');
    await page.getByTestId('open-input').setInputFiles({ name: 'Budget.xls', mimeType: 'application/vnd.ms-excel', buffer: Buffer.from(buildXls()) });
    await page.waitForFunction(() => (window as unknown as { spreadsheetEditor?: { doc: { fileName: string } } }).spreadsheetEditor?.doc.fileName === 'Budget.xlsx');
    expect(await shown(page, 'A1')).toBe('Region');
    expect(await shown(page, 'B2')).toBe('19.75');
  });

  test('formats from the overflow panel; lists stay inside the screen', async ({ page }) => {
    const errors = await openApp(page);
    await setData(page, SALES);
    await select(page, 'B2:D4');
    const width = page.viewportSize()!.width;
    // On a phone, Bold is in the overflow panel, which stays open for the next tool.
    await page.getByTestId('tb-more').tap();
    const panel = page.getByTestId('tb-overflow');
    const box = (await panel.boundingBox())!;
    expect(box.x).toBeGreaterThanOrEqual(0);
    expect(box.x + box.width).toBeLessThanOrEqual(width);
    await panel.getByTestId('tb-bold').tap();
    expect(await docEval(page, '(d) => d.styleAt(1, 1).font.bold')).toBe(true);
    expect(await docEval(page, '(d) => d.styleAt(3, 3).font.bold')).toBe(true);
    await expect(panel).toBeVisible();
    await page.getByTestId('tb-more').tap();
    await expect(panel).toHaveCount(0);
    await page.getByTestId('tb-size').tap();
    const list = page.getByTestId('tb-size-list');
    const sizes = (await list.boundingBox())!;
    expect(sizes.x).toBeGreaterThanOrEqual(0);
    expect(sizes.x + sizes.width).toBeLessThanOrEqual(width);
    await list.getByRole('option', { name: 'x-large', exact: true }).tap();
    expect(await docEval(page, '(d) => d.styleAt(3, 3).font.size')).toBe(18);
    expect(errors).toEqual([]);
  });

  test('inserts a chart from the overflow toolbar, below the data', async ({ page }) => {
    await openApp(page);
    await setData(page, SALES);
    await select(page, 'A1:D4');
    await page.getByTestId('tb-more').tap();
    await page.getByTestId('tb-chart').tap();
    await page.getByTestId('chart-type-pie').tap();
    await page.getByTestId('chart-insert').tap();
    const chart = page.getByTestId('object-chart');
    await expect(chart).toBeVisible();
    // The overflow panel closed and the chart sits under the data, inside the screen width.
    await expect(page.getByTestId('tb-overflow')).toHaveCount(0);
    expect(await docEval(page, '(d) => d.sheet.charts[0].from')).toMatchObject({ c: 0, r: 4 });
    const box = (await chart.boundingBox())!;
    expect(box.x + box.width).toBeLessThanOrEqual(page.viewportSize()!.width + 1);
  });

  test('menus open as bottom sheets; downloads the workbook; print preview fills the screen', async ({ page }) => {
    const errors = await openApp(page);
    await setData(page, SALES);
    await page.getByTestId('menu-file').tap();
    const sheet = page.locator('.sl-bottom-sheet');
    await expect(sheet).toBeVisible();
    // Download shows its own items in the sheet.
    await sheet.getByTestId('menu-download').tap();
    const download = page.waitForEvent('download');
    await sheet.getByTestId('download-xlsx').tap();
    const file = await download;
    expect(file.suggestedFilename()).toBe('Untitled spreadsheet.xlsx');
    const files = unzipSync(new Uint8Array(readFileSync((await file.path())!)));
    expect(strFromU8(files['xl/worksheets/sheet1.xml'])).toContain('<v>150</v>');
    await expect(sheet).toHaveCount(0);
    // Print preview.
    await page.getByTestId('menu-file').tap();
    await sheet.getByTestId('menu-print').tap();
    const dialog = page.getByTestId('print-dialog');
    await expect(dialog.getByTestId('print-page')).toHaveCount(1);
    const box = (await dialog.boundingBox())!;
    expect(box.width).toBeGreaterThan(page.viewportSize()!.width - 20);
    await dialog.getByRole('button', { name: 'Show options' }).tap();
    await expect(dialog.getByTestId('print-options')).toBeVisible();
    expect(errors).toEqual([]);
  });
});
