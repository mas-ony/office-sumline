import { readFileSync } from 'node:fs';
import { expect, test } from '@playwright/test';
import { strFromU8, unzipSync } from 'fflate';
import { SALES, clickCell, docEval, input, openApp, select, setData, shown, typeIn } from './helpers';

const KITCHEN = '/tests/fixtures/kitchen-sink.xlsx';

test.describe('desktop editor', () => {
  test('moves, reorders and deletes objects; shapes with text', async ({ page }) => {
    await openApp(page);
    await clickCell(page, 'B3');
    await page.getByTestId('tb-shape').click();
    const gallery = page.getByTestId('shapes-dialog');
    for (const group of ['Lines', 'Rectangles', 'Basic Forms', 'Block Arrows', 'Stars', 'Callouts', 'Equation Shapes', 'Flowchart']) await expect(gallery).toContainText(group);
    await gallery.getByTestId('shape-roundRect').click();
    const shape = page.getByTestId('object-shape');
    await expect(shape).toHaveCount(1);
    expect(await docEval(page, '(d) => d.sheet.shapes[0].from')).toMatchObject({ r: 2, c: 1 });
    await shape.dblclick();
    const settings = page.getByTestId('shape-settings');
    await expect(settings.getByTestId('shape-fill')).toHaveValue('#4F81BD');
    await settings.getByTestId('shape-text').fill('Hello');
    await settings.getByTestId('shape-save').click();
    await expect(shape).toContainText('Hello');

    // Drag to move.
    const box = (await shape.boundingBox())!;
    await page.mouse.move(box.x + box.width / 2, box.y + box.height / 2);
    await page.mouse.down();
    await page.mouse.move(box.x + box.width / 2 + 150, box.y + box.height / 2 + 60, { steps: 5 });
    await page.mouse.up();
    expect(await docEval(page, '(d) => d.sheet.shapes[0].from.c')).toBeGreaterThan(1);
    await page.keyboard.press('Control+z');
    expect(await docEval(page, '(d) => d.sheet.shapes[0].from.c')).toBe(1);

    // Object menu: delete.
    await shape.click({ button: 'right' });
    await expect(page.getByTestId('obj-front')).toBeVisible();
    await page.getByTestId('obj-delete').click();
    await expect(shape).toHaveCount(0);
  });

  test('opens a workbook with a chart and a picture and keeps the chart part unchanged', async ({ page }) => {
    const errors = await openApp(page, KITCHEN);
    await expect(page.getByTestId('tab-Sales')).toBeVisible();
    expect(await shown(page, 'D8')).toBe('$6,501.25');
    await expect(page.getByTestId('object-image')).toHaveCount(1);
    const chart = page.getByTestId('object-chart');
    await expect(chart).toHaveCount(1);
    await expect(chart).toContainText('Revenue by region');
    const download = page.waitForEvent('download');
    await page.keyboard.press('Control+s');
    const file = await (await download).path();
    const files = unzipSync(new Uint8Array(readFileSync(file!)));
    const original = unzipSync(new Uint8Array(readFileSync('tests/fixtures/kitchen-sink.xlsx')));
    expect(strFromU8(files['xl/charts/chart1.xml'])).toBe(strFromU8(original['xl/charts/chart1.xml']));
    expect(errors).toEqual([]);
  });

  test('data validation, filters, notes and freeze panes', async ({ page }) => {
    await openApp(page);
    await setData(page, SALES);
    // Filters on the data region.
    await clickCell(page, 'B2');
    await page.getByTestId('menu-data').click();
    await page.getByTestId('menu-filters').click();
    expect(await docEval(page, '(d) => !!d.sheet.autoFilter')).toBe(true);
    // Validation: a list.
    await clickCell(page, 'F2');
    await page.getByTestId('tb-validation').click();
    const dv = page.getByTestId('validations-dialog');
    await dv.getByTestId('dv-add').click();
    await dv.getByTestId('dv-type').selectOption('list');
    await dv.getByTestId('dv-value1').fill('Yes,No');
    await dv.getByTestId('dv-save').click();
    await page.getByTestId('modal-close').click();
    expect(await docEval(page, '(d) => d.sheet.validations.length')).toBe(1);
    // A note from the cell menu.
    await clickCell(page, 'A6', { button: 'right' });
    await page.getByTestId('ctx-add-notes').click();
    await page.getByTestId('note-input').fill('Check with finance');
    await page.getByTestId('note-save').click();
    expect(await docEval(page, '(d) => { const t = d.sheet.commentAt(5, 0); return t && { kind: t.kind, text: t.comments[0].text }; }')).toEqual({ kind: 'note', text: 'Check with finance' });
    // Freeze the first row from the View menu.
    await page.getByTestId('menu-view').click();
    await page.getByText('Freeze rows').hover();
    await page.getByTestId('freeze-rows-1').click();
    expect(await docEval(page, '(d) => d.sheet.freeze.rows')).toBe(1);
  });

  test('sheet tabs: insert, rename, hide and unhide', async ({ page }) => {
    const errors = await openApp(page);
    await page.getByTestId('tab-add').click();
    await expect(page.getByTestId('tab-Sheet2')).toBeVisible();
    await page.getByTestId('tab-Sheet2').dblclick();
    await page.getByTestId('tab-rename-input').fill('Budget');
    await page.getByTestId('tab-rename-input').press('Enter');
    await expect(page.getByTestId('tab-Budget')).toBeVisible();
    await page.getByTestId('tab-Budget').click({ button: 'right' });
    await page.getByTestId('tab-hide').click();
    await expect(page.getByTestId('tab-Budget')).toHaveCount(0);
    // Hidden worksheets are listed under Show worksheet.
    await page.getByTestId('tab-Sheet1').click({ button: 'right' });
    await page.getByRole('menuitem', { name: 'Show worksheet' }).hover();
    await page.getByRole('menuitem', { name: 'Budget' }).click();
    await expect(page.getByTestId('tab-Budget')).toBeVisible();
    expect(errors).toEqual([]);
  });

  test('start screen creates a workbook; typing values and formulas', async ({ page }) => {
    const errors: string[] = [];
    page.on('pageerror', (e) => errors.push(e.message));
    await page.goto('/');
    const start = page.getByTestId('start-screen');
    await expect(start).toContainText('Drop your Excel file here');
    await expect(start).toContainText('XLSX · XLS · XLSM · XLSB · CSV');
    await page.getByTestId('start-new').click();
    await expect(page.getByTestId('toolbar')).toBeVisible();
    // Row-by-row entry: Enter after Tabs returns to the first column.
    await clickCell(page, 'A1');
    for (const [i, v] of ['Item', 'Price', 'Qty'].entries()) {
      await page.keyboard.type(v);
      await page.keyboard.press(i < 2 ? 'Tab' : 'Enter');
    }
    for (const [i, v] of ['Pen', '1.5', '4'].entries()) {
      await page.keyboard.type(v);
      await page.keyboard.press(i < 2 ? 'Tab' : 'Enter');
    }
    expect(await shown(page, 'A2')).toBe('Pen');
    expect(await shown(page, 'C2')).toBe('4');
    await typeIn(page, 'D2', '=B2*C2');
    expect(await shown(page, 'D2')).toBe('6');
    // Suggestions while typing a function name.
    await clickCell(page, 'D3');
    await page.keyboard.type('=SU');
    await expect(page.getByTestId('formula-suggestions')).toBeVisible();
    await page.keyboard.type('M(D2,10)');
    await page.keyboard.press('Enter');
    expect(await input(page, 'D3')).toBe('=SUM(D2,10)');
    expect(await shown(page, 'D3')).toBe('16');
    expect(errors).toEqual([]);
  });

  test('menus, context menu and paste special', async ({ page }) => {
    const errors = await openApp(page);
    for (const name of ['file', 'edit', 'insert', 'format', 'data', 'view', 'help']) {
      await page.getByTestId(`menu-${name}`).click();
      await expect(page.getByRole('menu').first()).toBeVisible();
      await page.keyboard.press('Escape');
    }
    await setData(page, SALES);
    await typeIn(page, 'E2', '=B2+C2');
    await clickCell(page, 'E2', { button: 'right' });
    await expect(page.getByTestId('ctx-chart')).toBeVisible();
    await page.getByTestId('ctx-copy').click();
    // Values only: the result of the formula, not the formula.
    await clickCell(page, 'F2', { button: 'right' });
    await page.getByTestId('ctx-paste-special').hover();
    await page.getByTestId('paste-values').click();
    await expect.poll(() => input(page, 'F2')).toBe('270');
    expect(await input(page, 'E2')).toBe('=B2+C2');
    expect(errors).toEqual([]);
  });

  test('formats cells from the toolbar and the Format dialog', async ({ page }) => {
    const errors = await openApp(page);
    await setData(page, SALES);
    await select(page, 'B2:D4');
    await page.getByTestId('tb-bold').click();
    expect(await docEval(page, '(d) => d.styleAt(1, 1).font.bold')).toBe(true);
    expect(await docEval(page, '(d) => d.styleAt(3, 3).font.bold')).toBe(true);
    await page.getByTestId('tb-size').click();
    await page.getByRole('option', { name: 'x-large', exact: true }).click();
    expect(await docEval(page, '(d) => d.styleAt(3, 3).font.size')).toBe(18);
    // The Format dialog from the cell menu.
    await clickCell(page, 'B2', { button: 'right' });
    await page.getByTestId('ctx-format').click();
    const dialog = page.getByTestId('format-dialog');
    await dialog.getByRole('option', { name: 'Percentage' }).click();
    await dialog.getByTestId('format-ok').click();
    expect(await shown(page, 'B2')).toBe('12000.00%');
    await page.keyboard.press('Control+z');
    expect(await shown(page, 'B2')).toBe('120');
    expect(errors).toEqual([]);
  });

  test('inserts a chart, edits its settings and saves it in the XLSX file', async ({ page }) => {
    const errors = await openApp(page);
    await setData(page, SALES);
    await select(page, 'A1:D4');
    await page.getByTestId('tb-chart').click();
    const dialog = page.getByTestId('chart-insert-dialog');
    await expect(dialog.getByTestId('chart-range')).toHaveValue('A1:D4');
    await expect(dialog.locator('.sl-chart-type')).toHaveCount(23);
    await dialog.getByTestId('chart-type-column').click();
    await dialog.getByTestId('chart-insert').click();
    const chart = page.getByTestId('object-chart');
    await expect(chart).toHaveCount(1);
    await expect(chart.locator('svg rect')).not.toHaveCount(0);
    await expect(chart).toContainText('North');

    // Settings: line chart with a title.
    await chart.dblclick();
    const panel = page.getByTestId('chart-settings');
    await expect(panel.getByTestId('chart-type')).toHaveValue('column');
    await expect(panel.getByTestId('chart-series')).toHaveCount(3);
    await panel.getByTestId('chart-type').selectOption('line');
    await panel.getByTestId('chart-tab-style').click();
    await panel.getByTestId('chart-title').fill('Quarterly sales');
    await panel.getByTestId('chart-save').click();
    await expect(chart).toContainText('Quarterly sales');
    expect(await docEval(page, '(d) => d.sheet.charts[0].type')).toBe('line');

    // The chart follows the data.
    await page.evaluate(() => (window as unknown as { spreadsheetEditor: { setValue(r: string, v: number): void } }).spreadsheetEditor.setValue('B2', 999));
    await expect(chart).toContainText('999');

    // Saved as a chart part of the workbook.
    await page.getByTestId('menu-file').click();
    await page.getByTestId('menu-download').hover();
    const download = page.waitForEvent('download');
    await page.getByTestId('download-xlsx').click();
    const file = await (await download).path();
    const files = unzipSync(new Uint8Array(readFileSync(file!)));
    const xml = strFromU8(files['xl/charts/chart1.xml']);
    expect(xml).toContain('<c:lineChart>');
    expect(xml).toContain('Quarterly sales');
    expect(strFromU8(files['xl/drawings/drawing1.xml'])).toContain('<c:chart');
    expect(errors).toEqual([]);
  });

  test('sorts the data A to Z and Z to A', async ({ page }) => {
    const errors = await openApp(page);
    await setData(page, [['Region', 'Sales'], ['West', 5], ['East', 10], ['West', 15], ['North', 7]]);
    const rows = () => docEval(page, '(d) => [1,2,3,4].map((r) => d.displayText(r, 0) + d.displayText(r, 1))');
    // The column of the active cell sorts the data around it; the header row stays on top.
    await clickCell(page, 'A3');
    await page.getByTestId('menu-data').click();
    await page.getByText('Sorting', { exact: true }).hover();
    await page.getByTestId('sort-asc').click();
    expect(await rows()).toEqual(['East10', 'North7', 'West5', 'West15']);
    await clickCell(page, 'B3');
    await page.getByTestId('menu-data').click();
    await page.getByText('Sorting', { exact: true }).hover();
    await page.getByTestId('sort-desc').click();
    expect(await rows()).toEqual(['West15', 'East10', 'North7', 'West5']);
    expect(await shown(page, 'A1')).toBe('Region');
    expect(errors).toEqual([]);
  });

  test('downloads a CSV file; print preview and PDF download', async ({ page }) => {
    await openApp(page);
    await setData(page, SALES);
    await page.getByTestId('menu-file').click();
    await page.getByTestId('menu-download').hover();
    const csvDownload = page.waitForEvent('download');
    await page.getByTestId('download-csv').click();
    const csv = await csvDownload;
    expect(csv.suggestedFilename()).toBe('Untitled spreadsheet.csv');
    expect(readFileSync((await csv.path())!, 'utf8')).toContain('North,120,150,90');
    // Print: a live preview and a PDF.
    await page.keyboard.press('Control+p');
    const dialog = page.getByTestId('print-dialog');
    await expect(dialog.getByTestId('print-page')).toHaveCount(1);
    await expect(dialog.getByTestId('print-page').locator('svg text').first()).toBeVisible();
    for (const key of ['scope', 'paper', 'orientation', 'scale', 'margins', 'order', 'hAlign', 'vAlign']) await expect(dialog.getByTestId(`print-${key}`)).toBeVisible();
    await expect(dialog).toContainText('Repeat frozen rows');
    // Headings change the layout.
    await dialog.getByTestId('print-headings').check();
    await expect(dialog.getByTestId('print-page').locator('svg').first()).toContainText('A');
    const download = page.waitForEvent('download');
    await dialog.getByTestId('print-download').click();
    const file = await (await download).path();
    const pdf = readFileSync(file!);
    expect(pdf.subarray(0, 8).toString('latin1')).toBe('%PDF-1.4');
    expect(pdf.toString('latin1')).toContain('/Type /Page');
  });

  test('the page API opens a workbook and reads and writes cells', async ({ page }) => {
    const errors = await openApp(page);
    const bytes = [...readFileSync('tests/fixtures/kitchen-sink.xlsx')];
    await page.evaluate(async (data) => {
      const ed = (window as unknown as { spreadsheetEditor: { open(f: Blob, name?: string): Promise<void> } }).spreadsheetEditor;
      await ed.open(new Blob([new Uint8Array(data)]), 'kitchen-sink.xlsx');
    }, bytes);
    await expect(page.getByTestId('tab-Sales')).toBeVisible();
    expect(await shown(page, 'D8')).toBe('$6,501.25');
    // North sells 10 more units: a change event, and the total follows.
    const r = await page.evaluate(async () => {
      type Api = {
        getValue(ref: string): unknown;
        getText(ref: string): string;
        setValue(ref: string, v: unknown): void;
        select(ref: string): void;
        getSelected(): string;
        on(event: 'change', fn: () => void): unknown;
      };
      const ed = (window as unknown as { spreadsheetEditor: Api }).spreadsheetEditor;
      let changes = 0;
      ed.on('change', () => changes++);
      const before = ed.getValue('D8') as number;
      const price = ed.getValue('C3') as number;
      ed.setValue('B3', (ed.getValue('B3') as number) + 10);
      ed.select('B3');
      await new Promise((done) => setTimeout(done, 50));
      return { before, price, after: ed.getValue('D8') as number, text: ed.getText('D8'), changes, selected: ed.getSelected() };
    });
    expect(r.after).toBeCloseTo(r.before + 10 * r.price, 6);
    expect(r.changes).toBe(1);
    expect(r.selected).toBe('B3');
    expect(r.text).not.toBe('$6,501.25');
    expect(await shown(page, 'D8')).toBe(r.text);
    expect(errors).toEqual([]);
  });
});
