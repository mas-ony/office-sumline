/**
 * Editor commands shared by the top menu, the toolbar, the context menus and keyboard shortcuts.
 */
import { colName, isFullColumns, isFullRows, makeRange, rangesIntersect, type Range } from '../core/refs';
import { filterRegion } from '../core/filter';
import type { StylePatch } from '../core/model/style';
import type { UIContextValue } from './context';
import { copySelection, pasteClipboard, type PasteMode } from './clipboardActions';

type Ctx = UIContextValue;

function commitEditing(ctx: Ctx): void {
  if (ctx.grid?.isEditing()) ctx.grid.commitEdit();
}

export function undo(ctx: Ctx): void {
  if (ctx.grid?.isEditing()) {
    ctx.grid.cancelEdit();
    return;
  }
  ctx.doc.undo();
}

export function redo(ctx: Ctx): void {
  commitEditing(ctx);
  ctx.doc.redo();
}

export function cut(ctx: Ctx): void {
  copySelection(ctx.doc, ctx.grid, true);
}

export function copy(ctx: Ctx): void {
  copySelection(ctx.doc, ctx.grid, false);
}

export async function paste(ctx: Ctx, mode: PasteMode = 'all'): Promise<void> {
  commitEditing(ctx);
  const ok = await pasteClipboard(ctx.doc, mode);
  if (!ok) ctx.notify(window.isSecureContext ? 'Allow clipboard access, or press Ctrl+V to paste.' : 'Press Ctrl+V to paste from other apps (this page cannot read the clipboard directly).');
  ctx.grid?.focusIdle();
}

export function format(ctx: Ctx, patch: StylePatch, label = 'Format'): void {
  commitEditing(ctx);
  ctx.doc.format(patch, undefined, undefined, label);
}

export function toggleFont(ctx: Ctx, prop: 'bold' | 'italic' | 'underline' | 'strike'): void {
  commitEditing(ctx);
  ctx.doc.toggleFont(prop);
}

/** Format > Text > Bold/Italic/... (sets rather than toggles). */
export function setFontFlag(ctx: Ctx, prop: 'bold' | 'italic' | 'underline' | 'strike'): void {
  format(ctx, { font: { [prop]: prop === 'underline' ? 'single' : true } }, prop);
}

export function setNumberFormat(ctx: Ctx, code: string): void {
  commitEditing(ctx);
  ctx.doc.setNumberFormat(code);
}

/** Text rotation in degrees (positive = upwards), stored as the XLSX textRotation. */
export function rotate(ctx: Ctx, degrees: number): void {
  const rotation = degrees >= 0 ? degrees : 90 - degrees;
  format(ctx, { align: { rotation: rotation || null } }, 'Rotate text');
}

export function removeFormat(ctx: Ctx): void {
  commitEditing(ctx);
  ctx.doc.clear('formats');
}

export async function merge(ctx: Ctx): Promise<void> {
  commitEditing(ctx);
  const { doc } = ctx;
  const sheet = doc.sheet;
  const ranges = doc.selection.ranges;
  const range = doc.range;
  if (ctx.state.format.merged) {
    doc.unmerge();
    return;
  }
  if (range.r0 === range.r1 && range.c0 === range.c1) return;
  if (sheet.merges.some((m) => ranges.some((r) => rangesIntersect(m, r)))) {
    if (!(await ctx.confirm('This action will destroy any existing merged cells. Are you sure?'))) return;
  }
  let filled = 0;
  sheet.forEachCellIn(range, (_r, _c, cell) => {
    if (cell.v !== null && cell.v !== '') filled++;
  });
  if (filled > 1 && !(await ctx.confirm('The merged cells will retain the value of the top-left cell only. Are you sure?'))) return;
  doc.merge();
}

export function insertRows(ctx: Ctx, where: 'above' | 'below'): void {
  commitEditing(ctx);
  ctx.doc.insertRowsAtSelection(where);
}

export function insertCols(ctx: Ctx, where: 'left' | 'right'): void {
  commitEditing(ctx);
  ctx.doc.insertColsAtSelection(where);
}

export async function deleteRows(ctx: Ctx): Promise<void> {
  commitEditing(ctx);
  const range = ctx.doc.range;
  if (isFullColumns(range)) {
    ctx.notify('Select rows to delete, not whole columns.');
    return;
  }
  ctx.doc.deleteSelectedRows();
}

export async function deleteCols(ctx: Ctx): Promise<void> {
  commitEditing(ctx);
  const range = ctx.doc.range;
  if (isFullRows(range)) {
    ctx.notify('Select columns to delete, not whole rows.');
    return;
  }
  ctx.doc.deleteSelectedCols();
}

/** Appends a row after the last used row (column menu "Create a new row"). */
export function appendRow(ctx: Ctx): void {
  const used = ctx.doc.sheet.usedRange();
  const at = used ? used.r1 + 1 : 0;
  ctx.doc.insertRows(at, 1);
  ctx.doc.selectCell(at, ctx.doc.selection.active.c);
}

/** Whether the first row of a block looks like a header (text above numbers or dates). */
function hasHeader(ctx: Ctx, region: Range, col: number): boolean {
  if (ctx.doc.sheet.autoFilter && rangesIntersect(ctx.doc.sheet.autoFilter.ref, region)) return true;
  if (region.r1 <= region.r0) return false;
  const first = ctx.doc.cell(region.r0, col)?.v;
  const second = ctx.doc.cell(region.r0 + 1, col)?.v;
  if (typeof first === 'string' && (typeof second === 'number' || typeof second === 'boolean')) return true;
  // Every first-row cell is text while some column below holds numbers.
  let allText = true;
  for (let c = region.c0; c <= region.c1; c++) {
    const v = ctx.doc.cell(region.r0, c)?.v;
    if (v !== null && v !== undefined && typeof v !== 'string') allText = false;
  }
  if (!allText) return false;
  for (let c = region.c0; c <= region.c1; c++) if (typeof ctx.doc.cell(region.r0 + 1, c)?.v === 'number') return true;
  return false;
}

/** Sorts the data block around the selection by the active column. */
export function sortColumn(ctx: Ctx, ascending: boolean): void {
  commitEditing(ctx);
  const { doc } = ctx;
  const sel = doc.selection;
  const col = sel.active.c;
  const range = doc.range;
  const region = isFullColumns(range) || (range.r0 === range.r1 && range.c0 === range.c1) ? filterRegion(doc.sheet, makeRange(sel.active.r, col), { r: Math.max(0, range.r0 === 0 && isFullColumns(range) ? 0 : sel.active.r), c: col }) : range;
  const used = doc.sheet.usedRange(false);
  const area = region ?? (used ? { ...used } : null);
  if (!area) return;
  doc.sortRange(area, col, ascending, hasHeader(ctx, area, col));
}

export function toggleFilters(ctx: Ctx): void {
  commitEditing(ctx);
  const on = ctx.doc.toggleAutoFilter();
  if (!on && !ctx.doc.sheet.autoFilter) {
    if (!ctx.state.hasFilter) ctx.notify('Select a cell inside your data first.');
  }
}

export function freezeRows(ctx: Ctx, n: number): void {
  ctx.doc.setFreeze(n, ctx.doc.sheet.freeze.cols);
}

export function freezeCols(ctx: Ctx, n: number): void {
  ctx.doc.setFreeze(ctx.doc.sheet.freeze.rows, n);
}

/** View > Toggle footer row: totals of every used column under the grid. */
export function toggleFooter(ctx: Ctx): void {
  const sheet = ctx.doc.sheet;
  if (sheet.footer) {
    ctx.doc.setFooter(null);
    return;
  }
  const used = sheet.usedRange(false);
  const cols = Math.max(used ? used.c1 + 1 : 0, 13);
  const row = ['Total'];
  for (let c = 1; c < cols; c++) row.push(`=SUM(${colName(c)}:${colName(c)})`);
  ctx.doc.setFooter(row);
}

export function addComment(ctx: Ctx): void {
  const { r, c } = ctx.doc.selection.active;
  const m = ctx.doc.sheet.mergeAt(r, c);
  ctx.openComments(m?.r0 ?? r, m?.c0 ?? c);
}

export function editNote(ctx: Ctx): void {
  const { r, c } = ctx.doc.selection.active;
  const m = ctx.doc.sheet.mergeAt(r, c);
  ctx.openNote(m?.r0 ?? r, m?.c0 ?? c);
}

export function clearNote(ctx: Ctx): void {
  const { r, c } = ctx.doc.selection.active;
  ctx.doc.setNote(r, c, '');
}

export function newWorksheet(ctx: Ctx): void {
  commitEditing(ctx);
  ctx.doc.addSheet(undefined, ctx.doc.activeIndex + 1);
}

export async function renameWorksheet(ctx: Ctx, index = ctx.doc.activeIndex): Promise<void> {
  const sheet = ctx.doc.wb.sheets[index];
  if (!sheet) return;
  let value = sheet.name;
  for (;;) {
    const next = await ctx.prompt('Rename this worksheet', value, { title: 'Rename worksheet', ok: 'Rename' });
    if (next === null || next === sheet.name) return;
    const error = ctx.doc.renameSheet(index, next);
    if (!error) return;
    ctx.notify(error);
    value = next;
  }
}

export async function deleteWorksheet(ctx: Ctx, index = ctx.doc.activeIndex): Promise<void> {
  if (ctx.doc.wb.visibleSheets().length <= 1) {
    ctx.notify('You cannot delete the last worksheet.');
    return;
  }
  const name = ctx.doc.wb.sheets[index]?.name;
  if (!(await ctx.confirm(`Are you sure you want to delete the worksheet “${name}”?`, { ok: 'Delete', danger: true }))) return;
  ctx.doc.removeSheet(index);
}

/** Selects every cell (Ctrl+A). */
export function selectAll(ctx: Ctx): void {
  ctx.doc.selectAll();
}

/** Toolbar image upload: picks an image file and inserts it at the selection. */
export function uploadImage(ctx: Ctx): void {
  const input = document.createElement('input');
  input.type = 'file';
  input.accept = '.jpg,.jpeg,.png,.gif,.webp,.bmp,.svg,image/*';
  input.onchange = () => {
    const file = input.files?.[0];
    if (file) void ctx.grid?.insertImageFile(file, file.name);
  };
  input.click();
}
