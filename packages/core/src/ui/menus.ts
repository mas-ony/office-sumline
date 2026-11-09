/**
 * Menu contents: the top menu (File, Edit, Insert, Format, Data, View, Help) and the context menus
 * of cells, column and row headers, worksheet tabs and floating objects.
 */
import type { UIContextValue } from './context';
import type { MenuItem } from './menuItems';
import * as cmd from './commands';
import { FORMAT_PRESETS, MENU_FONT_SIZES, ZOOM_LEVELS } from './docState';

type Ctx = UIContextValue;

const isMac = typeof navigator !== 'undefined' && /Mac|iPhone|iPad/.test(navigator.platform || navigator.userAgent);
export const MOD = isMac ? '⌘' : 'Ctrl';
const key = (k: string) => `${MOD}+${k}`;

function pasteSpecial(ctx: Ctx): MenuItem[] {
  return [
    { title: 'Full paste', onClick: () => void cmd.paste(ctx, 'all'), testId: 'paste-full' },
    { type: 'line' },
    { title: 'Paste values', onClick: () => void cmd.paste(ctx, 'values'), testId: 'paste-values' },
    { title: 'Paste values and formulas', onClick: () => void cmd.paste(ctx, 'formulas'), testId: 'paste-formulas' },
    { title: 'Paste style', onClick: () => void cmd.paste(ctx, 'formats'), testId: 'paste-style' },
  ];
}

export function topMenus(ctx: Ctx): { title: string; items: () => MenuItem[] }[] {
  const ro = ctx.state.readOnly;
  return [
    {
      title: 'File',
      items: () => [
        { title: 'New Spreadsheet', icon: 'add_box', onClick: ctx.files.newWorkbook, testId: 'menu-new' },
        { title: 'Open', icon: 'folder', onClick: ctx.files.open, testId: 'menu-open' },
        { type: 'line' },
        { title: 'Share', icon: 'share', disabled: !ctx.files.share, onClick: ctx.files.share, testId: 'menu-share' },
        {
          title: 'Download',
          icon: 'file_download',
          testId: 'menu-download',
          submenu: [
            { title: 'XLSX', onClick: () => ctx.files.download('xlsx'), testId: 'download-xlsx' },
            { title: 'CSV', onClick: () => ctx.files.download('csv'), testId: 'download-csv' },
            { title: 'TSV', onClick: () => ctx.files.download('tsv'), testId: 'download-tsv' },
            { title: 'PDF', onClick: () => ctx.files.download('pdf'), testId: 'download-pdf' },
          ],
        },
        { title: 'Print', icon: 'print', shortcut: key('P'), onClick: ctx.files.print, testId: 'menu-print' },
        { type: 'line' },
        {
          title: 'Rename',
          icon: 'drive_file_rename_outline',
          testId: 'menu-rename',
          onClick: async () => {
            const base = ctx.doc.fileName.replace(/\.[^.]+$/, '');
            const name = await ctx.prompt('File name', base, { title: 'Rename', ok: 'Rename' });
            if (name && name.trim()) ctx.files.rename(name.trim());
          },
        },
        { title: 'Move to trash', icon: 'delete', disabled: true },
        { type: 'line' },
        { title: 'History', icon: 'history', disabled: true },
      ],
    },
    {
      title: 'Edit',
      items: () => [
        { title: 'Undo', icon: 'undo', shortcut: key('Z'), disabled: ro || !ctx.state.canUndo, onClick: () => cmd.undo(ctx), testId: 'menu-undo' },
        { title: 'Redo', icon: 'redo', shortcut: key('Y'), disabled: ro || !ctx.state.canRedo, onClick: () => cmd.redo(ctx), testId: 'menu-redo' },
        { type: 'line' },
        { title: 'Cut', icon: 'content_cut', shortcut: key('X'), disabled: ro, onClick: () => cmd.cut(ctx) },
        { title: 'Copy', icon: 'content_copy', shortcut: key('C'), onClick: () => cmd.copy(ctx) },
        { title: 'Paste', icon: 'content_paste', shortcut: key('V'), disabled: ro, onClick: () => void cmd.paste(ctx) },
        { type: 'line' },
        {
          title: 'Delete',
          icon: 'delete',
          disabled: ro,
          submenu: [
            { title: 'Selected rows', onClick: () => void cmd.deleteRows(ctx), testId: 'menu-delete-rows' },
            { title: 'Selected columns', onClick: () => void cmd.deleteCols(ctx), testId: 'menu-delete-cols' },
          ],
        },
        { title: 'Find and replace', icon: 'find_replace', shortcut: key('F'), onClick: () => ctx.openDialog({ kind: 'search', replace: true }), testId: 'menu-find' },
      ],
    },
    {
      title: 'Insert',
      items: () => [
        {
          title: 'Rows',
          icon: 'table_rows',
          disabled: ro,
          submenu: [
            { title: 'Add a new line before', onClick: () => cmd.insertRows(ctx, 'above'), testId: 'menu-row-before' },
            { title: 'Add a new line after', onClick: () => cmd.insertRows(ctx, 'below'), testId: 'menu-row-after' },
          ],
        },
        {
          title: 'Columns',
          icon: 'view_column',
          disabled: ro,
          submenu: [
            { title: 'Add a new column before', onClick: () => cmd.insertCols(ctx, 'left'), testId: 'menu-col-before' },
            { title: 'Add a new column after', onClick: () => cmd.insertCols(ctx, 'right'), testId: 'menu-col-after' },
          ],
        },
        { title: 'Create a new worksheet', disabled: ro, onClick: () => cmd.newWorksheet(ctx), testId: 'menu-new-sheet' },
        { type: 'line' },
        { title: 'Add chart', icon: 'addchart', disabled: ro, onClick: () => ctx.openDialog({ kind: 'chart-insert' }), testId: 'menu-chart' },
        { type: 'line' },
        { title: 'Link', icon: 'link', disabled: ro, onClick: () => ctx.openDialog({ kind: 'link' }), testId: 'menu-link' },
        { title: 'Checkbox', icon: 'check_box', disabled: true },
        { title: 'Comments', icon: 'comment', disabled: ro, onClick: () => cmd.addComment(ctx), testId: 'menu-comments' },
      ],
    },
    {
      title: 'Format',
      items: () => [
        {
          title: 'Format',
          icon: '123',
          disabled: ro,
          submenu: [
            ...FORMAT_PRESETS.slice(0, 4).map((p) => ({ title: p.label, shortcut: p.sample || undefined, checked: ctx.state.format.numFmt === p.code, onClick: () => cmd.setNumberFormat(ctx, p.code), testId: `fmt-${p.label.toLowerCase().replace(/\s+/g, '-')}` })),
            { type: 'line' },
            ...FORMAT_PRESETS.slice(4).map((p) => ({ title: p.label, shortcut: p.sample, checked: ctx.state.format.numFmt === p.code, onClick: () => cmd.setNumberFormat(ctx, p.code), testId: `fmt-${p.label.toLowerCase().replace(/\s+/g, '-')}` })),
            { type: 'line' },
            { title: 'Custom', onClick: () => ctx.openDialog({ kind: 'format' }), testId: 'fmt-custom' },
          ] as MenuItem[],
        },
        {
          title: 'Text',
          icon: 'format_bold',
          disabled: ro,
          submenu: [
            { title: 'Bold', checked: ctx.state.format.bold, onClick: () => cmd.toggleFont(ctx, 'bold'), testId: 'text-bold' },
            { title: 'Italic', checked: ctx.state.format.italic, onClick: () => cmd.toggleFont(ctx, 'italic'), testId: 'text-italic' },
            { title: 'Underline', checked: ctx.state.format.underline, onClick: () => cmd.toggleFont(ctx, 'underline'), testId: 'text-underline' },
            { title: 'Line through', checked: ctx.state.format.strike, onClick: () => cmd.toggleFont(ctx, 'strike'), testId: 'text-strike' },
          ],
        },
        {
          title: 'Rotate',
          disabled: ro,
          submenu: [
            { title: '+90deg', icon: 'text_rotate_up', onClick: () => cmd.rotate(ctx, 90), testId: 'rotate-90' },
            { title: '+45deg', icon: 'text_rotation_angleup', onClick: () => cmd.rotate(ctx, 45), testId: 'rotate-45' },
            { title: '0deg', icon: 'text_rotation_none', onClick: () => cmd.rotate(ctx, 0), testId: 'rotate-0' },
            { title: '-45deg', icon: 'text_rotation_angledown', onClick: () => cmd.rotate(ctx, -45), testId: 'rotate-m45' },
            { title: '-90deg', icon: 'text_rotation_down', onClick: () => cmd.rotate(ctx, -90), testId: 'rotate-m90' },
          ],
        },
        { type: 'line' },
        {
          title: 'Font size',
          icon: 'format_size',
          disabled: ro,
          submenu: MENU_FONT_SIZES.map((sz) => ({ title: String(sz), checked: ctx.state.format.fontSize === sz, onClick: () => cmd.format(ctx, { font: { size: sz } }, 'Font size'), testId: `size-${sz}` })),
        },
        { type: 'line' },
        { title: 'Alternate colors', icon: 'opacity', disabled: true },
        { title: 'Remove all format', icon: 'format_clear', shortcut: `${MOD}+\\`, disabled: ro, onClick: () => cmd.removeFormat(ctx), testId: 'menu-clear-format' },
      ],
    },
    {
      title: 'Data',
      items: () => [
        {
          title: 'Sorting',
          disabled: ro,
          submenu: [
            { title: 'Sorting column from A to Z', onClick: () => cmd.sortColumn(ctx, true), testId: 'sort-asc' },
            { title: 'Sorting column from Z to A', onClick: () => cmd.sortColumn(ctx, false), testId: 'sort-desc' },
          ],
        },
        { title: 'Toggle filters', icon: 'filter_alt', checked: ctx.state.hasFilter, disabled: ro, onClick: () => cmd.toggleFilters(ctx), testId: 'menu-filters' },
        { title: 'Data validations', icon: 'grading', disabled: ro, onClick: () => ctx.openDialog({ kind: 'validations' }), testId: 'menu-validations' },
      ],
    },
    {
      title: 'View',
      items: () => [
        {
          title: 'Freeze rows',
          icon: 'table_rows',
          submenu: [
            { title: 'No freeze', checked: ctx.state.freezeRows === 0, onClick: () => cmd.freezeRows(ctx, 0), testId: 'freeze-rows-0' },
            ...[1, 2, 3].map((n) => ({ title: `Freeze ${n}`, checked: ctx.state.freezeRows === n, onClick: () => cmd.freezeRows(ctx, n), testId: `freeze-rows-${n}` })),
          ],
        },
        {
          title: 'Freeze columns',
          icon: 'view_column',
          submenu: [
            { title: 'No freeze', checked: ctx.state.freezeCols === 0, onClick: () => cmd.freezeCols(ctx, 0), testId: 'freeze-cols-0' },
            ...[1, 2, 3].map((n) => ({ title: `Freeze ${n}`, checked: ctx.state.freezeCols === n, onClick: () => cmd.freezeCols(ctx, n), testId: `freeze-cols-${n}` })),
          ],
        },
        { type: 'line' },
        { title: 'Toggle footer row', icon: 'table_chart', checked: ctx.state.hasFooter, onClick: () => cmd.toggleFooter(ctx), testId: 'menu-footer' },
        { type: 'line' },
        { title: 'Gridlines', checked: ctx.doc.sheet.showGridLines, onClick: () => ctx.doc.setGridLines(!ctx.doc.sheet.showGridLines), testId: 'menu-gridlines' },
        {
          title: 'Zoom',
          submenu: ZOOM_LEVELS.map((z) => ({ title: `${Math.round(z * 100)}%`, checked: Math.abs(ctx.doc.sheet.zoom - z) < 0.001, onClick: () => ctx.grid?.setZoom(z) })),
        },
      ],
    },
    {
      title: 'Help',
      items: () => [
        { title: 'Keyboard shortcuts', icon: 'article', onClick: () => ctx.openDialog({ kind: 'shortcuts' }), testId: 'menu-shortcuts' },
        { title: 'About', icon: 'info', onClick: () => ctx.openDialog({ kind: 'about' }), testId: 'menu-about' },
        { type: 'line' },
        { title: 'Function list', icon: 'functions', onClick: () => ctx.openDialog({ kind: 'functions' }), testId: 'menu-functions' },
      ],
    },
  ];
}

export function cellMenu(ctx: Ctx): MenuItem[] {
  const ro = ctx.state.readOnly;
  const note = ctx.state.format.hasComment && ctx.state.format.commentKind === 'note';
  return [
    { title: 'Format cells', disabled: ro, onClick: () => ctx.openDialog({ kind: 'format' }), testId: 'ctx-format' },
    { type: 'line' },
    { title: 'Cut', icon: 'content_cut', shortcut: key('X'), disabled: ro, onClick: () => cmd.cut(ctx), testId: 'ctx-cut' },
    { title: 'Copy', icon: 'content_copy', shortcut: key('C'), onClick: () => cmd.copy(ctx), testId: 'ctx-copy' },
    { title: 'Paste', icon: 'content_paste', shortcut: key('V'), disabled: ro, onClick: () => void cmd.paste(ctx), testId: 'ctx-paste' },
    { title: 'Paste Special', icon: 'content_paste', disabled: ro, submenu: pasteSpecial(ctx), testId: 'ctx-paste-special' },
    { type: 'line' },
    { title: 'Insert chart', icon: 'addchart', disabled: ro, onClick: () => ctx.openDialog({ kind: 'chart-insert' }), testId: 'ctx-chart' },
    { type: 'line' },
    { title: 'Insert a new row before', disabled: ro, onClick: () => cmd.insertRows(ctx, 'above'), testId: 'ctx-row-before' },
    { title: 'Insert a new row after', disabled: ro, onClick: () => cmd.insertRows(ctx, 'below'), testId: 'ctx-row-after' },
    { title: 'Delete selected rows', disabled: ro, onClick: () => void cmd.deleteRows(ctx), testId: 'ctx-delete-rows' },
    { type: 'line' },
    { title: 'Comments', icon: 'rate_review', disabled: ro && !ctx.state.format.hasComment, onClick: () => cmd.addComment(ctx), testId: 'ctx-comments' },
    ...(note
      ? [
          { title: 'Edit notes', icon: 'notes', disabled: ro, onClick: () => cmd.editNote(ctx), testId: 'ctx-edit-notes' } as MenuItem,
          { title: 'Clear notes', disabled: ro, onClick: () => cmd.clearNote(ctx), testId: 'ctx-clear-notes' } as MenuItem,
        ]
      : [{ title: 'Add notes', icon: 'notes', disabled: ro || ctx.state.format.hasComment, onClick: () => cmd.editNote(ctx), testId: 'ctx-add-notes' } as MenuItem]),
    { type: 'line' },
    { title: 'Save as', icon: 'save', shortcut: key('S'), onClick: () => ctx.files.download('xlsx'), testId: 'ctx-save-as' },
    { title: 'About', icon: 'info', onClick: () => ctx.openDialog({ kind: 'about' }) },
  ];
}

export function columnMenu(ctx: Ctx): MenuItem[] {
  const ro = ctx.state.readOnly;
  const range = ctx.doc.range;
  const anyHidden = (() => {
    for (let c = Math.max(0, range.c0 - 1); c <= Math.min(range.c1 + 1, 16383); c++) if (ctx.doc.sheet.isColHidden(c)) return true;
    return false;
  })();
  return [
    { title: 'Format column', disabled: ro, onClick: () => ctx.openDialog({ kind: 'format' }), testId: 'col-format' },
    { type: 'line' },
    { title: 'Cut', icon: 'content_cut', shortcut: key('X'), disabled: ro, onClick: () => cmd.cut(ctx) },
    { title: 'Copy', icon: 'content_copy', shortcut: key('C'), onClick: () => cmd.copy(ctx) },
    { title: 'Paste', icon: 'content_paste', shortcut: key('V'), disabled: ro, onClick: () => void cmd.paste(ctx) },
    { title: 'Paste Special', icon: 'content_paste', disabled: ro, submenu: pasteSpecial(ctx) },
    { type: 'line' },
    { title: 'Insert a new column before', disabled: ro, onClick: () => cmd.insertCols(ctx, 'left'), testId: 'col-insert-before' },
    { title: 'Insert a new column after', disabled: ro, onClick: () => cmd.insertCols(ctx, 'right'), testId: 'col-insert-after' },
    { title: 'Delete selected columns', disabled: ro, onClick: () => void cmd.deleteCols(ctx), testId: 'col-delete' },
    { title: 'Rename this column', disabled: ro, onClick: () => ctx.openDialog({ kind: 'column-title', col: ctx.doc.selection.active.c }), testId: 'col-rename' },
    { type: 'line' },
    { title: 'Create a new row', disabled: ro, onClick: () => cmd.appendRow(ctx), testId: 'col-new-row' },
    { type: 'line' },
    { title: 'Order ascending', disabled: ro, onClick: () => cmd.sortColumn(ctx, true), testId: 'col-sort-asc' },
    { title: 'Order descending', disabled: ro, onClick: () => cmd.sortColumn(ctx, false), testId: 'col-sort-desc' },
    { type: 'line' },
    { title: 'Hide', disabled: ro, onClick: () => ctx.doc.setHidden('col', rangeIndexes(range.c0, range.c1), true), testId: 'col-hide' },
    { title: 'Show', disabled: ro || !anyHidden, onClick: () => ctx.doc.setHidden('col', rangeIndexes(Math.max(0, range.c0 - 1), range.c1 + 1), false), testId: 'col-show' },
    { type: 'line' },
    { title: 'Save as', icon: 'save', shortcut: key('S'), onClick: () => ctx.files.download('xlsx') },
    { title: 'About', icon: 'info', onClick: () => ctx.openDialog({ kind: 'about' }) },
  ];
}

export function rowMenu(ctx: Ctx): MenuItem[] {
  const ro = ctx.state.readOnly;
  const range = ctx.doc.range;
  const anyHidden = (() => {
    for (let r = Math.max(0, range.r0 - 1); r <= range.r1 + 1; r++) if (ctx.doc.sheet.isRowHidden(r)) return true;
    return false;
  })();
  return [
    { title: 'Cut', icon: 'content_cut', shortcut: key('X'), disabled: ro, onClick: () => cmd.cut(ctx) },
    { title: 'Copy', icon: 'content_copy', shortcut: key('C'), onClick: () => cmd.copy(ctx) },
    { title: 'Paste', icon: 'content_paste', shortcut: key('V'), disabled: ro, onClick: () => void cmd.paste(ctx) },
    { title: 'Paste Special', icon: 'content_paste', disabled: ro, submenu: pasteSpecial(ctx) },
    { type: 'line' },
    { title: 'Insert a new row before', disabled: ro, onClick: () => cmd.insertRows(ctx, 'above'), testId: 'row-insert-before' },
    { title: 'Insert a new row after', disabled: ro, onClick: () => cmd.insertRows(ctx, 'below'), testId: 'row-insert-after' },
    { title: 'Delete selected rows', disabled: ro, onClick: () => void cmd.deleteRows(ctx), testId: 'row-delete' },
    { type: 'line' },
    { title: 'Hide', disabled: ro, onClick: () => ctx.doc.setHidden('row', rangeIndexes(range.r0, Math.min(range.r1, range.r0 + 5000)), true), testId: 'row-hide' },
    { title: 'Show', disabled: ro || !anyHidden, onClick: () => ctx.doc.setHidden('row', rangeIndexes(Math.max(0, range.r0 - 1), Math.min(range.r1 + 1, range.r0 + 5000)), false), testId: 'row-show' },
    { type: 'line' },
    { title: 'Save as', icon: 'save', shortcut: key('S'), onClick: () => ctx.files.download('xlsx') },
    { title: 'About', icon: 'info', onClick: () => ctx.openDialog({ kind: 'about' }) },
  ];
}

export function tabMenu(ctx: Ctx, index: number): MenuItem[] {
  const ro = ctx.state.readOnly;
  const hidden = ctx.doc.wb.sheets.filter((s) => s.state === 'hidden');
  return [
    { title: 'Rename this worksheet', disabled: ro, onClick: () => void cmd.renameWorksheet(ctx, index), testId: 'tab-rename' },
    { title: 'Delete this worksheet', disabled: ro, onClick: () => void cmd.deleteWorksheet(ctx, index), testId: 'tab-delete' },
    { title: 'Duplicate this worksheet', disabled: ro, onClick: () => ctx.doc.duplicateSheet(index), testId: 'tab-duplicate' },
    { title: 'Hide this worksheet', disabled: ro || ctx.doc.wb.visibleSheets().length <= 1, onClick: () => ctx.doc.setSheetHidden(index, true), testId: 'tab-hide' },
    ...(hidden.length
      ? [{ title: 'Show worksheet', disabled: ro, submenu: hidden.map((s) => ({ title: s.name, onClick: () => ctx.doc.setSheetHidden(ctx.doc.wb.sheets.indexOf(s), false) })) } as MenuItem]
      : []),
    { type: 'line' },
    { title: 'Save as', icon: 'save', shortcut: key('S'), onClick: () => ctx.files.download('xlsx') },
    { title: 'About', icon: 'info', onClick: () => ctx.openDialog({ kind: 'about' }) },
  ];
}

export function objectMenu(ctx: Ctx, id: string): MenuItem[] {
  const ro = ctx.state.readOnly;
  const hit = ctx.doc.objectById(id);
  return [
    { title: 'Bring to front', icon: 'flip_to_front', disabled: ro, onClick: () => ctx.doc.reorderObject(id, 'front'), testId: 'obj-front' },
    { title: 'Bring to back', icon: 'flip_to_back', disabled: ro, onClick: () => ctx.doc.reorderObject(id, 'back'), testId: 'obj-back' },
    { type: 'line' },
    { title: 'Flip to front', icon: 'flip_to_front', disabled: ro, onClick: () => ctx.doc.reorderObject(id, 'forward'), testId: 'obj-forward' },
    { title: 'Flip to back', icon: 'flip_to_back', disabled: ro, onClick: () => ctx.doc.reorderObject(id, 'backward'), testId: 'obj-backward' },
    { type: 'line' },
    {
      title: 'Edit',
      icon: 'edit',
      disabled: ro || hit?.kind === 'image',
      onClick: () => (hit?.kind === 'chart' ? ctx.openDialog({ kind: 'chart-settings', id }) : hit?.kind === 'shape' ? ctx.openDialog({ kind: 'shape-settings', id }) : undefined),
      testId: 'obj-edit',
    },
    {
      title: 'Delete',
      icon: 'delete',
      disabled: ro,
      onClick: () => {
        ctx.doc.deleteObject(id);
        ctx.selectObject(null);
      },
      testId: 'obj-delete',
    },
    { type: 'line' },
    { title: 'About', icon: 'info', onClick: () => ctx.openDialog({ kind: 'about' }) },
  ];
}

function rangeIndexes(a: number, b: number): number[] {
  const out: number[] = [];
  for (let i = Math.min(a, b); i <= Math.max(a, b); i++) out.push(i);
  return out;
}
