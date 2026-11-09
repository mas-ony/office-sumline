/** The document state the interface shows, and the option lists of the menus and toolbar. */
import type { EditMode, SpreadsheetDocument } from '../core/document';
import { rangeName } from '../core/refs';
import { resolveColor } from '../core/model/color';
import { fillColor } from '../grid/render';

/** Fonts of the toolbar's font picker ("Default" is the workbook's default font). */
export const TOOLBAR_FONTS = ['Default', 'Verdana', 'Arial', 'Courier New'];

/** The toolbar's size picker uses CSS keywords; stored as points (16px = 12pt). */
export const TOOLBAR_SIZES: { label: string; pt: number; css: string }[] = [
  { label: 'x-small', pt: 7.5, css: 'x-small' },
  { label: 'small', pt: 10, css: 'small' },
  { label: 'medium', pt: 12, css: 'medium' },
  { label: 'large', pt: 13.5, css: 'large' },
  { label: 'x-large', pt: 18, css: 'x-large' },
];

/** Format > Font size (points, as in Excel). */
export const MENU_FONT_SIZES = [6, 7, 8, 9, 10, 11, 12, 14, 16, 18, 24, 36];

/** Format > Format presets: label, number format code, sample. */
export const FORMAT_PRESETS: { label: string; code: string; sample: string }[] = [
  { label: 'Text', code: '@', sample: '' },
  { label: 'Number', code: '#,##0.00', sample: '1,234.00' },
  { label: 'Percentage', code: '0.00%', sample: '10.00%' },
  { label: 'Currency', code: '"$" #,##0.00', sample: '$ 10.00' },
  { label: 'Date', code: 'dd/mm/yyyy', sample: '25/07/1998' },
  { label: 'Hour', code: 'h:mm:ss', sample: '12:59:59' },
  { label: 'Date and hour', code: 'dd/mm/yyyy h:mm:ss', sample: '25/07/1998 12:59:59' },
];

/** Custom format codes offered by the Format dialog. */
export const CUSTOM_FORMATS = [
  'General', '0', '0.00', '#,##0', '#,##0.00', '#,##0;-#,##0', '#,##0;[Red]-#,##0', '#,##0.00;-#,##0.00',
  '#,##0.00;[Red]-#,##0.00', '0%', '0.00%', '0.00E+00', '##0.0E+0', 'dd/mm/yyyy', 'dd/mmm/yy', 'dd/mmm', 'mmm/yy',
  'h:mm AM/PM', 'h:mm:ss AM/PM', 'hh:mm', 'hh:mm:ss', 'dd/mm/yyyy hh:mm', 'mm:ss', '@', '[h]:mm:ss',
  '_-* #,##0_-;-* #,##0_-;_-* "-"_-;_-@_-', '_-* #,##0.00_-;-* #,##0.00_-;_-* "-"??_-;_-@_-',
];

export const ZOOM_LEVELS = [0.5, 0.75, 0.9, 1, 1.25, 1.5, 2];

/** Document events after which the document state is computed again. */
export const DOC_STATE_EVENTS = ['change', 'selection', 'activeSheet', 'sheets', 'mode', 'history', 'file'] as const;

export interface FormatState {
  fontName: string;
  /** True when the font is the workbook default ("Default" in the picker). */
  defaultFont: boolean;
  fontSize: number;
  bold: boolean;
  italic: boolean;
  underline: boolean;
  strike: boolean;
  color: string;
  fill: string | null;
  hAlign: string;
  vAlign: string;
  wrap: 'wrap' | 'clip' | 'overflow';
  rotation: number;
  merged: boolean;
  numFmt: string;
  hasLink: boolean;
  hasComment: boolean;
  commentKind?: 'note' | 'threaded';
}

export interface DocState {
  version: number;
  mode: EditMode;
  readOnly: boolean;
  canUndo: boolean;
  canRedo: boolean;
  activeRef: string;
  rangeRef: string;
  multiCell: boolean;
  editText: string;
  format: FormatState;
  sheets: { name: string; index: number; state: string; color?: string }[];
  activeIndex: number;
  fileName: string;
  modified: boolean;
  hasFilter: boolean;
  hasFooter: boolean;
  freezeRows: number;
  freezeCols: number;
}

/** The state of a document for a version (a count of its events). */
export function computeDocState(doc: SpreadsheetDocument, version: number): DocState {
  const sel = doc.selection;
  const sheet = doc.sheet;
  const { r, c } = sel.active;
  const style = doc.styleAt(r, c);
  const theme = doc.wb.theme;
  const range = sel.ranges[sel.ranges.length - 1];
  const align = style.align ?? {};
  const m = sheet.mergeAt(r, c);
  const merged = sheet.merges.some((x) => x.r0 === range.r0 && x.c0 === range.c0 && x.r1 === range.r1 && x.c1 === range.c1) || (!!m && range.r0 === range.r1 && range.c0 === range.c1);
  const single = (range.r0 === range.r1 && range.c0 === range.c1) || (!!m && m.r0 === range.r0 && m.c0 === range.c0 && m.r1 === range.r1 && m.c1 === range.c1);
  const thread = sheet.commentAt(m?.r0 ?? r, m?.c0 ?? c);
  return {
    version,
    mode: doc.mode,
    readOnly: doc.isReadOnly(),
    canUndo: doc.history.canUndo(),
    canRedo: doc.history.canRedo(),
    activeRef: rangeName({ r0: r, c0: c, r1: r, c1: c }),
    rangeRef: single ? rangeName({ r0: r, c0: c, r1: r, c1: c }) : rangeName(range),
    multiCell: !single,
    editText: (() => {
      const anchor = doc.anchorOf(r, c);
      const cell = doc.cell(r, c);
      if (anchor && cell?.f === undefined) {
        const a = doc.cell(anchor.r, anchor.c);
        if (a?.f !== undefined && a.arr && !a.dyn) return `{=${a.f}}`;
        if (a?.f !== undefined) return `=${a.f}`;
      }
      if (cell?.f !== undefined && cell.arr && !cell.dyn) return `{=${cell.f}}`;
      return doc.editText(r, c);
    })(),
    format: {
      fontName: style.font.name,
      defaultFont: style.font.name === doc.wb.styles.defaultStyle.font.name,
      fontSize: style.font.size,
      bold: !!style.font.bold,
      italic: !!style.font.italic,
      underline: !!style.font.underline,
      strike: !!style.font.strike,
      color: `#${resolveColor(style.font.color, theme, '000000')}`,
      fill: fillColor(style, theme),
      hAlign: align.horizontal ?? 'general',
      vAlign: align.vertical ?? 'bottom',
      wrap: align.wrap ? 'wrap' : align.clip ? 'clip' : 'overflow',
      rotation: align.rotation ?? 0,
      merged,
      numFmt: style.numFmt,
      hasLink: !!sheet.hyperlinkAt(r, c),
      hasComment: !!thread,
      commentKind: thread?.kind,
    },
    sheets: doc.wb.sheets.map((s, index) => ({ name: s.name, index, state: s.state, color: s.tabColor ? `#${resolveColor(s.tabColor, theme, 'FFFFFF')}` : undefined })),
    activeIndex: doc.activeIndex,
    fileName: doc.fileName,
    modified: doc.modified,
    hasFilter: !!sheet.autoFilter,
    hasFooter: !!sheet.footer,
    freezeRows: sheet.freeze.rows,
    freezeCols: sheet.freeze.cols,
  };
}
