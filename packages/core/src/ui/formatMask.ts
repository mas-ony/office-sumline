/**
 * The Format Cells dialog's settings: the Mask tab's categories and the format code they make (and
 * read back), and the Editor tab's cell types, stored as number formats and data validation.
 */
import type { SpreadsheetDocument } from '../core/document';

export type FormatCategory = 'general' | 'number' | 'currency' | 'accounting' | 'percentage' | 'text' | 'custom';

export const FORMAT_CATEGORIES: { key: FormatCategory; label: string }[] = [
  { key: 'general', label: 'General' },
  { key: 'number', label: 'Number' },
  { key: 'currency', label: 'Currency' },
  { key: 'accounting', label: 'Accounting' },
  { key: 'percentage', label: 'Percentage' },
  { key: 'text', label: 'Text' },
  { key: 'custom', label: 'Custom' },
];

export interface MaskState {
  category: FormatCategory;
  decimals: string;
  thousands: boolean;
  symbol: string;
  negative: number;
  custom: string;
}

function decimalsPart(n: number): string {
  return n > 0 ? `.${'0'.repeat(Math.min(30, n))}` : '';
}

/** Number format code for the Mask tab settings. */
export function maskCode(s: MaskState): string {
  const d = Math.max(0, Math.min(30, Number(s.decimals) || 0));
  const sym = s.symbol.replace(/"/g, '');
  const withNeg = (base: string) => {
    switch (s.negative) {
      case 1: return `${base};[Red]${base}`;
      case 2: return `${base};\\(${base}\\)`;
      case 3: return `${base};[Red]\\(${base}\\)`;
      default: return base;
    }
  };
  switch (s.category) {
    case 'general': return 'General';
    case 'number': return withNeg(`${s.thousands ? '#,##0' : '0'}${decimalsPart(d)}`);
    case 'currency': return withNeg(`${sym ? `"${sym}"` : ''}#,##0${decimalsPart(d)}`);
    case 'accounting': {
      const n = `#,##0${decimalsPart(d)}`;
      const q = sym ? `"${sym}"` : '';
      return `_(${q}* ${n}_);_(${q}* \\(${n}\\);_(${q}* "-"${d ? '??' : ''}_);_(@_)`;
    }
    case 'percentage': return `0${decimalsPart(d)}%`;
    case 'text': return '@';
    case 'custom': return s.custom || 'General';
  }
}

/** Reads the Mask tab settings back from a format code. */
export function maskFromCode(code: string): MaskState {
  const base: MaskState = { category: 'custom', decimals: '2', thousands: true, symbol: '$', negative: 0, custom: code };
  if (/^General$/i.test(code)) return { ...base, category: 'general' };
  if (code === '@') return { ...base, category: 'text' };
  const dec = (/\.(0+)/.exec(code)?.[1].length ?? 0).toString();
  const neg = /;\[Red\]\\?\(/.test(code) ? 3 : /;\\?\(/.test(code) ? 2 : /;\[Red\]/.test(code) ? 1 : 0;
  if (/^0(\.0+)?%$/.test(code)) return { ...base, category: 'percentage', decimals: dec };
  if (/^_\(/.test(code)) return { ...base, category: 'accounting', decimals: dec, symbol: /"([^"]*)"/.exec(code)?.[1] ?? '' };
  const first = code.split(';')[0];
  if (/^"[^"]*"#,##0(\.0+)?$/.test(first)) return { ...base, category: 'currency', decimals: dec, symbol: /"([^"]*)"/.exec(first)?.[1] ?? '', negative: neg };
  if (/^(#,##)?0(\.0+)?$/.test(first)) return { ...base, category: 'number', decimals: dec, thousands: first.startsWith('#,##'), negative: neg };
  return base;
}

export type EditorType = 'text' | 'numeric' | 'percent' | 'dropdown' | 'checkbox' | 'calendar';

export const EDITOR_TYPES: { key: EditorType; label: string }[] = [
  { key: 'text', label: 'Text' },
  { key: 'numeric', label: 'Numeric' },
  { key: 'percent', label: 'Percent' },
  { key: 'dropdown', label: 'Dropdown' },
  { key: 'checkbox', label: 'Checkbox' },
  { key: 'calendar', label: 'Calendar' },
];

/** Applies the Editor tab to the selected cells: wrapping, and the cell type. */
export function applyCellEditor(doc: SpreadsheetDocument, editor: EditorType, options: string, wrap: boolean): void {
  doc.format({ align: { wrap: wrap || null } }, undefined, undefined, 'Wrap');
  const ranges = doc.selection.ranges;
  const sqref = ranges.map((x) => ({ ...x }));
  switch (editor) {
    case 'numeric':
      doc.setNumberFormat('General');
      break;
    case 'percent':
      doc.setNumberFormat('0%');
      break;
    case 'calendar':
      doc.setNumberFormat('dd/mm/yyyy');
      doc.addValidation({ sqref, type: 'date', operator: 'greaterThan', formula1: '1', allowBlank: true, showErrorMessage: true, errorStyle: 'stop' });
      break;
    case 'dropdown': {
      const items = options.split(/[\n,]/).map((s) => s.trim()).filter(Boolean);
      if (items.length) doc.addValidation({ sqref, type: 'list', formula1: `"${items.join(',').replace(/"/g, '')}"`, allowBlank: true, showErrorMessage: true, errorStyle: 'stop' });
      break;
    }
    case 'checkbox':
      doc.addValidation({ sqref, type: 'list', formula1: '"TRUE,FALSE"', allowBlank: true, showErrorMessage: true, errorStyle: 'stop' });
      break;
    default:
      break;
  }
}
