/** The Print dialog's page options: the drop-down choices and the switches. */
import { MARGINS, PAPER_SIZES, type PaperSize, type PrintOptions } from '../print/pages';

/** CSS pixels per point (the preview's page sizes). */
export const PT_TO_PX = 96 / 72;

export type PrintChoice<K extends keyof PrintOptions = keyof PrintOptions> = { key: K; label: string; options: { value: PrintOptions[K]; label: string }[] };

export const PRINT_CHOICES: PrintChoice[] = [
  {
    key: 'scope',
    label: 'Scope',
    options: [
      { value: 'sheet', label: 'Current worksheet' },
      { value: 'workbook', label: 'Entire workbook' },
      { value: 'selection', label: 'Selected cells' },
    ],
  },
  { key: 'paper', label: 'Page size', options: (Object.keys(PAPER_SIZES) as PaperSize[]).map((k) => ({ value: k, label: PAPER_SIZES[k].label })) },
  {
    key: 'orientation',
    label: 'Page orientation',
    options: [
      { value: 'portrait', label: 'Portrait' },
      { value: 'landscape', label: 'Landscape' },
    ],
  },
  {
    key: 'scale',
    label: 'Scale',
    options: [
      { value: 'normal', label: 'Normal (100%)' },
      { value: 'fitWidth', label: 'Fit to page width' },
      { value: 'fitHeight', label: 'Fit to page height' },
      { value: 'fitPage', label: 'Fit to page' },
    ],
  },
  { key: 'margins', label: 'Margins', options: (Object.keys(MARGINS) as (keyof typeof MARGINS)[]).map((k) => ({ value: k, label: MARGINS[k].label })) },
  {
    key: 'order',
    label: 'Order of pages',
    options: [
      { value: 'leftToRight', label: 'Left to right' },
      { value: 'topToBottom', label: 'Top to bottom' },
    ],
  },
  {
    key: 'hAlign',
    label: 'Horizontal alignment',
    options: [
      { value: 'left', label: 'Left' },
      { value: 'center', label: 'Center' },
      { value: 'right', label: 'Right' },
    ],
  },
  {
    key: 'vAlign',
    label: 'Vertical alignment',
    options: [
      { value: 'top', label: 'Top' },
      { value: 'center', label: 'Center' },
      { value: 'bottom', label: 'Bottom' },
    ],
  },
] as PrintChoice[];

export const PRINT_SWITCHES: { key: keyof PrintOptions; label: string }[] = [
  { key: 'gridlines', label: 'Show grid lines' },
  { key: 'notes', label: 'Show notes' },
  { key: 'headings', label: 'Show row and column identifiers' },
  { key: 'pageNumbers', label: 'Show page number' },
  { key: 'spreadsheetTitle', label: 'Show spreadsheet title' },
  { key: 'worksheetTitle', label: 'Show worksheet title' },
  { key: 'date', label: 'Current date' },
  { key: 'time', label: 'Current time' },
  { key: 'repeatRows', label: 'Repeat frozen rows' },
  { key: 'repeatCols', label: 'Repeat frozen columns' },
  { key: 'objects', label: 'Show pictures, charts and shapes' },
];
