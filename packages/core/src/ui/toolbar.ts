/** The toolbar's items, and how many fit in its width (the rest go to the "more" panel). */
import type { BorderKind } from '../core/document';
import type { BorderLineStyle } from '../core/model/style';
import type { IconName } from './icons';

export const ALIGN_ICONS: Record<string, IconName> = { left: 'format_align_left', center: 'format_align_center', right: 'format_align_right', justify: 'format_align_justify' };
export const VALIGN_ICONS: Record<string, IconName> = { top: 'vertical_align_top', center: 'vertical_align_center', bottom: 'vertical_align_bottom' };

export const BORDER_KINDS: { kind: BorderKind; icon: IconName; label: string }[] = [
  { kind: 'all', icon: 'border_all', label: 'All borders' },
  { kind: 'outside', icon: 'border_outer', label: 'Outer borders' },
  { kind: 'inside', icon: 'border_inner', label: 'Inner borders' },
  { kind: 'horizontal', icon: 'border_horizontal', label: 'Horizontal borders' },
  { kind: 'vertical', icon: 'border_vertical', label: 'Vertical borders' },
  { kind: 'left', icon: 'border_left', label: 'Left border' },
  { kind: 'top', icon: 'border_top', label: 'Top border' },
  { kind: 'right', icon: 'border_right', label: 'Right border' },
  { kind: 'bottom', icon: 'border_bottom', label: 'Bottom border' },
  { kind: 'none', icon: 'border_clear', label: 'Clear borders' },
];

export type LineKind = 'solid' | 'dotted' | 'dashed' | 'double';
export const LINE_KINDS: LineKind[] = ['solid', 'dotted', 'dashed', 'double'];

/** Toolbar thickness (1-5 px) and line kind to an XLSX border style. */
export function borderStyleFor(width: number, kind: LineKind): BorderLineStyle {
  if (kind === 'double') return 'double';
  if (kind === 'dotted') return width >= 2 ? 'mediumDashDotDot' : 'dotted';
  if (kind === 'dashed') return width >= 2 ? 'mediumDashed' : 'dashed';
  return width >= 3 ? 'thick' : width === 2 ? 'medium' : 'thin';
}

export type ToolKey =
  | 'undo' | 'redo' | 'save' | 'd1' | 'font' | 'size' | 'align' | 'bold' | 'color' | 'fill' | 'valign' | 'merge'
  | 'border' | 'd2' | 'image' | 'search' | 'chart' | 'rule' | 'shape';

const W_BTN = 33;
const W_PICK = 47;

/** The toolbar's items in order, with their widths in px; the keys starting with 'd' are dividers. */
export const TOOLBAR_ITEMS: readonly { key: ToolKey; width: number }[] = [
  { key: 'undo', width: W_BTN },
  { key: 'redo', width: W_BTN },
  { key: 'save', width: W_BTN },
  { key: 'd1', width: 11 },
  { key: 'font', width: 122 },
  { key: 'size', width: W_PICK },
  { key: 'align', width: W_PICK },
  { key: 'bold', width: W_BTN },
  { key: 'color', width: W_BTN },
  { key: 'fill', width: W_BTN },
  { key: 'valign', width: W_PICK },
  { key: 'merge', width: W_BTN },
  { key: 'border', width: W_PICK },
  { key: 'd2', width: 11 },
  { key: 'image', width: W_BTN },
  { key: 'search', width: W_BTN },
  { key: 'chart', width: W_BTN },
  { key: 'rule', width: W_BTN },
  { key: 'shape', width: W_BTN },
];

export const isDivider = (key: ToolKey): boolean => key.startsWith('d');

/** The items shown in a toolbar `avail` px wide, and those moved to the "more" panel (without dividers). */
export function fitToolbar(avail: number): { shown: ToolKey[]; overflow: ToolKey[] } {
  const items = TOOLBAR_ITEMS;
  const total = items.reduce((s, i) => s + i.width, 0);
  let visible = items.length;
  if (total > avail) {
    // The "more" button takes 40 px.
    let used = 40;
    visible = 0;
    for (const it of items) {
      if (used + it.width > avail) break;
      used += it.width;
      visible++;
    }
    // Never end the visible part with a divider.
    while (visible > 0 && isDivider(items[visible - 1].key)) visible--;
  }
  return {
    shown: items.slice(0, visible).map((i) => i.key),
    overflow: items.slice(visible).filter((i) => !isDivider(i.key)).map((i) => i.key),
  };
}
