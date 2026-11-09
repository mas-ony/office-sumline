/** Entries of the menus (top menus, context menus and lists), as the interfaces draw them. */
import type { IconName } from './icons';

export interface MenuItem {
  type?: 'line' | 'header';
  title?: string;
  icon?: IconName;
  shortcut?: string;
  disabled?: boolean;
  checked?: boolean;
  submenu?: MenuItem[];
  onClick?: () => void;
  testId?: string;
  /** Keeps the menu open after the click (toggles). */
  keepOpen?: boolean;
}

/** Where a menu opens: at a point, or under (or beside) an element's rectangle. */
export type MenuAnchor = { x: number; y: number } | DOMRect;

function isRect(at: MenuAnchor): at is DOMRect {
  return 'width' in at;
}

/** Places a fixed-position menu at its anchor (below or to the right of a rectangle), inside the viewport. */
export function placeMenu(el: HTMLElement, at: MenuAnchor, side: 'below' | 'right'): void {
  const vw = window.innerWidth;
  const vh = window.innerHeight;
  const w = el.offsetWidth;
  const h = el.offsetHeight;
  let x: number;
  let y: number;
  if (isRect(at)) {
    if (side === 'right') {
      x = at.right - 2;
      y = at.top - 6;
      if (x + w > vw - 4) x = Math.max(4, at.left - w + 2);
    } else {
      x = at.left;
      y = at.bottom + 2;
      if (y + h > vh - 4 && at.top - h - 2 > 4) y = at.top - h - 2;
    }
  } else {
    x = at.x;
    y = at.y;
    if (x + w > vw - 4) x = Math.max(4, vw - w - 4);
    if (y + h > vh - 4) y = Math.max(4, at.y - h);
  }
  if (x + w > vw - 4) x = Math.max(4, vw - w - 4);
  if (y + h > vh - 4) y = Math.max(4, vh - h - 4);
  el.style.left = `${Math.round(x)}px`;
  el.style.top = `${Math.round(y)}px`;
}
