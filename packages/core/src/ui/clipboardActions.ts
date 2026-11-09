/**
 * Clipboard commands for menus and buttons. Keyboard shortcuts use native copy/paste events
 * (handled by the grid); menus cannot, so copying goes through execCommand('copy') with a
 * one-off listener, and pasting uses the async Clipboard API where the browser allows it (secure
 * contexts) and otherwise the editor's internal clipboard.
 */
import type { SpreadsheetDocument } from '../core/document';
import type { GridView } from '../grid/GridView';

export function copySelection(doc: SpreadsheetDocument, grid: GridView | null, cut: boolean): boolean {
  const payload = cut ? doc.cut() : doc.copy();
  if (!payload) return false;
  grid?.markClip();
  let written = false;
  const onCopy = (e: ClipboardEvent) => {
    if (!e.clipboardData) return;
    e.clipboardData.setData('text/plain', payload.text);
    e.clipboardData.setData('text/html', payload.html);
    e.preventDefault();
    e.stopImmediatePropagation();
    written = true;
  };
  document.addEventListener('copy', onCopy, true);
  try {
    document.execCommand('copy');
  } catch {
    // Not supported; the internal clipboard still works.
  } finally {
    document.removeEventListener('copy', onCopy, true);
  }
  if (!written) void navigator.clipboard?.writeText?.(payload.text).catch(() => undefined);
  return true;
}

export type PasteMode = 'all' | 'values' | 'formats' | 'formulas';

/** Returns false when nothing could be pasted (no clipboard access and nothing copied here). */
export async function pasteClipboard(doc: SpreadsheetDocument, mode: PasteMode = 'all'): Promise<boolean> {
  const clip = navigator.clipboard;
  if (clip && window.isSecureContext) {
    try {
      let html: string | undefined;
      let text: string | undefined;
      if (clip.read) {
        const items = await clip.read();
        for (const item of items) {
          if (!html && item.types.includes('text/html')) html = await (await item.getType('text/html')).text();
          if (text === undefined && item.types.includes('text/plain')) text = await (await item.getType('text/plain')).text();
        }
      } else text = await clip.readText();
      if (html !== undefined || text !== undefined) {
        doc.paste({ text, html }, mode);
        return true;
      }
    } catch {
      // Permission denied or unsupported: fall back to the internal clipboard.
    }
  }
  if (doc.hasClip()) {
    doc.paste({}, mode);
    return true;
  }
  return false;
}
