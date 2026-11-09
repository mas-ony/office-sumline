/**
 * Formula editing help: which function name is being typed (for autocomplete) and which function
 * call and argument the caret is in (for the argument hint).
 */
import { type FunctionInfo, functionInfo, matchFunctions } from './functionCatalog';
import { tokenize } from '../core/formula/tokenizer';
import { REF_COLORS } from '../grid/GridView';

export interface FormulaContext {
  /** Identifier being typed and its position. */
  ident?: { text: string; start: number; end: number };
  suggestions: FunctionInfo[];
  /** Innermost function call around the caret. */
  call?: { info: FunctionInfo; argIndex: number };
}

const IDENT_BEFORE = /[A-Za-z_][A-Za-z0-9_.]*$/;
const IDENT_AFTER = /^[A-Za-z0-9_.]*/;

export function analyzeFormula(text: string, caret: number, names: string[] = []): FormulaContext | null {
  if (!text.startsWith('=')) return null;
  const before = text.slice(0, caret);
  const stack: { name: string | null; argIndex: number }[] = [];
  let inString = false;
  let inSheet = false;
  for (let i = 1; i < before.length; i++) {
    const ch = before[i];
    if (inString) {
      if (ch === '"') {
        if (before[i + 1] === '"') i++;
        else inString = false;
      }
      continue;
    }
    if (inSheet) {
      if (ch === "'") {
        if (before[i + 1] === "'") i++;
        else inSheet = false;
      }
      continue;
    }
    if (ch === '"') inString = true;
    else if (ch === "'") inSheet = true;
    else if (ch === '(') {
      const m = IDENT_BEFORE.exec(before.slice(0, i));
      stack.push({ name: m ? m[0].toUpperCase().replace(/^_XLFN\./, '') : null, argIndex: 0 });
    } else if (ch === '{') stack.push({ name: null, argIndex: 0 });
    else if (ch === ')' || ch === '}') stack.pop();
    else if (ch === ',' || ch === ';') {
      const top = stack[stack.length - 1];
      if (top) top.argIndex++;
    }
  }
  if (inString || inSheet) return { suggestions: [] };

  let call: FormulaContext['call'];
  for (let i = stack.length - 1; i >= 0; i--) {
    const entry = stack[i];
    if (!entry.name) continue;
    const info = functionInfo(entry.name);
    if (info) call = { info, argIndex: entry.argIndex };
    break;
  }

  const m = IDENT_BEFORE.exec(before);
  let suggestions: FunctionInfo[] = [];
  let ident: FormulaContext['ident'];
  if (m) {
    const start = caret - m[0].length;
    const prev = before[start - 1] ?? '';
    // Only at the start of an operand (not after a sheet name, a number or a colon).
    if (/[=(,;+\-*/^&<>{ ]|^$/.test(prev) && !/^\d/.test(m[0])) {
      const end = caret + (IDENT_AFTER.exec(text.slice(caret))?.[0].length ?? 0);
      ident = { text: m[0], start, end };
      suggestions = matchFunctions(m[0]);
      const upper = m[0].toUpperCase();
      for (const n of names) {
        if (n.toUpperCase().startsWith(upper)) suggestions.push({ name: n, category: 'Name', args: [], description: 'Defined name' });
      }
      // Nothing to suggest when the identifier is already a complete name.
      if (suggestions.length === 1 && suggestions[0].name === upper && text[end] === '(') suggestions = [];
    }
  }
  return { ident, suggestions, call };
}

/** Index of the argument to highlight in a signature (repeating groups wrap around). */
export function highlightedArg(info: FunctionInfo, argIndex: number): number {
  const args = info.args;
  const ell = args.indexOf('…');
  if (ell < 0 || argIndex < ell) return Math.min(argIndex, args.length - 1);
  let start = ell - 1;
  while (start > 0 && !args[start].startsWith('[')) start--;
  const size = Math.max(1, ell - start);
  return start + ((argIndex - start) % size);
}

/** Inserts a picked function (or name) in place of the identifier being typed. */
export function applySuggestion(text: string, ident: { start: number; end: number }, item: FunctionInfo): { text: string; caret: number } {
  const isName = item.category === 'Name';
  const after = text.slice(ident.end);
  const insert = isName || after.startsWith('(') ? item.name : `${item.name}(`;
  const next = text.slice(0, ident.start) + insert + after;
  const caret = ident.start + insert.length + (!isName && after.startsWith('(') ? 1 : 0);
  return { text: next, caret };
}

/** What the suggestion list or the argument hint shows while a formula is edited. */
export interface AssistView {
  /** The list is shown (Tab, or Enter after moving through it, applies an entry). */
  showList: boolean;
  suggestions: FunctionInfo[];
  /** The highlighted entry of the list. */
  index: number;
  /** The function call around the caret, for the argument hint (when the list is not shown). */
  call?: { info: FunctionInfo; argIndex: number };
  visible: boolean;
}

/** An edit as the assist needs it (the grid's edit state). */
interface AssistEdit {
  active: boolean;
  text: string;
  caret: number;
}

/**
 * The suggestion list and the argument hint of the cell editor and the formula bar: the entries
 * for the text at the caret, the highlighted one, and the keys that move through and apply them.
 */
export class FormulaAssistState {
  view: AssistView = { showList: false, suggestions: [], index: 0, visible: false };
  private ctx: FormulaContext | null = null;
  private edit: AssistEdit = { active: false, text: '', caret: 0 };
  private identKey: string | null = null;
  private index = 0;
  private picked = false;
  private dismissed: string | null = null;

  /**
   * `names`: defined names offered with the functions; `replace` writes the edit text after a
   * suggestion is applied; `changed` is called whenever `view` changes.
   */
  constructor(private readonly names: string[], private readonly replace: (text: string, caret: number) => void, private readonly changed: () => void) {}

  /** The edit changed (text, caret, or editing started or ended). */
  update(edit: AssistEdit): void {
    if (edit.active === this.edit.active && edit.text === this.edit.text && edit.caret === this.edit.caret) return;
    this.edit = { active: edit.active, text: edit.text, caret: edit.caret };
    this.ctx = edit.active ? analyzeFormula(edit.text, edit.caret, this.names) : null;
    const identKey = this.ctx?.ident ? `${this.ctx.ident.start}:${this.ctx.ident.text}` : null;
    if (identKey !== this.identKey) {
      this.identKey = identKey;
      this.index = 0;
      this.picked = false;
    }
    this.refresh();
  }

  /** Applies a suggestion to the edit. */
  apply(item: FunctionInfo): void {
    if (!this.ctx?.ident) return;
    const res = applySuggestion(this.edit.text, this.ctx.ident, item);
    this.replace(res.text, res.caret);
  }

  /** Handles a key while editing; returns true when the key was consumed. */
  handleKey(e: KeyboardEvent): boolean {
    const { showList, suggestions } = this.view;
    if (!showList || e.isComposing) return false;
    const n = suggestions.length;
    switch (e.key) {
      case 'ArrowDown':
        this.index = (this.index + 1) % n;
        this.picked = true;
        this.refresh();
        return true;
      case 'ArrowUp':
        this.index = (this.index - 1 + n) % n;
        this.picked = true;
        this.refresh();
        return true;
      case 'Tab':
        this.apply(suggestions[this.index] ?? suggestions[0]);
        return true;
      case 'Enter':
        if (!this.picked || e.shiftKey || e.altKey || e.ctrlKey || e.metaKey) return false;
        this.apply(suggestions[this.index] ?? suggestions[0]);
        return true;
      case 'Escape':
        this.dismissed = this.identKey;
        this.refresh();
        return true;
      default:
        return false;
    }
  }

  private refresh(): void {
    const suggestions = (this.ctx?.suggestions ?? []).slice(0, 30);
    const showList = suggestions.length > 0 && this.dismissed !== this.identKey;
    const call = !showList ? this.ctx?.call : undefined;
    const visible = this.edit.active && (showList || !!call);
    const v = this.view;
    if (v.showList === showList && v.index === this.index && v.call === call && v.visible === visible && v.suggestions.length === suggestions.length && v.suggestions.every((s, i) => s === suggestions[i])) return;
    this.view = { showList, suggestions, index: this.index, call, visible };
    this.changed();
  }
}

/** A run of the formula bar's text; references are coloured like their highlights on the grid. */
export interface FormulaSegment {
  text: string;
  color?: string;
  /** Position in the text (a stable key). */
  start: number;
}

/** The formula bar's text as runs: references in the colours of the grid's highlights, strings green. */
export function formulaSegments(text: string): FormulaSegment[] {
  if (!text.startsWith('=')) return [{ text, start: 0 }];
  try {
    const out: FormulaSegment[] = [{ text: '=', start: 0 }];
    let pos = 1;
    let i = 0;
    for (const t of tokenize(text.slice(1))) {
      const start = t.start + 1;
      const end = t.end + 1;
      if (start > pos) out.push({ text: text.slice(pos, start), start: pos });
      const s = text.slice(start, end);
      if (t.type === 'ref' && t.ref && t.ref.kind !== 'error') out.push({ text: s, color: REF_COLORS[i++ % REF_COLORS.length], start });
      else if (t.type === 'string') out.push({ text: s, color: '#0b8043', start });
      else out.push({ text: s, start });
      pos = Math.max(pos, end);
    }
    if (pos < text.length) out.push({ text: text.slice(pos), start: pos });
    return out;
  } catch {
    return [{ text, start: 0 }];
  }
}
