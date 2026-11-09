/** Cleans up a formula the way Excel does on entry: upper-case names, closed parentheses. */
import { tokenize, refToText } from '../formula/tokenizer';
import { registry } from '../formula/functions/helpers';
import { quoteSheetName } from '../refs';
import '../formula/functions';

export function normalizeFormulaInput(text: string): string {
  let src = text.startsWith('=') ? text.slice(1) : text;
  // Close an unterminated string.
  const quotes = (src.match(/"/g) ?? []).length;
  if (quotes % 2 === 1) src += '"';
  let tokens;
  try {
    tokens = tokenize(src);
  } catch {
    return src;
  }
  let depth = 0;
  let out = '';
  for (const t of tokens) {
    switch (t.type) {
      case 'func': {
        const name = String(t.value);
        const upper = name.toUpperCase();
        out += registry.has(upper) || upper === 'LET' || upper.startsWith('_XL') ? upper : name;
        break;
      }
      case 'ref':
        out += t.ref && t.ref.kind !== 'error' ? refToText(t.ref, quoteSheetName) : t.text.toUpperCase();
        break;
      case 'bool':
        out += t.text.toUpperCase();
        break;
      case 'error':
        out += t.text.toUpperCase();
        break;
      case 'open':
        depth++;
        out += t.text;
        break;
      case 'close':
        depth--;
        out += t.text;
        break;
      default:
        out += t.text;
    }
  }
  while (depth-- > 0) out += ')';
  return out;
}
