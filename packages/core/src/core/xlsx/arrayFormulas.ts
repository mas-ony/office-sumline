/**
 * Decides whether a formula typed in the editor (dynamic array semantics) must be saved as an
 * array formula so that Excel evaluates it the same way. Formulas that only pass ranges to
 * functions that take ranges (SUM(A1:A3)) can be saved as plain formulas.
 */
import { parseFormula, type Node } from '../formula/parser';
import { registry } from '../formula/functions/helpers';
import '../formula/functions';

const ARRAY_RESULT = new Set([
  'SEQUENCE', 'RANDARRAY', 'FILTER', 'SORT', 'SORTBY', 'UNIQUE', 'TRANSPOSE', 'MMULT', 'MINVERSE', 'MUNIT', 'FREQUENCY',
  'TEXTSPLIT', 'VSTACK', 'HSTACK', 'TOCOL', 'TOROW', 'WRAPROWS', 'WRAPCOLS', 'TAKE', 'DROP', 'CHOOSEROWS', 'CHOOSECOLS',
  'EXPAND', 'MODE.MULT', 'LINEST', 'LOGEST', 'TREND', 'GROWTH', 'REGEXEXTRACT',
]);

function isMultiRef(n: Node): boolean {
  if (n.t === 'ref') return n.ref.kind !== 'cell' && n.ref.kind !== 'error';
  if (n.t === 'bin' && (n.op === ':' || n.op === ' ')) return true;
  if (n.t === 'arr') return n.rows.length * (n.rows[0]?.length ?? 0) > 1;
  return false;
}

function walk(n: Node, scalarPosition: boolean): boolean {
  switch (n.t) {
    case 'ref':
    case 'bin':
      if (n.t === 'bin' && n.op !== ':' && n.op !== ' ') return walk(n.l, true) || walk(n.r, true);
      return scalarPosition && isMultiRef(n);
    case 'arr':
      return scalarPosition && isMultiRef(n);
    case 'un':
    case 'pct':
      return walk(n.arg, true);
    case 'spill':
      return true;
    case 'union':
      return n.items.some((i) => walk(i, false));
    case 'func': {
      if (ARRAY_RESULT.has(n.name)) return true;
      const def = registry.get(n.name);
      const scalar = def?.scalar;
      return n.args.some((a, i) => walk(a, scalar === 'all' || (Array.isArray(scalar) && scalar.includes(i))));
    }
    default:
      return false;
  }
}

const WORD = /(\$?[A-Za-z_\\][A-Za-z0-9_.$]*)(\s*\()?/g;
const CELL = /^\$?[A-Za-z]{1,3}\$?\d+$/;

/**
 * Answers without parsing for the common case: only single-cell references, literals, operators
 * and functions. Returns null when a range, name, array constant or string needs the parser.
 */
function quickAnswer(formula: string): boolean | null {
  if (/[:{#@"'!\[]/.test(formula)) return null;
  for (const m of formula.matchAll(WORD)) {
    const word = m[1];
    if (m[2]) {
      if (ARRAY_RESULT.has(word.toUpperCase())) return true;
      continue;
    }
    if (CELL.test(word) || /^(TRUE|FALSE)$/i.test(word)) continue;
    // Exponents of numbers such as 1E5 start inside a number literal.
    const before = formula[m.index! - 1];
    if (before && /[0-9.]/.test(before)) continue;
    return null;
  }
  return false;
}

const cache = new Map<string, boolean>();

export function needsArrayFormula(formula: string): boolean {
  const quick = quickAnswer(formula);
  if (quick !== null) return quick;
  let v = cache.get(formula);
  if (v === undefined) {
    try {
      v = walk(parseFormula(formula), true);
    } catch {
      v = false;
    }
    if (cache.size > 5000) cache.clear();
    cache.set(formula, v);
  }
  return v;
}
