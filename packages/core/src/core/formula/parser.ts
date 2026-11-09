import { CellError, type ErrorCode } from '../values';
import { tokenize, type RefParts, type Token } from './tokenizer';

export type Node =
  | { t: 'num'; v: number }
  | { t: 'str'; v: string }
  | { t: 'bool'; v: boolean }
  | { t: 'err'; v: CellError }
  | { t: 'ref'; ref: RefParts }
  | { t: 'name'; name: string; sheet?: string }
  | { t: 'struct'; text: string; table: string }
  | { t: 'func'; name: string; args: Node[] }
  | { t: 'un'; op: '-' | '+' | '@'; arg: Node }
  | { t: 'pct'; arg: Node }
  | { t: 'bin'; op: string; l: Node; r: Node }
  | { t: 'union'; items: Node[] }
  | { t: 'arr'; rows: Node[][] }
  | { t: 'spill'; arg: Node }
  | { t: 'missing' };

export class FormulaSyntaxError extends Error {
  constructor(message: string, readonly position: number) {
    super(message);
  }
}

/** Function names stored in files with a prefix (_xlfn., _xlws.) are normalised without it. */
export function normalizeFunctionName(name: string): string {
  let n = name.toUpperCase();
  if (n.startsWith('_XLFN.')) n = n.slice(6);
  if (n.startsWith('_XLWS.')) n = n.slice(6);
  if (n.startsWith('_XLL.')) n = n.slice(5);
  return n;
}

const BINARY_BP: Record<string, number> = {
  '=': 10,
  '<>': 10,
  '<': 10,
  '>': 10,
  '<=': 10,
  '>=': 10,
  '&': 20,
  '+': 30,
  '-': 30,
  '*': 40,
  '/': 40,
  '^': 50,
};
const PERCENT_BP = 60;
const PREFIX_BP = 70;
const INTERSECT_BP = 85;
const RANGE_BP = 90;

class Parser {
  private pos = 0;
  constructor(private readonly tokens: Token[]) {}

  private peek(skipWs = true): Token | undefined {
    let p = this.pos;
    if (skipWs) while (this.tokens[p]?.type === 'ws') p++;
    return this.tokens[p];
  }

  private next(): Token | undefined {
    while (this.tokens[this.pos]?.type === 'ws') this.pos++;
    return this.tokens[this.pos++];
  }

  private expect(type: Token['type']): Token {
    const t = this.next();
    if (!t || t.type !== type) throw new FormulaSyntaxError(`Expected ${type}`, t?.start ?? -1);
    return t;
  }

  parse(): Node {
    if (!this.peek()) throw new FormulaSyntaxError('Empty formula', 0);
    const node = this.expr(0);
    const rest = this.peek();
    if (rest) throw new FormulaSyntaxError(`Unexpected "${rest.text}"`, rest.start);
    return node;
  }

  private startsOperand(t: Token | undefined): boolean {
    if (!t) return false;
    return ['ref', 'name', 'func', 'open', 'structured'].includes(t.type);
  }

  private expr(minBp: number): Node {
    let left = this.prefix();
    for (;;) {
      // Intersection: whitespace between two reference expressions.
      const raw = this.tokens[this.pos];
      if (raw?.type === 'ws' && INTERSECT_BP > minBp && isReferenceNode(left)) {
        const after = this.peek();
        if (this.startsOperand(after)) {
          this.pos++;
          const right = this.expr(INTERSECT_BP);
          left = { t: 'bin', op: ' ', l: left, r: right };
          continue;
        }
      }
      const t = this.peek();
      if (!t) break;
      if (t.type === 'colon' && RANGE_BP > minBp) {
        this.next();
        const right = this.expr(RANGE_BP);
        left = { t: 'bin', op: ':', l: left, r: right };
        continue;
      }
      if (t.type === 'spill') {
        this.next();
        left = { t: 'spill', arg: left };
        continue;
      }
      if (t.type !== 'op') break;
      const op = String(t.value);
      if (op === '%') {
        if (PERCENT_BP <= minBp) break;
        this.next();
        left = { t: 'pct', arg: left };
        continue;
      }
      const bp = BINARY_BP[op];
      if (bp === undefined || bp <= minBp) break;
      this.next();
      // All Excel binary operators are left associative (2^3^2 = 64).
      const right = this.expr(bp);
      left = { t: 'bin', op, l: left, r: right };
    }
    return left;
  }

  private prefix(): Node {
    const t = this.next();
    if (!t) throw new FormulaSyntaxError('Unexpected end of formula', -1);
    switch (t.type) {
      case 'number':
        return { t: 'num', v: Number(t.value) };
      case 'string':
        return { t: 'str', v: String(t.value) };
      case 'bool':
        return { t: 'bool', v: Boolean(t.value) };
      case 'error':
        return { t: 'err', v: CellError.of(String(t.value) as ErrorCode) };
      case 'ref':
        return { t: 'ref', ref: t.ref! };
      case 'name':
        return { t: 'name', name: String(t.value), sheet: t.sheet };
      case 'structured':
        return { t: 'struct', text: t.text, table: String(t.value ?? '') };
      case 'func':
        return this.call(t);
      case 'open': {
        const items: Node[] = [this.expr(0)];
        while (this.peek()?.type === 'comma') {
          this.next();
          items.push(this.expr(0));
        }
        this.expect('close');
        return items.length === 1 ? items[0] : { t: 'union', items };
      }
      case 'brace-open':
        return this.array();
      case 'op': {
        const op = String(t.value);
        if (op === '-' || op === '+') {
          const arg = this.expr(PREFIX_BP);
          return { t: 'un', op, arg };
        }
        throw new FormulaSyntaxError(`Unexpected operator "${op}"`, t.start);
      }
      case 'at': {
        const arg = this.expr(PREFIX_BP);
        return { t: 'un', op: '@', arg };
      }
      default:
        throw new FormulaSyntaxError(`Unexpected "${t.text}"`, t.start);
    }
  }

  private call(t: Token): Node {
    const name = normalizeFunctionName(String(t.value));
    this.expect('open');
    const args: Node[] = [];
    if (this.peek()?.type === 'close') {
      this.next();
      return { t: 'func', name, args };
    }
    for (;;) {
      const p = this.peek();
      if (p?.type === 'comma' || p?.type === 'close') args.push({ t: 'missing' });
      else args.push(this.expr(0));
      const sep = this.next();
      if (!sep) throw new FormulaSyntaxError('Missing ")"', -1);
      if (sep.type === 'close') break;
      if (sep.type !== 'comma') throw new FormulaSyntaxError(`Unexpected "${sep.text}"`, sep.start);
    }
    return { t: 'func', name, args };
  }

  private array(): Node {
    const rows: Node[][] = [[]];
    for (;;) {
      const t = this.peek();
      if (!t) throw new FormulaSyntaxError('Missing "}"', -1);
      if (t.type === 'brace-close') {
        this.next();
        break;
      }
      rows[rows.length - 1].push(this.arrayItem());
      const sep = this.next();
      if (!sep) throw new FormulaSyntaxError('Missing "}"', -1);
      if (sep.type === 'brace-close') break;
      if (sep.type === 'semicolon') rows.push([]);
      else if (sep.type !== 'comma') throw new FormulaSyntaxError(`Unexpected "${sep.text}" in array`, sep.start);
    }
    const width = rows[0].length;
    if (rows.some((r) => r.length !== width)) throw new FormulaSyntaxError('Array rows must have the same size', -1);
    return { t: 'arr', rows };
  }

  private arrayItem(): Node {
    const t = this.next();
    if (!t) throw new FormulaSyntaxError('Unexpected end of array', -1);
    if (t.type === 'op' && (t.value === '-' || t.value === '+')) {
      const n = this.next();
      if (n?.type !== 'number') throw new FormulaSyntaxError('Expected a number', n?.start ?? -1);
      return { t: 'num', v: t.value === '-' ? -Number(n.value) : Number(n.value) };
    }
    switch (t.type) {
      case 'number':
        return { t: 'num', v: Number(t.value) };
      case 'string':
        return { t: 'str', v: String(t.value) };
      case 'bool':
        return { t: 'bool', v: Boolean(t.value) };
      case 'error':
        return { t: 'err', v: CellError.of(String(t.value) as ErrorCode) };
      default:
        throw new FormulaSyntaxError('Arrays can only contain constants', t.start);
    }
  }
}

function isReferenceNode(n: Node): boolean {
  if (n.t === 'ref' || n.t === 'name' || n.t === 'struct') return true;
  if (n.t === 'bin' && (n.op === ':' || n.op === ' ')) return true;
  if (n.t === 'func') return ['INDEX', 'OFFSET', 'INDIRECT', 'CHOOSE', 'IF', 'XLOOKUP'].includes(n.name);
  return n.t === 'union';
}

export function parseFormula(formula: string): Node {
  const src = formula.startsWith('=') ? formula.slice(1) : formula;
  return new Parser(tokenize(src)).parse();
}

/** Collects static references and names in an AST (for dependency tracking). */
export function collectRefs(node: Node, refs: RefParts[] = [], names: { name: string; sheet?: string }[] = [], funcs: Set<string> = new Set()): { refs: RefParts[]; names: { name: string; sheet?: string }[]; funcs: Set<string> } {
  switch (node.t) {
    case 'ref':
      refs.push(node.ref);
      break;
    case 'name':
      names.push({ name: node.name, sheet: node.sheet });
      break;
    case 'func':
      funcs.add(node.name);
      for (const a of node.args) collectRefs(a, refs, names, funcs);
      break;
    case 'un':
    case 'pct':
    case 'spill':
      collectRefs(node.arg, refs, names, funcs);
      break;
    case 'bin':
      collectRefs(node.l, refs, names, funcs);
      collectRefs(node.r, refs, names, funcs);
      break;
    case 'union':
      for (const i of node.items) collectRefs(i, refs, names, funcs);
      break;
    default:
      break;
  }
  return { refs, names, funcs };
}
