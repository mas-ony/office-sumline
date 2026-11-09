import { canonicalPrefixFor, NS } from './namespaces';

/**
 * A small, fast XML DOM tailored to OOXML parts. It keeps element order, attribute values and
 * significant text exactly, so elements the editor does not understand can be written back
 * unchanged. (Adapted from the DOCX editor; the SpreadsheetML main namespace is unprefixed.)
 */
export interface XmlElement {
  type: 'element';
  name: string;
  attrs: Record<string, string>;
  children: XmlNode[];
}

export interface XmlText {
  type: 'text';
  text: string;
}

export type XmlNode = XmlElement | XmlText;

export interface XmlDocument {
  root: XmlElement;
  /** Namespace declarations (prefix -> URI) that are not canonical and must be re-declared on write. */
  extraNamespaces: Record<string, string>;
}

export interface ParseOptions {
  /** Rewrite prefixes to the canonical prefixes (default true). */
  normalize?: boolean;
}

/** Elements whose whitespace-only text content is significant. */
const TEXT_ELEMENTS = new Set(['t', 'a:t', 'v', 'f', 'formula', 'formula1', 'formula2', 'xm:f', 'tc:text', 'x:Anchor']);

const MC_PREFIX_LIST_ATTRS = new Set(['mc:Ignorable', 'mc:MustUnderstand', 'Requires']);
const MC_QNAME_LIST_ATTRS = new Set(['mc:ProcessContent', 'mc:PreserveElements', 'mc:PreserveAttributes']);

const ENTITY_RE = /&(#x[0-9a-fA-F]+|#[0-9]+|lt|gt|amp|quot|apos);/g;

export function decodeEntities(s: string): string {
  if (s.indexOf('&') === -1) return s;
  return s.replace(ENTITY_RE, (_m, ent: string) => {
    switch (ent) {
      case 'lt': return '<';
      case 'gt': return '>';
      case 'amp': return '&';
      case 'quot': return '"';
      case 'apos': return "'";
      default: {
        const code = ent[1] === 'x' ? parseInt(ent.slice(2), 16) : parseInt(ent.slice(1), 10);
        try {
          return String.fromCodePoint(code);
        } catch {
          return '';
        }
      }
    }
  });
}

function isWhitespace(s: string): boolean {
  for (let i = 0; i < s.length; i++) {
    const c = s.charCodeAt(i);
    if (c !== 32 && c !== 9 && c !== 10 && c !== 13) return false;
  }
  return true;
}

type Scope = Map<string, string>;

class NamespaceMapper {
  readonly extra: Record<string, string> = {};
  private unknownByUri = new Map<string, string>();
  private counter = 0;

  constructor(private readonly normalize: boolean) {}

  prefixForUri(uri: string, preferred: string): string {
    if (!this.normalize) return preferred;
    const canon = canonicalPrefixFor(uri);
    if (canon !== undefined) return canon;
    const known = this.unknownByUri.get(uri);
    if (known !== undefined) return known;
    let prefix = preferred || 'ns';
    const taken = (p: string) => p in NS || Object.prototype.hasOwnProperty.call(this.extra, p);
    if (taken(prefix)) {
      do prefix = `ns${++this.counter}`; while (taken(prefix));
    }
    this.unknownByUri.set(uri, prefix);
    this.extra[prefix] = uri;
    return prefix;
  }

  private qualify(prefix: string, local: string): string {
    return prefix ? `${prefix}:${local}` : local;
  }

  resolve(qname: string, scope: Scope, isAttr: boolean): string {
    if (!this.normalize) return qname;
    const idx = qname.indexOf(':');
    if (idx < 0) {
      if (isAttr) return qname;
      const uri = scope.get('');
      if (!uri) return qname;
      const canon = canonicalPrefixFor(uri);
      if (canon !== undefined && canon !== 'xml') return this.qualify(canon, qname);
      return this.qualify(this.prefixForUri(uri, 'ns'), qname);
    }
    const prefix = qname.slice(0, idx);
    if (prefix === 'xml') return qname;
    const uri = scope.get(prefix);
    if (!uri) return qname;
    return this.qualify(this.prefixForUri(uri, prefix), qname.slice(idx + 1));
  }

  /** Translate a whitespace separated list of prefixes or qualified names (markup compatibility). */
  translateList(value: string, scope: Scope, qnames: boolean): string {
    if (!this.normalize) return value;
    return value
      .split(/\s+/)
      .filter(Boolean)
      .map((token) => {
        if (qnames) {
          const idx = token.indexOf(':');
          if (idx < 0) return token;
          const uri = scope.get(token.slice(0, idx));
          return uri ? this.qualify(this.prefixForUri(uri, token.slice(0, idx)), token.slice(idx + 1)) : token;
        }
        const uri = scope.get(token);
        return uri ? this.prefixForUri(uri, token) : token;
      })
      .join(' ');
  }
}

/** Parse an XML string. Throws on unbalanced markup. */
export function parseXml(src: string, opts: ParseOptions = {}): XmlDocument {
  const normalize = opts.normalize !== false;
  const mapper = new NamespaceMapper(normalize);
  const n = src.length;
  let i = 0;
  if (src.charCodeAt(0) === 0xfeff) i = 1;

  const stack: { el: XmlElement; scope: Scope; preserve: boolean }[] = [];
  let root: XmlElement | null = null;
  const rootScope: Scope = new Map([['xml', NS.xml]]);

  const addText = (raw: string, decode = true) => {
    const top = stack[stack.length - 1];
    if (!top) return;
    const t = decode ? decodeEntities(raw) : raw;
    if (!t) return;
    if (isWhitespace(t) && !top.preserve) return;
    const last = top.el.children[top.el.children.length - 1];
    if (last && last.type === 'text') last.text += t;
    else top.el.children.push({ type: 'text', text: t });
  };

  while (i < n) {
    const lt = src.indexOf('<', i);
    if (lt === -1) {
      addText(src.slice(i));
      break;
    }
    if (lt > i) addText(src.slice(i, lt));
    const c = src.charCodeAt(lt + 1);
    if (c === 63 /* ? */) {
      const end = src.indexOf('?>', lt + 2);
      i = end === -1 ? n : end + 2;
      continue;
    }
    if (c === 33 /* ! */) {
      if (src.startsWith('<!--', lt)) {
        const end = src.indexOf('-->', lt + 4);
        i = end === -1 ? n : end + 3;
      } else if (src.startsWith('<![CDATA[', lt)) {
        const end = src.indexOf(']]>', lt + 9);
        addText(src.slice(lt + 9, end === -1 ? n : end), false);
        i = end === -1 ? n : end + 3;
      } else {
        let j = lt + 2;
        let depth = 0;
        while (j < n) {
          const ch = src[j];
          if (ch === '[') depth++;
          else if (ch === ']') depth--;
          else if (ch === '>' && depth <= 0) break;
          j++;
        }
        i = j + 1;
      }
      continue;
    }
    if (c === 47 /* / */) {
      const end = src.indexOf('>', lt + 2);
      if (end === -1) throw new Error('Unterminated end tag');
      if (!stack.length) throw new Error('Unexpected end tag');
      stack.pop();
      i = end + 1;
      continue;
    }

    // Start tag.
    let j = lt + 1;
    while (j < n) {
      const ch = src.charCodeAt(j);
      if (ch === 32 || ch === 9 || ch === 10 || ch === 13 || ch === 47 || ch === 62) break;
      j++;
    }
    const rawName = src.slice(lt + 1, j);
    const rawAttrs: [string, string][] = [];
    let selfClosing = false;
    while (j < n) {
      let ch = src.charCodeAt(j);
      while (ch === 32 || ch === 9 || ch === 10 || ch === 13) ch = src.charCodeAt(++j);
      if (ch === 62 /* > */) {
        j++;
        break;
      }
      if (ch === 47 /* / */) {
        selfClosing = true;
        const end = src.indexOf('>', j);
        j = end === -1 ? n : end + 1;
        break;
      }
      const nameStart = j;
      while (j < n) {
        ch = src.charCodeAt(j);
        if (ch === 61 || ch === 32 || ch === 9 || ch === 10 || ch === 13 || ch === 62 || ch === 47) break;
        j++;
      }
      const attrName = src.slice(nameStart, j);
      while (src.charCodeAt(j) !== 61 && j < n && src.charCodeAt(j) !== 62) j++;
      if (src.charCodeAt(j) !== 61) {
        rawAttrs.push([attrName, '']);
        continue;
      }
      j++;
      while (j < n && /\s/.test(src[j])) j++;
      const quote = src[j];
      if (quote !== '"' && quote !== "'") throw new Error(`Malformed attribute ${attrName}`);
      const valEnd = src.indexOf(quote, j + 1);
      if (valEnd === -1) throw new Error('Unterminated attribute value');
      rawAttrs.push([attrName, decodeEntities(src.slice(j + 1, valEnd))]);
      j = valEnd + 1;
    }
    i = j;

    const parent = stack[stack.length - 1];
    let scope: Scope = parent ? parent.scope : rootScope;
    for (const [name, value] of rawAttrs) {
      if (name === 'xmlns' || name.startsWith('xmlns:')) {
        if (scope === (parent ? parent.scope : rootScope)) scope = new Map(scope);
        const prefix = name === 'xmlns' ? '' : name.slice(6);
        scope.set(prefix, value);
        if (prefix) mapper.prefixForUri(value, prefix);
      }
    }
    const attrs: Record<string, string> = {};
    for (const [name, value] of rawAttrs) {
      if (normalize && (name === 'xmlns' || name.startsWith('xmlns:'))) continue;
      const resolved = mapper.resolve(name, scope, true);
      if (MC_PREFIX_LIST_ATTRS.has(resolved)) attrs[resolved] = mapper.translateList(value, scope, false);
      else if (MC_QNAME_LIST_ATTRS.has(resolved)) attrs[resolved] = mapper.translateList(value, scope, true);
      else attrs[resolved] = value;
    }
    const el: XmlElement = { type: 'element', name: mapper.resolve(rawName, scope, false), attrs, children: [] };
    if (parent) parent.el.children.push(el);
    else if (!root) root = el;
    else throw new Error('Multiple root elements');
    if (!selfClosing) {
      const preserve =
        attrs['xml:space'] === 'preserve' || TEXT_ELEMENTS.has(el.name) || (parent?.preserve ?? false);
      stack.push({ el, scope, preserve });
    }
  }
  if (!root) throw new Error('Empty XML document');
  return { root, extraNamespaces: mapper.extra };
}

// eslint-disable-next-line no-control-regex
const INVALID_XML_CHARS = /[\u0000-\u0008\u000B\u000C\u000E-\u001F￾￿]|[\uD800-\uDBFF](?![\uDC00-\uDFFF])|(?<![\uD800-\uDBFF])[\uDC00-\uDFFF]/g;

export function stripInvalidXmlChars(s: string): string {
  return s.replace(INVALID_XML_CHARS, '');
}

export function escapeText(s: string): string {
  s = stripInvalidXmlChars(s);
  if (!/[&<>]/.test(s)) return s;
  return s.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
}

export function escapeAttr(s: string): string {
  s = stripInvalidXmlChars(s);
  if (!/[&<>"\t\n\r]/.test(s)) return s;
  return s
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/\t/g, '&#9;')
    .replace(/\n/g, '&#10;')
    .replace(/\r/g, '&#13;');
}

function writeNode(node: XmlNode, out: string[]): void {
  if (node.type === 'text') {
    out.push(escapeText(node.text));
    return;
  }
  out.push('<', node.name);
  for (const key in node.attrs) out.push(' ', key, '="', escapeAttr(node.attrs[key]), '"');
  if (!node.children.length) {
    out.push('/>');
    return;
  }
  out.push('>');
  for (const child of node.children) writeNode(child, out);
  out.push('</', node.name, '>');
}

/** Serialize a fragment (no declarations added). */
export function serializeNode(node: XmlNode): string {
  const out: string[] = [];
  writeNode(node, out);
  return out.join('');
}

export const XML_DECL = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n';

export interface SerializeOptions {
  declaration?: boolean;
  /** Namespace declarations to add to the root element, prefix -> URI ('' for default). */
  namespaces?: Record<string, string>;
}

export function serializeXml(root: XmlElement, opts: SerializeOptions = {}): string {
  const out: string[] = [];
  if (opts.declaration !== false) out.push(XML_DECL);
  const nsAttrs: Record<string, string> = {};
  for (const [prefix, uri] of Object.entries(opts.namespaces ?? {})) {
    nsAttrs[prefix ? `xmlns:${prefix}` : 'xmlns'] = uri;
  }
  writeNode({ ...root, attrs: { ...nsAttrs, ...root.attrs } }, out);
  return out.join('');
}

/** Collect the prefixes used by element and attribute names in a subtree. */
export function collectPrefixes(node: XmlNode, into = new Set<string>()): Set<string> {
  if (node.type !== 'element') return into;
  const idx = node.name.indexOf(':');
  if (idx > 0) into.add(node.name.slice(0, idx));
  for (const key in node.attrs) {
    const k = key.indexOf(':');
    if (k > 0) into.add(key.slice(0, k));
  }
  for (const value of [node.attrs['mc:Ignorable']]) {
    if (value) for (const p of value.split(/\s+/)) if (p) into.add(p);
  }
  for (const c of node.children) collectPrefixes(c, into);
  return into;
}

/** Namespace declarations needed by a subtree: canonical ones plus `extra` for unknown prefixes. */
export function namespacesFor(node: XmlNode, extra: Record<string, string> = {}, includeDefault = true): Record<string, string> {
  const out: Record<string, string> = {};
  if (includeDefault) out[''] = NS[''];
  for (const p of collectPrefixes(node)) {
    if (p === 'xml' || p === 'xmlns') continue;
    const uri = NS[p] ?? extra[p];
    if (uri) out[p] = uri;
  }
  return out;
}

// ---------------------------------------------------------------------------------------------
// Construction and query helpers

export function el(name: string, attrs: Record<string, string | number | boolean | undefined | null> = {}, children: (XmlNode | null | undefined | false)[] = []): XmlElement {
  const a: Record<string, string> = {};
  for (const [k, v] of Object.entries(attrs)) {
    if (v === undefined || v === null || v === false) continue;
    a[k] = v === true ? '1' : String(v);
  }
  return { type: 'element', name, attrs: a, children: children.filter(Boolean) as XmlNode[] };
}

export function txt(text: string): XmlText {
  return { type: 'text', text };
}

export function elements(e: XmlElement | undefined): XmlElement[] {
  if (!e) return [];
  const out: XmlElement[] = [];
  for (const c of e.children) if (c.type === 'element') out.push(c);
  return out;
}

export function child(e: XmlElement | undefined, name: string): XmlElement | undefined {
  if (!e) return undefined;
  for (const c of e.children) if (c.type === 'element' && c.name === name) return c;
  return undefined;
}

export function childrenNamed(e: XmlElement | undefined, name: string): XmlElement[] {
  if (!e) return [];
  const out: XmlElement[] = [];
  for (const c of e.children) if (c.type === 'element' && c.name === name) out.push(c);
  return out;
}

export function attr(e: XmlElement | undefined, name: string): string | undefined {
  return e ? e.attrs[name] : undefined;
}

export function numAttr(e: XmlElement | undefined, name: string): number | undefined {
  const v = e?.attrs[name];
  if (v === undefined || v === '') return undefined;
  const n = Number(v);
  return Number.isFinite(n) ? n : undefined;
}

/** OOXML boolean attribute ('1', 'true', 'on'). */
export function boolAttr(e: XmlElement | undefined, name: string, dflt = false): boolean {
  const v = e?.attrs[name];
  if (v === undefined) return dflt;
  return v === '1' || v === 'true' || v === 'on';
}

/** `val` attribute of a child element (fonts use this a lot). */
export function childVal(e: XmlElement | undefined, name: string): string | undefined {
  return attr(child(e, name), 'val');
}

export function textContent(e: XmlNode): string {
  if (e.type === 'text') return e.text;
  let s = '';
  for (const c of e.children) s += textContent(c);
  return s;
}

export function findFirst(e: XmlElement, name: string): XmlElement | undefined {
  for (const c of e.children) {
    if (c.type !== 'element') continue;
    if (c.name === name) return c;
    const found = findFirst(c, name);
    if (found) return found;
  }
  return undefined;
}

export function findAll(e: XmlElement, name: string, out: XmlElement[] = []): XmlElement[] {
  for (const c of e.children) {
    if (c.type !== 'element') continue;
    if (c.name === name) out.push(c);
    findAll(c, name, out);
  }
  return out;
}

export function cloneXml<T extends XmlNode>(node: T): T {
  if (node.type === 'text') return { type: 'text', text: node.text } as T;
  return {
    type: 'element',
    name: node.name,
    attrs: { ...node.attrs },
    children: node.children.map(cloneXml),
  } as T;
}

const FRAGMENT_DECLS = Object.entries(NS)
  .filter(([p]) => p !== 'xml')
  .map(([p, uri]) => (p ? `xmlns:${p}="${uri}"` : `xmlns="${uri}"`))
  .join(' ');

/** Parse serialized fragments produced by `serializeNode` (canonical prefixes, no declarations). */
export function parseFragments(xml: string | null | undefined): XmlElement[] {
  if (!xml) return [];
  return elements(parseXml(`<frag ${FRAGMENT_DECLS}>${xml}</frag>`).root);
}

export function parseFragment(xml: string): XmlElement {
  const first = parseFragments(xml)[0];
  if (!first) throw new Error('Empty fragment');
  return first;
}
