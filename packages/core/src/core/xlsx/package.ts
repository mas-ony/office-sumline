import { strFromU8, strToU8, unzipSync, zipSync, type Zippable } from 'fflate';
import { attr, el, elements, parseXml, serializeXml, type XmlElement } from '../xml/xml';

export interface Relationship {
  id: string;
  type: string;
  target: string;
  targetMode?: 'External';
}

const REL_BASE = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/';
const STRICT_REL_BASE = 'http://purl.oclc.org/ooxml/officeDocument/relationships/';

export const REL = {
  officeDocument: `${REL_BASE}officeDocument`,
  worksheet: `${REL_BASE}worksheet`,
  chartsheet: `${REL_BASE}chartsheet`,
  styles: `${REL_BASE}styles`,
  sharedStrings: `${REL_BASE}sharedStrings`,
  theme: `${REL_BASE}theme`,
  calcChain: `${REL_BASE}calcChain`,
  drawing: `${REL_BASE}drawing`,
  vmlDrawing: `${REL_BASE}vmlDrawing`,
  image: `${REL_BASE}image`,
  hyperlink: `${REL_BASE}hyperlink`,
  comments: `${REL_BASE}comments`,
  table: `${REL_BASE}table`,
  sheetMetadata: `${REL_BASE}sheetMetadata`,
  extendedProps: `${REL_BASE}extended-properties`,
  coreProps: 'http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties',
  threadedComment: 'http://schemas.microsoft.com/office/2017/10/relationships/threadedComment',
  person: 'http://schemas.microsoft.com/office/2017/10/relationships/person',
} as const;

/** Compare a relationship type with a known one, accepting the ISO 29500 strict variant. */
export function relTypeIs(type: string, known: string): boolean {
  if (type === known) return true;
  if (known.startsWith(REL_BASE)) return type === STRICT_REL_BASE + known.slice(REL_BASE.length);
  return false;
}

export const CT = {
  workbook: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml',
  workbookMacro: 'application/vnd.ms-excel.sheet.macroEnabled.main+xml',
  template: 'application/vnd.openxmlformats-officedocument.spreadsheetml.template.main+xml',
  worksheet: 'application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml',
  styles: 'application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml',
  sharedStrings: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml',
  theme: 'application/vnd.openxmlformats-officedocument.theme+xml',
  comments: 'application/vnd.openxmlformats-officedocument.spreadsheetml.comments+xml',
  threadedComments: 'application/vnd.ms-excel.threadedcomments+xml',
  person: 'application/vnd.ms-excel.person+xml',
  drawing: 'application/vnd.openxmlformats-officedocument.drawing+xml',
  vml: 'application/vnd.openxmlformats-officedocument.vmlDrawing',
  metadata: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheetMetadata+xml',
  calcChain: 'application/vnd.openxmlformats-officedocument.spreadsheetml.calcChain+xml',
  coreProps: 'application/vnd.openxmlformats-package.core-properties+xml',
  extendedProps: 'application/vnd.openxmlformats-officedocument.extended-properties+xml',
  rels: 'application/vnd.openxmlformats-package.relationships+xml',
  xml: 'application/xml',
} as const;

export const IMAGE_CONTENT_TYPES: Record<string, string> = {
  png: 'image/png',
  jpg: 'image/jpeg',
  jpeg: 'image/jpeg',
  gif: 'image/gif',
  bmp: 'image/bmp',
  tif: 'image/tiff',
  tiff: 'image/tiff',
  svg: 'image/svg+xml',
  webp: 'image/webp',
  emf: 'image/x-emf',
  wmf: 'image/x-wmf',
};

export function extensionForMime(mime: string): string {
  for (const [ext, type] of Object.entries(IMAGE_CONTENT_TYPES)) if (type === mime) return ext === 'jpeg' ? 'jpg' : ext;
  return 'png';
}

export function normalizePath(p: string): string {
  const parts: string[] = [];
  for (const seg of p.replace(/\\/g, '/').split('/')) {
    if (!seg || seg === '.') continue;
    if (seg === '..') parts.pop();
    else parts.push(seg);
  }
  return parts.join('/');
}

export function dirname(p: string): string {
  const i = p.lastIndexOf('/');
  return i < 0 ? '' : p.slice(0, i);
}

export function relsPathFor(partPath: string): string {
  const dir = dirname(partPath);
  const base = partPath.slice(dir ? dir.length + 1 : 0);
  return dir ? `${dir}/_rels/${base}.rels` : `_rels/${base}.rels`;
}

export class ContentTypes {
  defaults = new Map<string, string>();
  overrides = new Map<string, string>();

  static parse(xml: string): ContentTypes {
    const ct = new ContentTypes();
    const { root } = parseXml(xml, { normalize: false });
    for (const e of elements(root)) {
      const local = e.name.replace(/^.*:/, '');
      if (local === 'Default') ct.defaults.set((attr(e, 'Extension') ?? '').toLowerCase(), attr(e, 'ContentType') ?? '');
      else if (local === 'Override') ct.overrides.set(normalizePath(attr(e, 'PartName') ?? ''), attr(e, 'ContentType') ?? '');
    }
    return ct;
  }

  typeOf(path: string): string | undefined {
    const o = this.overrides.get(normalizePath(path));
    if (o) return o;
    const ext = path.slice(path.lastIndexOf('.') + 1).toLowerCase();
    return this.defaults.get(ext);
  }

  serialize(): string {
    const children: XmlElement[] = [];
    for (const [ext, type] of this.defaults) children.push(el('Default', { Extension: ext, ContentType: type }));
    for (const [part, type] of this.overrides) children.push(el('Override', { PartName: `/${part}`, ContentType: type }));
    return serializeXml(el('Types', {}, children), {
      namespaces: { '': 'http://schemas.openxmlformats.org/package/2006/content-types' },
    });
  }
}

/** An Open Packaging Conventions (ZIP) container. Paths are stored without a leading slash. */
export class OoxmlPackage {
  readonly files = new Map<string, Uint8Array>();
  contentTypes = new ContentTypes();
  private relsCache = new Map<string, Relationship[]>();

  static fromBytes(bytes: Uint8Array): OoxmlPackage {
    let entries: Record<string, Uint8Array>;
    try {
      entries = unzipSync(bytes);
    } catch (e) {
      throw new Error(`Not a valid XLSX (zip) file: ${(e as Error).message}`);
    }
    const pkg = new OoxmlPackage();
    for (const [name, data] of Object.entries(entries)) {
      if (name.endsWith('/')) continue;
      pkg.files.set(normalizePath(name), data);
    }
    const ct = pkg.findCaseInsensitive('[Content_Types].xml');
    if (!ct) throw new Error('Not a valid XLSX file: [Content_Types].xml is missing');
    pkg.contentTypes = ContentTypes.parse(strFromU8(pkg.files.get(ct)!));
    pkg.files.delete(ct);
    return pkg;
  }

  private findCaseInsensitive(path: string): string | undefined {
    if (this.files.has(path)) return path;
    const lower = path.toLowerCase();
    for (const k of this.files.keys()) if (k.toLowerCase() === lower) return k;
    return undefined;
  }

  /** Every part path in the package. */
  paths(): string[] {
    return [...this.files.keys()];
  }

  has(path: string): boolean {
    return this.findCaseInsensitive(normalizePath(path)) !== undefined;
  }

  get(path: string): Uint8Array | undefined {
    const key = this.findCaseInsensitive(normalizePath(path));
    return key ? this.files.get(key) : undefined;
  }

  getText(path: string): string | undefined {
    const data = this.get(path);
    return data ? strFromU8(data) : undefined;
  }

  set(path: string, data: Uint8Array, contentType?: string): void {
    const p = normalizePath(path);
    const existing = this.findCaseInsensitive(p);
    if (existing && existing !== p) this.files.delete(existing);
    this.files.set(p, data);
    if (contentType) this.contentTypes.overrides.set(p, contentType);
  }

  setText(path: string, text: string, contentType?: string): void {
    this.set(path, strToU8(text), contentType);
  }

  delete(path: string): void {
    const key = this.findCaseInsensitive(normalizePath(path));
    if (!key) return;
    this.files.delete(key);
    this.contentTypes.overrides.delete(key);
    this.relsCache.delete(key);
    this.files.delete(relsPathFor(key));
  }

  ensureDefaultContentType(ext: string, type: string): void {
    const e = ext.toLowerCase();
    if (!this.contentTypes.defaults.has(e)) this.contentTypes.defaults.set(e, type);
  }

  /** Relationships of a part ('' for the package-level relationships). */
  getRels(partPath: string): Relationship[] {
    const key = partPath ? normalizePath(partPath) : '';
    const cached = this.relsCache.get(key);
    if (cached) return cached;
    const relsPath = key ? relsPathFor(key) : '_rels/.rels';
    const text = this.getText(relsPath);
    const rels: Relationship[] = [];
    if (text) {
      const { root } = parseXml(text, { normalize: false });
      for (const r of elements(root)) {
        if (r.name.replace(/^.*:/, '') !== 'Relationship') continue;
        rels.push({
          id: attr(r, 'Id') ?? '',
          type: attr(r, 'Type') ?? '',
          target: attr(r, 'Target') ?? '',
          targetMode: attr(r, 'TargetMode') === 'External' ? 'External' : undefined,
        });
      }
    }
    this.relsCache.set(key, rels);
    return rels;
  }

  setRels(partPath: string, rels: Relationship[]): void {
    const key = partPath ? normalizePath(partPath) : '';
    this.relsCache.set(key, rels);
    const relsPath = key ? relsPathFor(key) : '_rels/.rels';
    if (!rels.length) {
      this.files.delete(relsPath);
      return;
    }
    const xml = serializeXml(
      el(
        'Relationships',
        {},
        rels.map((r) => el('Relationship', { Id: r.id, Type: r.type, Target: r.target, TargetMode: r.targetMode })),
      ),
      { namespaces: { '': 'http://schemas.openxmlformats.org/package/2006/relationships' } },
    );
    this.files.set(relsPath, strToU8(xml));
  }

  /** Resolve a relationship target of `fromPart` to an absolute package path. */
  resolveTarget(fromPart: string, target: string): string {
    if (target.startsWith('/')) return normalizePath(target);
    return normalizePath(`${dirname(fromPart)}/${target}`);
  }

  /** Target string for a relationship from `fromPart` to `toPart`. */
  relativeTarget(fromPart: string, toPart: string): string {
    const fromDir = dirname(fromPart).split('/').filter(Boolean);
    const to = normalizePath(toPart).split('/');
    let i = 0;
    while (i < fromDir.length && i < to.length - 1 && fromDir[i] === to[i]) i++;
    return [...fromDir.slice(i).map(() => '..'), ...to.slice(i)].join('/');
  }

  findRel(partPath: string, type: string): Relationship | undefined {
    return this.getRels(partPath).find((r) => relTypeIs(r.type, type));
  }

  findRels(partPath: string, type: string): Relationship[] {
    return this.getRels(partPath).filter((r) => relTypeIs(r.type, type));
  }

  /** Path of the workbook part (usually xl/workbook.xml). */
  workbookPath(): string {
    const rel = this.findRel('', REL.officeDocument);
    if (rel) return this.resolveTarget('', rel.target);
    for (const [p, t] of this.contentTypes.overrides) if (t === CT.workbook || t === CT.template || t === CT.workbookMacro) return p;
    if (this.has('xl/workbook.xml')) return 'xl/workbook.xml';
    throw new Error('Not a valid XLSX file: no workbook part');
  }

  /** Allocate a relationship id that is unused for the part. */
  nextRelId(partPath: string): string {
    const used = new Set(this.getRels(partPath).map((r) => r.id));
    let n = used.size + 1;
    while (used.has(`rId${n}`)) n++;
    return `rId${n}`;
  }

  /** Add a relationship and return its id. */
  addRel(partPath: string, type: string, target: string, external = false): string {
    const rels = [...this.getRels(partPath)];
    const id = this.nextRelId(partPath);
    rels.push({ id, type, target, targetMode: external ? 'External' : undefined });
    this.setRels(partPath, rels);
    return id;
  }

  /** A free part name such as xl/drawings/drawing3.xml. */
  uniquePartName(dir: string, base: string, ext: string): string {
    let n = 1;
    while (this.has(`${dir}/${base}${n}.${ext}`)) n++;
    return `${dir}/${base}${n}.${ext}`;
  }

  toBytes(): Uint8Array {
    const zippable: Zippable = {};
    zippable['[Content_Types].xml'] = [strToU8(this.contentTypes.serialize()), { level: 6 }];
    const names = [...this.files.keys()].sort((a, b) => {
      // Package relationships first, then the workbook, then the rest alphabetically.
      const rank = (p: string) => (p === '_rels/.rels' ? 0 : p.startsWith('xl/workbook') ? 1 : 2);
      return rank(a) - rank(b) || a.localeCompare(b);
    });
    for (const name of names) {
      const data = this.files.get(name)!;
      const isMedia = /\.(png|jpe?g|gif|webp)$/i.test(name);
      // Very large parts (big sheets) use fast compression: twice as fast, slightly larger files.
      zippable[name] = [data, { level: isMedia ? 0 : data.length > 2_000_000 ? 1 : 6 }];
    }
    return zipSync(zippable);
  }

  clone(): OoxmlPackage {
    const pkg = new OoxmlPackage();
    for (const [k, v] of this.files) pkg.files.set(k, v);
    pkg.contentTypes.defaults = new Map(this.contentTypes.defaults);
    pkg.contentTypes.overrides = new Map(this.contentTypes.overrides);
    return pkg;
  }
}
