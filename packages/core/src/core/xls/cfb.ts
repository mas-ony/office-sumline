/**
 * Reader for Compound File Binary containers (the "OLE2" files that hold Excel 97-2003
 * workbooks): sectors, the allocation tables, the directory and the mini stream.
 */

const ENDOFCHAIN = 0xfffffffe;
const FREESECT = 0xffffffff;

export interface CfbEntry {
  name: string;
  type: number;
  start: number;
  size: number;
}

export class CompoundFile {
  readonly entries: CfbEntry[] = [];
  private sectorSize: number;
  private miniSectorSize: number;
  private miniCutoff: number;
  private fat: number[] = [];
  private miniFat: number[] = [];
  private miniStream: Uint8Array = new Uint8Array(0);
  private view: DataView;

  constructor(private readonly data: Uint8Array) {
    if (data.length < 512 || data[0] !== 0xd0 || data[1] !== 0xcf || data[2] !== 0x11 || data[3] !== 0xe0) throw new Error('Not an Excel 97-2003 file.');
    this.view = new DataView(data.buffer, data.byteOffset, data.byteLength);
    const u16 = (o: number) => this.view.getUint16(o, true);
    const u32 = (o: number) => this.view.getUint32(o, true);
    this.sectorSize = 1 << u16(0x1e);
    this.miniSectorSize = 1 << u16(0x20);
    const fatSectors = u32(0x2c);
    const firstDir = u32(0x30);
    this.miniCutoff = u32(0x38) || 4096;
    const firstMiniFat = u32(0x3c);
    let difatSector = u32(0x44);
    const difatCount = u32(0x48);
    // The sector numbers of the FAT: 109 in the header, the rest in DIFAT sectors.
    const fatSectorList: number[] = [];
    for (let i = 0; i < 109 && fatSectorList.length < fatSectors; i++) {
      const s = u32(0x4c + i * 4);
      if (s !== FREESECT) fatSectorList.push(s);
    }
    const perDifat = this.sectorSize / 4 - 1;
    for (let k = 0; k < difatCount && difatSector !== ENDOFCHAIN && difatSector !== FREESECT; k++) {
      const off = this.sectorOffset(difatSector);
      for (let i = 0; i < perDifat && fatSectorList.length < fatSectors; i++) {
        const s = this.view.getUint32(off + i * 4, true);
        if (s !== FREESECT) fatSectorList.push(s);
      }
      difatSector = this.view.getUint32(off + perDifat * 4, true);
    }
    for (const s of fatSectorList) {
      const off = this.sectorOffset(s);
      for (let i = 0; i < this.sectorSize / 4 && off + i * 4 + 4 <= data.length; i++) this.fat.push(this.view.getUint32(off + i * 4, true));
    }
    // Directory.
    const dir = this.chain(firstDir);
    for (let off = 0; off + 128 <= dir.length; off += 128) {
      const dv = new DataView(dir.buffer, dir.byteOffset + off, 128);
      const nameLen = dv.getUint16(64, true);
      let name = '';
      for (let i = 0; i + 2 <= Math.min(64, nameLen) - 2; i += 2) name += String.fromCharCode(dv.getUint16(i, true));
      this.entries.push({ name, type: dv.getUint8(66), start: dv.getUint32(116, true), size: dv.getUint32(120, true) });
    }
    const root = this.entries[0];
    if (root) {
      this.miniStream = this.chain(root.start, root.size);
      const mf = firstMiniFat !== ENDOFCHAIN ? this.chain(firstMiniFat) : new Uint8Array(0);
      const mv = new DataView(mf.buffer, mf.byteOffset, mf.byteLength);
      for (let i = 0; i + 4 <= mf.length; i += 4) this.miniFat.push(mv.getUint32(i, true));
    }
  }

  private sectorOffset(s: number): number {
    return (s + 1) * this.sectorSize;
  }

  /** Bytes of a sector chain (at most `size` bytes when given). */
  private chain(start: number, size = Infinity): Uint8Array {
    const parts: Uint8Array[] = [];
    let total = 0;
    let s = start;
    const seen = new Set<number>();
    while (s !== ENDOFCHAIN && s !== FREESECT && s < this.fat.length + 1 && !seen.has(s) && total < size) {
      seen.add(s);
      const off = this.sectorOffset(s);
      if (off >= this.data.length) break;
      const part = this.data.subarray(off, Math.min(off + this.sectorSize, this.data.length));
      parts.push(part);
      total += part.length;
      s = this.fat[s] ?? ENDOFCHAIN;
    }
    return concat(parts, size);
  }

  private miniChain(start: number, size: number): Uint8Array {
    const parts: Uint8Array[] = [];
    let total = 0;
    let s = start;
    const seen = new Set<number>();
    while (s !== ENDOFCHAIN && s !== FREESECT && !seen.has(s) && total < size) {
      seen.add(s);
      const off = s * this.miniSectorSize;
      if (off >= this.miniStream.length) break;
      const part = this.miniStream.subarray(off, Math.min(off + this.miniSectorSize, this.miniStream.length));
      parts.push(part);
      total += part.length;
      s = this.miniFat[s] ?? ENDOFCHAIN;
    }
    return concat(parts, size);
  }

  /** A stream by name (case-insensitive), or null. */
  stream(...names: string[]): Uint8Array | null {
    const wanted = names.map((n) => n.toLowerCase());
    const e = this.entries.find((x) => x.type === 2 && wanted.includes(x.name.toLowerCase()));
    if (!e) return null;
    return e.size < this.miniCutoff ? this.miniChain(e.start, e.size) : this.chain(e.start, e.size);
  }
}

function concat(parts: Uint8Array[], size: number): Uint8Array {
  const total = Math.min(size, parts.reduce((a, p) => a + p.length, 0));
  const out = new Uint8Array(total);
  let at = 0;
  for (const p of parts) {
    if (at >= total) break;
    const n = Math.min(p.length, total - at);
    out.set(p.subarray(0, n), at);
    at += n;
  }
  return out;
}
