/**
 * Small hand-built Excel 97-2003 (.xls) and Excel binary (.xlsb) workbooks for the reader tests.
 * They follow the published record layouts ([MS-XLS], [MS-XLSB], [MS-CFB]).
 */
import { strToU8, zipSync } from 'fflate';

class Bytes {
  private parts: number[] = [];
  u8(v: number) {
    this.parts.push(v & 0xff);
    return this;
  }
  u16(v: number) {
    return this.u8(v).u8(v >> 8);
  }
  u32(v: number) {
    return this.u16(v & 0xffff).u16((v >>> 16) & 0xffff);
  }
  i32(v: number) {
    return this.u32(v >>> 0);
  }
  f64(v: number) {
    const b = new DataView(new ArrayBuffer(8));
    b.setFloat64(0, v, true);
    for (let i = 0; i < 8; i++) this.u8(b.getUint8(i));
    return this;
  }
  bytes(b: ArrayLike<number>) {
    for (let i = 0; i < b.length; i++) this.u8(b[i]);
    return this;
  }
  /** 8-bit characters with a count and a flag byte (BIFF8 strings). */
  str8(s: string, cchSize: 1 | 2) {
    if (cchSize === 1) this.u8(s.length);
    else this.u16(s.length);
    this.u8(0);
    for (const ch of s) this.u8(ch.charCodeAt(0));
    return this;
  }
  /** 32-bit count of UTF-16 characters (XLSB strings). */
  wide(s: string) {
    this.u32(s.length);
    for (let i = 0; i < s.length; i++) this.u16(s.charCodeAt(i));
    return this;
  }
  get length() {
    return this.parts.length;
  }
  toArray() {
    return Uint8Array.from(this.parts);
  }
}

const b = () => new Bytes();

// -------------------------------------------------------------------------------------------------
// Compound file with a single "Workbook" stream (regular sectors only).

function compoundFile(stream: Uint8Array): Uint8Array {
  const SECTOR = 512;
  const padded = new Uint8Array(Math.max(4096, Math.ceil(stream.length / SECTOR) * SECTOR));
  padded.set(stream);
  const streamSectors = padded.length / SECTOR;
  const header = b().bytes([0xd0, 0xcf, 0x11, 0xe0, 0xa1, 0xb1, 0x1a, 0xe1]).bytes(new Array(16).fill(0)).u16(0x3e).u16(3).u16(0xfffe).u16(9).u16(6).bytes(new Array(6).fill(0))
    .u32(0).u32(1).u32(1).u32(0).u32(4096).u32(0xfffffffe).u32(0).u32(0xfffffffe).u32(0).u32(0);
  for (let i = 1; i < 109; i++) header.u32(0xffffffff);
  const fat = b().u32(0xfffffffd).u32(0xfffffffe);
  for (let i = 0; i < streamSectors; i++) fat.u32(i === streamSectors - 1 ? 0xfffffffe : 3 + i);
  while (fat.length < SECTOR) fat.u32(0xffffffff);
  const entry = (name: string, type: number, child: number, start: number, size: number) => {
    const e = b();
    for (let i = 0; i < 32; i++) e.u16(i < name.length ? name.charCodeAt(i) : 0);
    e.u16(name ? (name.length + 1) * 2 : 0).u8(type).u8(1).u32(0xffffffff).u32(0xffffffff).u32(child).bytes(new Array(16).fill(0)).u32(0).bytes(new Array(16).fill(0)).u32(start).u32(size).u32(0);
    return e.toArray();
  };
  // Streams under 4096 bytes belong in the mini stream: the padded size keeps this one in sectors.
  const dir = b().bytes(entry('Root Entry', 5, 1, 0xfffffffe, 0)).bytes(entry('Workbook', 2, 0xffffffff, 2, padded.length)).bytes(entry('', 0, 0xffffffff, 0, 0)).bytes(entry('', 0, 0xffffffff, 0, 0));
  return b().bytes(header.toArray()).bytes(fat.toArray()).bytes(dir.toArray()).bytes(padded).toArray();
}

const rec = (type: number, body: Bytes | Uint8Array) => {
  const data = body instanceof Bytes ? body.toArray() : body;
  return b().u16(type).u16(data.length).bytes(data).toArray();
};

function xf(font: number, fmt: number, style: boolean, extra: { fill?: number; bold?: boolean; border?: boolean; wrap?: boolean } = {}): Bytes {
  const borderBits = extra.border ? 1 | (1 << 4) | (1 << 8) | (1 << 12) : 0;
  return b()
    .u16(font)
    .u16(fmt)
    .u16(style ? 0xfff5 : 0x0001)
    .u8(extra.wrap ? 0x08 | 0x20 : 0x20)
    .u8(0)
    .u8(0)
    .u8(style ? 0 : 0xf8)
    .u32(borderBits)
    .u32(extra.fill !== undefined ? 1 << 26 : 0)
    .u16(extra.fill !== undefined ? extra.fill | (65 << 7) : 64 | (65 << 7));
}

/**
 * Two sheets. Sheet1: A1 "Region" (SST), B1 12.5, C1 7 (RK int), D1 0.25 (RK x100), E1 TRUE, F1
 * #DIV/0!, A2 =B1*2, B2 =SUM(B1:D1), C2 ="a"&"b", D2 =Sheet2!A1, E2 =IF(B1>1,"big","small"), A3:A4
 * shared formula =A2+1, A6:B6 merged, bold red-filled B1, column B 20 wide, row 1 30pt, first row
 * frozen; a name MyRange = Sheet1!$A$1:$B$2. Sheet2: A1 "Other" (SST string split by CONTINUE).
 */
export function buildXls(): Uint8Array {
  const globals: Uint8Array[] = [];
  globals.push(rec(0x0809, b().u16(0x0600).u16(0x0005).u16(0).u16(0).u32(0).u32(0)));
  globals.push(rec(0x0042, b().u16(1200)));
  globals.push(rec(0x0022, b().u16(0)));
  const font = (name: string, bold: boolean, color = 0x7fff) => rec(0x0031, b().u16(200).u16(0).u16(color).u16(bold ? 700 : 400).u16(0).u8(0).u8(0).u8(0).u8(0).str8(name, 1));
  globals.push(font('Arial', false), font('Arial', true), font('Arial', false), font('Arial', false), font('Arial', true, 10));
  globals.push(rec(0x041e, b().u16(164).str8('0.0%', 2)));
  for (let i = 0; i < 15; i++) globals.push(rec(0x00e0, xf(0, 0, true)));
  globals.push(rec(0x00e0, xf(0, 0, false))); // 15: default cell XF
  globals.push(rec(0x00e0, xf(5, 0, false, { fill: 10, border: true }))); // 16: bold red text (font index 5: index 4 does not exist), red fill, thin borders
  globals.push(rec(0x00e0, xf(0, 164, false, { wrap: true }))); // 17: 0.0% wrapped
  const boundAt: number[] = [];
  const bound = (name: string) => {
    boundAt.push(globals.reduce((a, g) => a + g.length, 0) + 4);
    return rec(0x0085, b().u32(0).u8(0).u8(0).str8(name, 1));
  };
  globals.push(bound('Sheet1'));
  globals.push(bound('Sheet2'));
  globals.push(rec(0x01ae, b().u16(2).u16(0x0401)));
  globals.push(rec(0x0017, b().u16(2).u16(0).u16(0).u16(0).u16(0).u16(1).u16(1)));
  // NAME MyRange: ptgArea3d (0x3b) ixti 0, rows 0..1, cols 0..1 absolute.
  const nameFormula = b().u8(0x3b).u16(0).u16(0).u16(1).u16(0).u16(1).toArray();
  globals.push(rec(0x0018, b().u16(0).u8(0).u8(7).u16(nameFormula.length).u16(0).u16(0).u32(0).u8(0).bytes([...'MyRange'].map((c) => c.charCodeAt(0))).bytes(nameFormula)));
  // SST: "Region" and "Other", the second split across a CONTINUE record (new flag byte there).
  const sstBody = b().u32(2).u32(2).str8('Region', 2).u16(5).u8(0).bytes([79, 116]);
  globals.push(rec(0x00fc, sstBody), rec(0x003c, b().u8(0).bytes([104, 101, 114])));
  globals.push(rec(0x000a, b()));

  const cell = (type: number, row: number, col: number, ixfe: number, rest: Bytes) => rec(type, b().u16(row).u16(col).u16(ixfe).bytes(rest.toArray()));
  const formula = (row: number, col: number, value: Bytes, rgce: Bytes, ixfe = 15) => {
    const tokens = rgce.toArray();
    return rec(0x0006, b().u16(row).u16(col).u16(ixfe).bytes(value.toArray()).u16(0).u32(0).u16(tokens.length).bytes(tokens));
  };
  const num = (v: number) => b().f64(v);
  const special = (kind: number, val = 0) => b().u8(kind).u8(0).u8(val).u8(0).u8(0).u8(0).u16(0xffff);
  const sheet1: Uint8Array[] = [];
  sheet1.push(rec(0x0809, b().u16(0x0600).u16(0x0010).u16(0).u16(0).u32(0).u32(0)));
  sheet1.push(rec(0x0055, b().u16(8)));
  sheet1.push(rec(0x007d, b().u16(1).u16(1).u16(20 * 256).u16(15).u16(2).u16(0)));
  sheet1.push(rec(0x0208, b().u16(0).u16(0).u16(6).u16(600).u16(0).u16(0).u16(0x40).u16(15)));
  sheet1.push(cell(0x00fd, 0, 0, 15, b().u32(0)));
  sheet1.push(cell(0x0203, 0, 1, 16, b().f64(12.5)));
  sheet1.push(cell(0x027e, 0, 2, 15, b().u32((7 << 2) | 2)));
  sheet1.push(cell(0x027e, 0, 3, 17, b().u32((25 << 2) | 3)));
  sheet1.push(cell(0x0205, 0, 4, 15, b().u8(1).u8(0)));
  sheet1.push(cell(0x0205, 0, 5, 15, b().u8(0x07).u8(1)));
  sheet1.push(formula(1, 0, num(25), b().u8(0x24).u16(0).u16(0xc001).u8(0x1e).u16(2).u8(0x05)));
  sheet1.push(formula(1, 1, num(19.75), b().u8(0x25).u16(0).u16(0).u16(0xc001).u16(0xc003).u8(0x22).u8(1).u16(4)));
  sheet1.push(formula(1, 2, special(0), b().u8(0x17).u8(1).u8(0).u8(97).u8(0x17).u8(1).u8(0).u8(98).u8(0x08)));
  sheet1.push(rec(0x0207, b().str8('ab', 2)));
  sheet1.push(formula(1, 3, special(0), b().u8(0x3a).u16(1).u16(0).u16(0)));
  sheet1.push(rec(0x0207, b().str8('Other', 2)));
  sheet1.push(formula(1, 4, special(0), b().u8(0x24).u16(0).u16(0xc001).u8(0x1e).u16(1).u8(0x0d).u8(0x17).u8(3).u8(0).bytes([98, 105, 103]).u8(0x17).u8(5).u8(0).bytes([115, 109, 97, 108, 108]).u8(0x42).u8(3).u16(1)));
  sheet1.push(rec(0x0207, b().str8('big', 2)));
  // Shared formula A3:A4 = (cell above) + 1, anchored at A3.
  sheet1.push(formula(2, 0, num(26), b().u8(0x01).u16(2).u16(0)));
  sheet1.push(rec(0x04bc, b().u16(2).u16(3).u8(0).u8(0).u8(0).u8(2).u16(9).u8(0x2c).u16(0xffff).u16(0xc000).u8(0x1e).u16(1).u8(0x03)));
  sheet1.push(formula(3, 0, num(27), b().u8(0x01).u16(2).u16(0)));
  sheet1.push(rec(0x00e5, b().u16(1).u16(5).u16(5).u16(0).u16(1)));
  sheet1.push(rec(0x023e, b().u16(0x0002 | 0x0008 | 0x0100).u16(0).u16(0).u16(64).u16(0).u16(0).u16(0).u32(0)));
  sheet1.push(rec(0x0041, b().u16(0).u16(1).u16(1).u16(0).u8(2)));
  sheet1.push(rec(0x000a, b()));
  const sheet2: Uint8Array[] = [];
  sheet2.push(rec(0x0809, b().u16(0x0600).u16(0x0010).u16(0).u16(0).u32(0).u32(0)));
  sheet2.push(cell(0x00fd, 0, 0, 15, b().u32(1)));
  sheet2.push(rec(0x000a, b()));
  const join = (parts: Uint8Array[]) => {
    const out = new Uint8Array(parts.reduce((a, p) => a + p.length, 0));
    let at = 0;
    for (const p of parts) {
      out.set(p, at);
      at += p.length;
    }
    return out;
  };
  const g = join(globals);
  const s1 = join(sheet1);
  const s2 = join(sheet2);
  const stream = join([g, s1, s2]);
  const view = new DataView(stream.buffer);
  view.setUint32(boundAt[0], g.length, true);
  view.setUint32(boundAt[1], g.length + s1.length, true);
  return compoundFile(stream);
}

// -------------------------------------------------------------------------------------------------
// XLSB

function brt(type: number, body: Bytes): number[] {
  const data = body.toArray();
  const out: number[] = [];
  if (type < 0x80) out.push(type);
  else out.push((type & 0x7f) | 0x80, type >> 7);
  let size = data.length;
  do {
    let byte = size & 0x7f;
    size >>= 7;
    if (size) byte |= 0x80;
    out.push(byte);
  } while (size);
  return [...out, ...data];
}

const color = (kind: number, index: number, rgb: [number, number, number] = [0, 0, 0]) => b().u8((kind << 1) | 1).u8(index).u16(0).u8(rgb[0]).u8(rgb[1]).u8(rgb[2]).u8(255);

/**
 * One sheet "Data": A1 "Name" (shared string), B1 3.5, C1 42 (RK), D1 TRUE, A2 ="x" (string
 * formula result), B2 =B1*2, C2 =SUM(B1:C1), A3:B3 merged, B1 bold with a yellow fill, column B
 * hidden, two frozen rows; a name Total = Data!$B$2.
 */
export function buildXlsb(): Uint8Array {
  const wb = [
    ...brt(153, b().u32(0).u32(0).wide('')),
    ...brt(156, b().u32(0).u32(1).wide('rId1').wide('Data')),
    ...brt(355, b()),
    ...brt(362, b().u32(1).u32(0).i32(0).i32(0)),
    ...brt(39, b().u32(0).u8(0).u32(0xffffffff).wide('Total').u32(9).u8(0x3a).u16(0).u32(1).u16(1).u32(0)),
  ];
  const sst = [...brt(19, b().u8(0).wide('Name'))];
  const styles = [
    ...brt(44, b().u16(164).wide('0.000')),
    ...brt(43, b().u16(220).u16(0).u16(400).u16(0).u8(0).u8(2).u8(0).u8(0).bytes(color(3, 1).toArray()).u8(2).wide('Calibri')),
    ...brt(43, b().u16(220).u16(0).u16(700).u16(0).u8(0).u8(2).u8(0).u8(0).bytes(color(2, 0, [0x12, 0x34, 0x56]).toArray()).u8(2).wide('Calibri')),
    ...brt(45, b().u32(0).bytes(color(1, 64).toArray()).bytes(color(1, 65).toArray())),
    ...brt(45, b().u32(1).bytes(color(2, 0, [255, 255, 0]).toArray()).bytes(color(1, 65).toArray())),
    ...brt(46, b().u8(0).bytes(b().u8(0).u8(0).bytes(color(1, 64).toArray()).toArray()).bytes(b().u8(0).u8(0).bytes(color(1, 64).toArray()).toArray()).bytes(b().u8(0).u8(0).bytes(color(1, 64).toArray()).toArray()).bytes(b().u8(0).u8(0).bytes(color(1, 64).toArray()).toArray()).bytes(b().u8(0).u8(0).bytes(color(1, 64).toArray()).toArray())),
    ...brt(617, b().u32(2)),
    ...brt(47, b().u16(0xffff).u16(0).u16(0).u16(0).u16(0).u8(0).u8(0).u16(0x1010).u8(0).u8(0)),
    ...brt(47, b().u16(0).u16(164).u16(1).u16(1).u16(0).u8(0).u8(0).u16(0x1010 | 0x0002).u8(0).u8(0)),
    ...brt(618, b()),
  ];
  const cellHead = (col: number, style: number) => b().u32(col).u32(style);
  const formula = (rgce: Bytes) => {
    const t = rgce.toArray();
    return b().u32(t.length).bytes(t).u32(0);
  };
  const sheet = [
    ...brt(485, b().u32(0xffffffff).u16(8).u16(300)),
    ...brt(60, b().u32(1).u32(1).u32(15 * 256).u32(0).u16(1)),
    ...brt(137, b().u16(4).u32(0).u32(0).u32(0).u8(64).u8(0).u16(100)),
    ...brt(151, b().f64(0).f64(2).u32(2).u32(0).u32(2).u8(1)),
    ...brt(0, b().u32(0).u32(0).u16(300).u8(0).u8(0).u8(0).u32(0)),
    ...brt(7, cellHead(0, 0).u32(0)),
    ...brt(5, cellHead(1, 1).f64(3.5)),
    ...brt(2, cellHead(2, 0).u32((42 << 2) | 2)),
    ...brt(4, cellHead(3, 0).u8(1)),
    ...brt(0, b().u32(1).u32(0).u16(300).u8(0).u8(0).u8(0).u32(0)),
    ...brt(8, cellHead(0, 0).wide('x').u16(0).bytes(formula(b().u8(0x17).u16(1).u16(120)).toArray())),
    ...brt(9, cellHead(1, 0).f64(7).u16(0).bytes(formula(b().u8(0x24).u32(0).u16(0xc001).u8(0x1e).u16(2).u8(0x05)).toArray())),
    ...brt(9, cellHead(2, 0).f64(45.5).u16(0).bytes(formula(b().u8(0x25).u32(0).u32(0).u16(0xc001).u16(0xc002).u8(0x22).u8(1).u16(4)).toArray())),
    ...brt(176, b().u32(2).u32(2).u32(0).u32(1)),
  ];
  const ct = `<?xml version="1.0" encoding="UTF-8"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="bin" ContentType="application/vnd.ms-excel.sheet.binary.macroEnabled.main"/></Types>`;
  const rels = `<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.bin"/></Relationships>`;
  const wbRels = `<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.bin"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings" Target="sharedStrings.bin"/><Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.bin"/></Relationships>`;
  return zipSync({
    '[Content_Types].xml': strToU8(ct),
    '_rels/.rels': strToU8(rels),
    'xl/workbook.bin': Uint8Array.from(wb),
    'xl/_rels/workbook.bin.rels': strToU8(wbRels),
    'xl/sharedStrings.bin': Uint8Array.from(sst),
    'xl/styles.bin': Uint8Array.from(styles),
    'xl/worksheets/sheet1.bin': Uint8Array.from(sheet),
  });
}
