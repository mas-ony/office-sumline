/**
 * Formulas stored as parsed tokens ("Ptg" streams) in Excel 97-2003 (BIFF8) and Excel binary
 * (XLSB / BIFF12) workbooks, turned back into formula text. The two formats share the token
 * set; they differ in the size of row, column and string fields.
 */
import { colName } from '../refs';
import { quoteSheetName } from '../refs';

/** Built-in function names by index, with their fixed argument count (-1: variable). */
const FUNCS: Record<number, [string, number]> = {
  0: ['COUNT', -1], 1: ['IF', -1], 2: ['ISNA', 1], 3: ['ISERROR', 1], 4: ['SUM', -1], 5: ['AVERAGE', -1], 6: ['MIN', -1], 7: ['MAX', -1],
  8: ['ROW', -1], 9: ['COLUMN', -1], 10: ['NA', 0], 11: ['NPV', -1], 12: ['STDEV', -1], 13: ['DOLLAR', -1], 14: ['FIXED', -1], 15: ['SIN', 1],
  16: ['COS', 1], 17: ['TAN', 1], 18: ['ATAN', 1], 19: ['PI', 0], 20: ['SQRT', 1], 21: ['EXP', 1], 22: ['LN', 1], 23: ['LOG10', 1], 24: ['ABS', 1],
  25: ['INT', 1], 26: ['SIGN', 1], 27: ['ROUND', 2], 28: ['LOOKUP', -1], 29: ['INDEX', -1], 30: ['REPT', 2], 31: ['MID', 3], 32: ['LEN', 1],
  33: ['VALUE', 1], 34: ['TRUE', 0], 35: ['FALSE', 0], 36: ['AND', -1], 37: ['OR', -1], 38: ['NOT', 1], 39: ['MOD', 2], 40: ['DCOUNT', 3],
  41: ['DSUM', 3], 42: ['DAVERAGE', 3], 43: ['DMIN', 3], 44: ['DMAX', 3], 45: ['DSTDEV', 3], 46: ['VAR', -1], 47: ['DVAR', 3], 48: ['TEXT', 2],
  49: ['LINEST', -1], 50: ['TREND', -1], 51: ['LOGEST', -1], 52: ['GROWTH', -1], 56: ['PV', -1], 57: ['FV', -1], 58: ['NPER', -1], 59: ['PMT', -1],
  60: ['RATE', -1], 61: ['MIRR', 3], 62: ['IRR', -1], 63: ['RAND', 0], 64: ['MATCH', -1], 65: ['DATE', 3], 66: ['TIME', 3], 67: ['DAY', 1],
  68: ['MONTH', 1], 69: ['YEAR', 1], 70: ['WEEKDAY', -1], 71: ['HOUR', 1], 72: ['MINUTE', 1], 73: ['SECOND', 1], 74: ['NOW', 0], 75: ['AREAS', 1],
  76: ['ROWS', 1], 77: ['COLUMNS', 1], 78: ['OFFSET', -1], 82: ['SEARCH', -1], 83: ['TRANSPOSE', 1], 86: ['TYPE', 1], 97: ['ATAN2', 2], 98: ['ASIN', 1],
  99: ['ACOS', 1], 100: ['CHOOSE', -1], 101: ['HLOOKUP', -1], 102: ['VLOOKUP', -1], 105: ['ISREF', 1], 109: ['LOG', -1], 111: ['CHAR', 1],
  112: ['LOWER', 1], 113: ['UPPER', 1], 114: ['PROPER', 1], 115: ['LEFT', -1], 116: ['RIGHT', -1], 117: ['EXACT', 2], 118: ['TRIM', 1],
  119: ['REPLACE', 4], 120: ['SUBSTITUTE', -1], 121: ['CODE', 1], 124: ['FIND', -1], 125: ['CELL', -1], 126: ['ISERR', 1], 127: ['ISTEXT', 1],
  128: ['ISNUMBER', 1], 129: ['ISBLANK', 1], 130: ['T', 1], 131: ['N', 1], 140: ['DATEVALUE', 1], 141: ['TIMEVALUE', 1], 142: ['SLN', 3], 143: ['SYD', 4],
  144: ['DDB', -1], 148: ['INDIRECT', -1], 162: ['CLEAN', 1], 163: ['MDETERM', 1], 164: ['MINVERSE', 1], 165: ['MMULT', 2], 167: ['IPMT', -1],
  168: ['PPMT', -1], 169: ['COUNTA', -1], 183: ['PRODUCT', -1], 184: ['FACT', 1], 189: ['DPRODUCT', 3], 190: ['ISNONTEXT', 1], 193: ['STDEVP', -1],
  194: ['VARP', -1], 195: ['DSTDEVP', 3], 196: ['DVARP', 3], 197: ['TRUNC', -1], 198: ['ISLOGICAL', 1], 199: ['DCOUNTA', 3], 204: ['DOLLAR', -1],
  205: ['FINDB', -1], 206: ['SEARCHB', -1], 207: ['REPLACEB', 4], 208: ['LEFTB', -1], 209: ['RIGHTB', -1], 210: ['MIDB', 3], 211: ['LENB', 1],
  212: ['ROUNDUP', 2], 213: ['ROUNDDOWN', 2], 214: ['ASC', 1], 215: ['DBCS', 1], 216: ['RANK', -1], 219: ['ADDRESS', -1], 220: ['DAYS360', -1],
  221: ['TODAY', 0], 222: ['VDB', -1], 227: ['MEDIAN', -1], 228: ['SUMPRODUCT', -1], 229: ['SINH', 1], 230: ['COSH', 1], 231: ['TANH', 1],
  232: ['ASINH', 1], 233: ['ACOSH', 1], 234: ['ATANH', 1], 235: ['DGET', 3], 244: ['INFO', 1], 247: ['DB', -1], 252: ['FREQUENCY', 2],
  261: ['ERROR.TYPE', 1], 269: ['AVEDEV', -1], 270: ['BETADIST', -1], 271: ['GAMMALN', 1], 272: ['BETAINV', -1], 273: ['BINOMDIST', 4],
  274: ['CHIDIST', 2], 275: ['CHIINV', 2], 276: ['COMBIN', 2], 277: ['CONFIDENCE', 3], 278: ['CRITBINOM', 3], 279: ['EVEN', 1], 280: ['EXPONDIST', 3],
  281: ['FDIST', 3], 282: ['FINV', 3], 283: ['FISHER', 1], 284: ['FISHERINV', 1], 285: ['FLOOR', 2], 286: ['GAMMADIST', 4], 287: ['GAMMAINV', 3],
  288: ['CEILING', 2], 289: ['HYPGEOMDIST', 4], 290: ['LOGNORMDIST', 3], 291: ['LOGINV', 3], 292: ['NEGBINOMDIST', 3], 293: ['NORMDIST', 4],
  294: ['NORMSDIST', 1], 295: ['NORMINV', 3], 296: ['NORMSINV', 1], 297: ['STANDARDIZE', 3], 298: ['ODD', 1], 299: ['PERMUT', 2], 300: ['POISSON', 3],
  301: ['TDIST', 3], 302: ['WEIBULL', 4], 303: ['SUMXMY2', 2], 304: ['SUMX2MY2', 2], 305: ['SUMX2PY2', 2], 306: ['CHITEST', 2], 307: ['CORREL', 2],
  308: ['COVAR', 2], 309: ['FORECAST', 3], 310: ['FTEST', 2], 311: ['INTERCEPT', 2], 312: ['PEARSON', 2], 313: ['RSQ', 2], 314: ['STEYX', 2],
  315: ['SLOPE', 2], 316: ['TTEST', 4], 317: ['PROB', -1], 318: ['DEVSQ', -1], 319: ['GEOMEAN', -1], 320: ['HARMEAN', -1], 321: ['SUMSQ', -1],
  322: ['KURT', -1], 323: ['SKEW', -1], 324: ['ZTEST', -1], 325: ['LARGE', 2], 326: ['SMALL', 2], 327: ['QUARTILE', 2], 328: ['PERCENTILE', 2],
  329: ['PERCENTRANK', -1], 330: ['MODE', -1], 331: ['TRIMMEAN', 2], 332: ['TINV', 2], 336: ['CONCATENATE', -1], 337: ['POWER', 2], 342: ['RADIANS', 1],
  343: ['DEGREES', 1], 344: ['SUBTOTAL', -1], 345: ['SUMIF', -1], 346: ['COUNTIF', 2], 347: ['COUNTBLANK', 1], 350: ['ISPMT', 4], 351: ['DATEDIF', 3],
  352: ['DATESTRING', 1], 353: ['NUMBERSTRING', 2], 354: ['ROMAN', -1], 358: ['GETPIVOTDATA', -1], 359: ['HYPERLINK', -1], 360: ['PHONETIC', 1],
  361: ['AVERAGEA', -1], 362: ['MAXA', -1], 363: ['MINA', -1], 364: ['STDEVPA', -1], 365: ['VARPA', -1], 366: ['STDEVA', -1], 367: ['VARA', -1],
  368: ['BAHTTEXT', 1], 369: ['THAIDAYOFWEEK', 1], 370: ['THAIDIGIT', 1], 371: ['THAIMONTHOFYEAR', 1], 372: ['THAINUMSOUND', 1], 373: ['THAINUMSTRING', 1],
  374: ['THAISTRINGLENGTH', 1], 375: ['ISTHAIDIGIT', 1], 376: ['ROUNDBAHTDOWN', 1], 377: ['ROUNDBAHTUP', 1], 378: ['THAIYEAR', 1], 379: ['RTD', -1],
  380: ['CUBEVALUE', -1], 381: ['CUBEMEMBER', -1], 382: ['CUBEMEMBERPROPERTY', -1], 383: ['CUBERANKEDMEMBER', -1], 384: ['HEX2BIN', -1],
  385: ['HEX2DEC', 1], 386: ['HEX2OCT', -1], 387: ['DEC2BIN', -1], 388: ['DEC2HEX', -1], 389: ['DEC2OCT', -1], 390: ['OCT2BIN', -1],
  391: ['OCT2HEX', -1], 392: ['OCT2DEC', 1], 393: ['BIN2DEC', 1], 394: ['BIN2OCT', -1], 395: ['BIN2HEX', -1], 396: ['IMSUB', 2], 397: ['IMDIV', 2],
  398: ['IMPOWER', 2], 399: ['IMABS', 1], 400: ['IMSQRT', 1], 401: ['IMLN', 1], 402: ['IMLOG2', 1], 403: ['IMLOG10', 1], 404: ['IMSIN', 1],
  405: ['IMCOS', 1], 406: ['IMEXP', 1], 407: ['IMARGUMENT', 1], 408: ['IMCONJUGATE', 1], 409: ['IMAGINARY', 1], 410: ['IMREAL', 1], 411: ['COMPLEX', -1],
  412: ['IMSUM', -1], 413: ['IMPRODUCT', -1], 414: ['SERIESSUM', 4], 415: ['FACTDOUBLE', 1], 416: ['SQRTPI', 1], 417: ['QUOTIENT', 2], 418: ['DELTA', -1],
  419: ['GESTEP', -1], 420: ['ISEVEN', 1], 421: ['ISODD', 1], 422: ['MROUND', 2], 423: ['ERF', -1], 424: ['ERFC', 1], 425: ['BESSELJ', 2],
  426: ['BESSELK', 2], 427: ['BESSELY', 2], 428: ['BESSELI', 2], 429: ['XIRR', -1], 430: ['XNPV', 3], 431: ['PRICEMAT', -1], 432: ['YIELDMAT', -1],
  433: ['INTRATE', -1], 434: ['RECEIVED', -1], 435: ['DISC', -1], 436: ['PRICEDISC', -1], 437: ['YIELDDISC', -1], 438: ['TBILLEQ', 3],
  439: ['TBILLPRICE', 3], 440: ['TBILLYIELD', 3], 441: ['PRICE', -1], 442: ['YIELD', -1], 443: ['DOLLARDE', 2], 444: ['DOLLARFR', 2], 445: ['NOMINAL', 2],
  446: ['EFFECT', 2], 447: ['CUMPRINC', 6], 448: ['CUMIPMT', 6], 449: ['EDATE', 2], 450: ['EOMONTH', 2], 451: ['YEARFRAC', -1], 452: ['COUPDAYBS', -1],
  453: ['COUPDAYS', -1], 454: ['COUPDAYSNC', -1], 455: ['COUPNCD', -1], 456: ['COUPNUM', -1], 457: ['COUPPCD', -1], 458: ['DURATION', -1],
  459: ['MDURATION', -1], 460: ['ODDLPRICE', -1], 461: ['ODDLYIELD', -1], 462: ['ODDFPRICE', -1], 463: ['ODDFYIELD', -1], 464: ['RANDBETWEEN', 2],
  465: ['WEEKNUM', -1], 466: ['AMORDEGRC', -1], 467: ['AMORLINC', -1], 468: ['CONVERT', 3], 469: ['ACCRINT', -1], 470: ['ACCRINTM', -1],
  471: ['WORKDAY', -1], 472: ['NETWORKDAYS', -1], 473: ['GCD', -1], 474: ['MULTINOMIAL', -1], 475: ['LCM', -1], 476: ['FVSCHEDULE', 2],
  477: ['CUBEKPIMEMBER', -1], 478: ['CUBESET', -1], 479: ['CUBESETCOUNT', 1], 480: ['IFERROR', 2], 481: ['COUNTIFS', -1], 482: ['SUMIFS', -1],
  483: ['AVERAGEIF', -1], 484: ['AVERAGEIFS', -1],
};

const ERRORS: Record<number, string> = { 0x00: '#NULL!', 0x07: '#DIV/0!', 0x0f: '#VALUE!', 0x17: '#REF!', 0x1d: '#NAME?', 0x24: '#NUM!', 0x2a: '#N/A', 0x2b: '#GETTING_DATA' };

export function errorText(code: number): string {
  return ERRORS[code] ?? '#N/A';
}

const BINARY: Record<number, string> = { 0x03: '+', 0x04: '-', 0x05: '*', 0x06: '/', 0x07: '^', 0x08: '&', 0x09: '<', 0x0a: '<=', 0x0b: '=', 0x0c: '>=', 0x0d: '>', 0x0e: '<>', 0x0f: ' ', 0x10: ',', 0x11: ':' };

export interface FormulaContext {
  /** 'biff8' (XLS) or 'biff12' (XLSB): field sizes differ. */
  format: 'biff8' | 'biff12';
  /** Cell holding the formula (relative references of shared formulas are offsets from it). */
  row: number;
  col: number;
  /** Sheet name of an XTI index (3-D references), or null for external workbooks. */
  sheetOf(ixti: number): { first: string; last: string } | null;
  /** Defined name text by 1-based index. */
  nameOf(index: number): string | null;
  /** External or add-in name (ptgNameX). */
  externName(ixti: number, index: number): string | null;
  /** Formula text of a shared or array formula anchored at a cell (ptgExp), or null. */
  shared?(row: number, col: number): string | null;
}

class Reader {
  pos = 0;
  constructor(readonly b: Uint8Array, readonly v: DataView, start = 0) {
    this.pos = start;
  }
  u8(): number {
    return this.b[this.pos++];
  }
  u16(): number {
    const x = this.v.getUint16(this.pos, true);
    this.pos += 2;
    return x;
  }
  i16(): number {
    const x = this.v.getInt16(this.pos, true);
    this.pos += 2;
    return x;
  }
  u32(): number {
    const x = this.v.getUint32(this.pos, true);
    this.pos += 4;
    return x;
  }
  i32(): number {
    const x = this.v.getInt32(this.pos, true);
    this.pos += 4;
    return x;
  }
  f64(): number {
    const x = this.v.getFloat64(this.pos, true);
    this.pos += 8;
    return x;
  }
  skip(n: number): void {
    this.pos += n;
  }
  /** BIFF8 string: count, flag byte, then 8- or 16-bit characters. */
  str8(cchSize: 1 | 2): string {
    const cch = cchSize === 1 ? this.u8() : this.u16();
    const flags = this.u8();
    let s = '';
    if (flags & 1) {
      for (let i = 0; i < cch; i++) s += String.fromCharCode(this.u16());
    } else {
      for (let i = 0; i < cch; i++) s += String.fromCharCode(this.u8());
    }
    return s;
  }
  /** BIFF12 string: 16-bit count of UTF-16 characters. */
  str12(): string {
    const cch = this.u16();
    let s = '';
    for (let i = 0; i < cch; i++) s += String.fromCharCode(this.u16());
    return s;
  }
}

const quoteStr = (s: string) => `"${s.replace(/"/g, '""')}"`;

function arrayLiteral(r: Reader, format: 'biff8' | 'biff12'): string {
  let cols: number;
  let rows: number;
  if (format === 'biff8') {
    cols = r.u8() + 1;
    rows = r.u16() + 1;
  } else {
    rows = r.u32();
    cols = r.u32();
  }
  const out: string[] = [];
  for (let i = 0; i < rows; i++) {
    const row: string[] = [];
    for (let j = 0; j < cols; j++) {
      const t = r.u8();
      switch (t) {
        case 0x00:
          r.skip(8);
          row.push('');
          break;
        case 0x01: {
          const n = r.f64();
          row.push(String(n));
          break;
        }
        case 0x02:
          row.push(quoteStr(format === 'biff8' ? r.str8(2) : r.str12()));
          break;
        case 0x04: {
          const b = r.u8();
          r.skip(format === 'biff8' ? 7 : 0);
          row.push(b ? 'TRUE' : 'FALSE');
          break;
        }
        case 0x10: {
          const e = r.u8();
          r.skip(format === 'biff8' ? 7 : 0);
          row.push(errorText(e));
          break;
        }
        default:
          throw new Error('array');
      }
    }
    out.push(row.join(','));
  }
  return `{${out.join(';')}}`;
}

function cellRef(row: number, col: number, rowRel: boolean, colRel: boolean): string {
  return `${colRel ? '' : '$'}${colName(col)}${rowRel ? '' : '$'}${row + 1}`;
}

/**
 * Formula text (without '=') of a token stream, or null when it uses tokens that cannot be
 * represented (the cell then keeps its last value).
 */
export function decodeFormula(rgce: Uint8Array, extra: Uint8Array, ctx: FormulaContext): string | null {
  const view = new DataView(rgce.buffer, rgce.byteOffset, rgce.byteLength);
  const r = new Reader(rgce, view);
  const ex = new Reader(extra, new DataView(extra.buffer, extra.byteOffset, extra.byteLength));
  const b12 = ctx.format === 'biff12';
  const maxRow = b12 ? 1048576 : 65536;
  const maxCol = b12 ? 16384 : 256;
  const stack: string[] = [];
  const pop = () => {
    const v = stack.pop();
    if (v === undefined) throw new Error('stack');
    return v;
  };
  /** Row and column of a reference token (absolute, or relative to the formula cell for N tokens). */
  const readRowCol = (relativeToCell: boolean): string => {
    const rowRaw = b12 ? r.u32() : r.u16();
    const colRaw = r.u16();
    const rowRel = (colRaw & 0x8000) !== 0;
    const colRel = (colRaw & 0x4000) !== 0;
    let row = rowRaw;
    let col = colRaw & 0x3fff;
    if (relativeToCell) {
      if (rowRel) row = (ctx.row + (b12 ? rowRaw | 0 : (rowRaw << 16) >> 16) + maxRow) % maxRow;
      if (colRel) col = (ctx.col + (b12 ? (col << 18) >> 18 : (col << 24) >> 24) + maxCol) % maxCol;
    }
    return cellRef(row, col, rowRel, colRel);
  };
  const readArea = (relativeToCell: boolean): string => {
    const r0 = b12 ? r.u32() : r.u16();
    const r1 = b12 ? r.u32() : r.u16();
    const c0 = r.u16();
    const c1 = r.u16();
    const fix = (rowRaw: number, colRaw: number) => {
      const rowRel = (colRaw & 0x8000) !== 0;
      const colRel = (colRaw & 0x4000) !== 0;
      let row = rowRaw;
      let col = colRaw & 0x3fff;
      if (relativeToCell) {
        if (rowRel) row = (ctx.row + (b12 ? rowRaw | 0 : (rowRaw << 16) >> 16) + maxRow) % maxRow;
        if (colRel) col = (ctx.col + (b12 ? (col << 18) >> 18 : (col << 24) >> 24) + maxCol) % maxCol;
      }
      return { row, col, rowRel, colRel };
    };
    const a = fix(r0, c0);
    const b = fix(r1, c1);
    // Whole columns and rows.
    if (a.row === 0 && b.row === maxRow - 1 && !a.rowRel && !b.rowRel) return `${a.colRel ? '' : '$'}${colName(a.col)}:${b.colRel ? '' : '$'}${colName(b.col)}`;
    if (a.col === 0 && b.col === maxCol - 1 && !a.colRel && !b.colRel) return `${a.rowRel ? '' : '$'}${a.row + 1}:${b.rowRel ? '' : '$'}${b.row + 1}`;
    return `${cellRef(a.row, a.col, a.rowRel, a.colRel)}:${cellRef(b.row, b.col, b.rowRel, b.colRel)}`;
  };
  const sheetPrefix = (ixti: number): string => {
    const s = ctx.sheetOf(ixti);
    if (!s) throw new Error('external');
    return s.first === s.last ? `${quoteSheetName(s.first)}!` : `${quoteSheetName(`${s.first}:${s.last}`)}!`;
  };
  while (r.pos < rgce.length) {
    const ptg = r.u8();
    const base = ptg >= 0x20 ? (ptg & 0x1f) | 0x20 : ptg;
    switch (base) {
      case 0x01: {
        // Shared or array formula.
        const row = b12 ? r.u32() : r.u16();
        const col = b12 ? ctx.col : r.u16();
        const text = ctx.shared?.(row, col);
        if (text === null || text === undefined) return null;
        return text;
      }
      case 0x03: case 0x04: case 0x05: case 0x06: case 0x07: case 0x08: case 0x09: case 0x0a: case 0x0b: case 0x0c: case 0x0d: case 0x0e: case 0x0f: case 0x10: case 0x11: {
        const b = pop();
        const a = pop();
        stack.push(`${a}${BINARY[base]}${b}`);
        break;
      }
      case 0x12:
        stack.push(`+${pop()}`);
        break;
      case 0x13:
        stack.push(`-${pop()}`);
        break;
      case 0x14:
        stack.push(`${pop()}%`);
        break;
      case 0x15:
        stack.push(`(${pop()})`);
        break;
      case 0x16:
        stack.push('');
        break;
      case 0x17:
        stack.push(quoteStr(b12 ? r.str12() : r.str8(1)));
        break;
      case 0x19: {
        const flags = r.u8();
        if (flags & 0x04) {
          // CHOOSE jump table.
          const cases = r.u16();
          r.skip((cases + 1) * 2);
        } else {
          r.skip(2);
          if (flags & 0x10) stack.push(`SUM(${pop()})`);
        }
        break;
      }
      case 0x1c:
        stack.push(errorText(r.u8()));
        break;
      case 0x1d:
        stack.push(r.u8() ? 'TRUE' : 'FALSE');
        break;
      case 0x1e:
        stack.push(String(r.u16()));
        break;
      case 0x1f:
        stack.push(String(r.f64()));
        break;
      case 0x20:
        // Array constant (its values follow the tokens).
        r.skip(b12 ? 14 : 7);
        stack.push(arrayLiteral(ex, ctx.format));
        break;
      case 0x21:
      case 0x22: {
        let argc: number;
        let index: number;
        if (base === 0x22) {
          argc = r.u8() & 0x7f;
          index = r.u16() & 0x7fff;
        } else {
          index = r.u16();
          argc = FUNCS[index]?.[1] ?? -1;
          if (argc < 0) return null;
        }
        const args: string[] = [];
        for (let i = 0; i < argc; i++) args.unshift(pop());
        if (index === 255) {
          // User-defined or add-in function: its name is the first "argument".
          const name = args.shift() ?? '';
          stack.push(`${name}(${args.join(',')})`);
          break;
        }
        const name = FUNCS[index]?.[0];
        if (!name) return null;
        stack.push(`${name}(${args.join(',')})`);
        break;
      }
      case 0x23: {
        const name = ctx.nameOf(r.u32());
        if (!name) return null;
        stack.push(name);
        break;
      }
      case 0x24:
        stack.push(readRowCol(false));
        break;
      case 0x25:
        stack.push(readArea(false));
        break;
      case 0x26:
        // Memory area: the reference expression follows; its extra data lists the areas.
        r.skip(6);
        {
          const count = b12 ? ex.u32() : ex.u16();
          ex.skip(count * (b12 ? 12 : 8));
        }
        break;
      case 0x27:
      case 0x28:
        r.skip(6);
        break;
      case 0x29:
        r.skip(2);
        break;
      case 0x2a:
        r.skip(b12 ? 6 : 4);
        stack.push('#REF!');
        break;
      case 0x2b:
        r.skip(b12 ? 12 : 8);
        stack.push('#REF!');
        break;
      case 0x2c:
        stack.push(readRowCol(true));
        break;
      case 0x2d:
        stack.push(readArea(true));
        break;
      case 0x39: {
        const ixti = r.u16();
        const index = r.u32();
        const name = ctx.externName(ixti, index);
        if (!name) return null;
        stack.push(name);
        break;
      }
      case 0x3a: {
        const ixti = r.u16();
        stack.push(`${sheetPrefix(ixti)}${readRowCol(false)}`);
        break;
      }
      case 0x3b: {
        const ixti = r.u16();
        stack.push(`${sheetPrefix(ixti)}${readArea(false)}`);
        break;
      }
      case 0x3c: {
        const ixti = r.u16();
        r.skip(b12 ? 6 : 4);
        stack.push(`${sheetPrefix(ixti)}#REF!`);
        break;
      }
      case 0x3d: {
        const ixti = r.u16();
        r.skip(b12 ? 12 : 8);
        stack.push(`${sheetPrefix(ixti)}#REF!`);
        break;
      }
      default:
        return null;
    }
  }
  if (stack.length !== 1) return null;
  return stack[0];
}

export { Reader };
