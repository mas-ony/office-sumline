/**
 * Canonical namespace prefixes. Parsed SpreadsheetML parts are normalised so the main namespace
 * has no prefix (`sheetData`, `c`, `v`) and every other namespace uses the prefix below, whatever
 * prefixes the producing application chose (including ISO 29500 "strict" namespaces).
 */
export const NS: Record<string, string> = {
  '': 'http://schemas.openxmlformats.org/spreadsheetml/2006/main',
  r: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
  mc: 'http://schemas.openxmlformats.org/markup-compatibility/2006',
  x14ac: 'http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac',
  x14: 'http://schemas.microsoft.com/office/spreadsheetml/2009/9/main',
  x15: 'http://schemas.microsoft.com/office/spreadsheetml/2010/11/main',
  x15ac: 'http://schemas.microsoft.com/office/spreadsheetml/2010/11/ac',
  x16r2: 'http://schemas.microsoft.com/office/spreadsheetml/2015/02/main',
  xr: 'http://schemas.microsoft.com/office/spreadsheetml/2014/revision',
  xr2: 'http://schemas.microsoft.com/office/spreadsheetml/2015/revision2',
  xr3: 'http://schemas.microsoft.com/office/spreadsheetml/2016/revision3',
  xr6: 'http://schemas.microsoft.com/office/spreadsheetml/2016/revision6',
  xr10: 'http://schemas.microsoft.com/office/spreadsheetml/2016/revision10',
  xcalcf: 'http://schemas.microsoft.com/office/spreadsheetml/2018/calcfeatures',
  xda: 'http://schemas.microsoft.com/office/spreadsheetml/2017/dynamicarray',
  xlrd: 'http://schemas.microsoft.com/office/spreadsheetml/2017/richdata',
  xm: 'http://schemas.microsoft.com/office/excel/2006/main',
  xdr: 'http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing',
  a: 'http://schemas.openxmlformats.org/drawingml/2006/main',
  a14: 'http://schemas.microsoft.com/office/drawing/2010/main',
  a16: 'http://schemas.microsoft.com/office/drawing/2014/main',
  c: 'http://schemas.openxmlformats.org/drawingml/2006/chart',
  pic: 'http://schemas.openxmlformats.org/drawingml/2006/picture',
  tc: 'http://schemas.microsoft.com/office/spreadsheetml/2018/threadedcomments',
  v: 'urn:schemas-microsoft-com:vml',
  o: 'urn:schemas-microsoft-com:office:office',
  x: 'urn:schemas-microsoft-com:office:excel',
  xe: 'urn:schemas-xlsx-editor:extensions',
  xml: 'http://www.w3.org/XML/1998/namespace',
};

/** ISO/IEC 29500 strict namespaces mapped onto their transitional canonical prefix. */
const STRICT_ALIASES: Record<string, string> = {
  'http://purl.oclc.org/ooxml/spreadsheetml/main': '',
  'http://purl.oclc.org/ooxml/officeDocument/relationships': 'r',
  'http://purl.oclc.org/ooxml/drawingml/main': 'a',
  'http://purl.oclc.org/ooxml/drawingml/spreadsheetDrawing': 'xdr',
  'http://purl.oclc.org/ooxml/drawingml/chart': 'c',
  'http://purl.oclc.org/ooxml/drawingml/picture': 'pic',
};

const URI_TO_PREFIX = new Map<string, string>();
for (const [prefix, uri] of Object.entries(NS)) URI_TO_PREFIX.set(uri, prefix);
for (const [uri, prefix] of Object.entries(STRICT_ALIASES)) URI_TO_PREFIX.set(uri, prefix);

export function canonicalPrefixFor(uri: string): string | undefined {
  return URI_TO_PREFIX.get(uri);
}

export const MAIN_NS = NS[''];
export const REL_NS = NS.r;
