// Generates XLSX test fixtures in tests/fixtures. Run: node scripts/make-fixtures.mjs
import { writeFileSync, mkdirSync } from 'node:fs';
import { deflateSync } from 'node:zlib';
import { zipSync, strToU8 } from 'fflate';

const OUT = new URL('../tests/fixtures/', import.meta.url);
mkdirSync(OUT, { recursive: true });

/**
 * Zip a package with the same time on every entry, so that a rerun writes the same bytes. fflate
 * stores local time, so local midnight on 1 January 2024 gives the same bytes in every time zone.
 */
const zip = (parts) => zipSync(parts, { mtime: new Date(2024, 0, 1) });

const DECL = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n';
const MAIN = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main';
const R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const PR = 'http://schemas.openxmlformats.org/package/2006/relationships';
const REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const esc = (s) => String(s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');

// --- tiny PNG encoder ---------------------------------------------------------------------------
function crc32(buf) {
  let c, crc = 0xffffffff;
  for (let n = 0; n < buf.length; n++) {
    c = (crc ^ buf[n]) & 0xff;
    for (let k = 0; k < 8; k++) c = c & 1 ? 0xedb88320 ^ (c >>> 1) : c >>> 1;
    crc = (crc >>> 8) ^ c;
  }
  return (crc ^ 0xffffffff) >>> 0;
}
function chunk(type, data) {
  const len = Buffer.alloc(4); len.writeUInt32BE(data.length);
  const td = Buffer.concat([Buffer.from(type), data]);
  const crc = Buffer.alloc(4); crc.writeUInt32BE(crc32(td));
  return Buffer.concat([len, td, crc]);
}
function png(w, h, pixel) {
  const raw = Buffer.alloc((w * 3 + 1) * h);
  for (let y = 0; y < h; y++) {
    raw[y * (w * 3 + 1)] = 0;
    for (let x = 0; x < w; x++) {
      const [r, g, b] = pixel(x, y);
      const o = y * (w * 3 + 1) + 1 + x * 3;
      raw[o] = r; raw[o + 1] = g; raw[o + 2] = b;
    }
  }
  const ihdr = Buffer.alloc(13);
  ihdr.writeUInt32BE(w, 0); ihdr.writeUInt32BE(h, 4); ihdr[8] = 8; ihdr[9] = 2;
  return Buffer.concat([Buffer.from([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a]), chunk('IHDR', ihdr), chunk('IDAT', deflateSync(raw)), chunk('IEND', Buffer.alloc(0))]);
}

const colName = (n) => { let s = ''; n += 1; while (n) { const r = (n - 1) % 26; s = String.fromCharCode(65 + r) + s; n = Math.floor((n - 1) / 26); } return s; };

/** A small worksheet builder. */
class SheetBuilder {
  constructor(shared) { this.rows = new Map(); this.shared = shared; }
  put(r, c, xml) { if (!this.rows.has(r)) this.rows.set(r, []); this.rows.get(r).push([c, xml]); }
  s(r, c, text, st = 0) { this.put(r, c, `<c r="${colName(c)}${r}" t="s"${st ? ` s="${st}"` : ''}><v>${this.shared.add(text)}</v></c>`); }
  n(r, c, v, st = 0) { this.put(r, c, `<c r="${colName(c)}${r}"${st ? ` s="${st}"` : ''}><v>${v}</v></c>`); }
  f(r, c, formula, cached, st = 0, t = null, extra = '') { this.put(r, c, `<c r="${colName(c)}${r}"${st ? ` s="${st}"` : ''}${t ? ` t="${t}"` : ''}${extra}><f>${esc(formula)}</f><v>${cached}</v></c>`); }
  raw(r, c, xml) { this.put(r, c, xml); }
  xml({ before = '', after = '', cols = '', views = '', fmt = '<sheetFormatPr defaultRowHeight="15"/>', heights = {} } = {}) {
    const body = [...this.rows.keys()].sort((a, b) => a - b).map((r) => {
      const cells = this.rows.get(r).sort((a, b) => a[0] - b[0]).map((x) => x[1]).join('');
      const ht = heights[r] ? ` ht="${heights[r]}" customHeight="1"` : '';
      return `<row r="${r}"${ht}>${cells}</row>`;
    }).join('');
    return `${DECL}<worksheet xmlns="${MAIN}" xmlns:r="${R}" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:x14ac="http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac" mc:Ignorable="x14ac">${views}${fmt}${cols}<sheetData>${body}</sheetData>${before}${after}</worksheet>`;
  }
}

class Shared {
  constructor() { this.list = []; this.count = 0; }
  add(s) { this.count++; let i = this.list.indexOf(s); if (i < 0) { i = this.list.length; this.list.push(s); } return i; }
  xml() { return `${DECL}<sst xmlns="${MAIN}" count="${this.count}" uniqueCount="${this.list.length}">${this.list.map((s) => `<si><t xml:space="preserve">${esc(s)}</t></si>`).join('')}</sst>`; }
}

const STYLES = `${DECL}<styleSheet xmlns="${MAIN}">
<numFmts count="4"><numFmt numFmtId="164" formatCode="&quot;$&quot;#,##0.00"/><numFmt numFmtId="165" formatCode="0.0%"/><numFmt numFmtId="166" formatCode="yyyy-mm-dd"/><numFmt numFmtId="167" formatCode="#,##0.00;[Red](#,##0.00)"/></numFmts>
<fonts count="6">
<font><sz val="11"/><color theme="1"/><name val="Calibri"/><family val="2"/><scheme val="minor"/></font>
<font><b/><sz val="11"/><color theme="1"/><name val="Calibri"/><family val="2"/><scheme val="minor"/></font>
<font><b/><sz val="16"/><color rgb="FFFFFFFF"/><name val="Calibri"/><family val="2"/><scheme val="minor"/></font>
<font><i/><sz val="11"/><color rgb="FFC00000"/><name val="Arial"/><family val="2"/></font>
<font><u/><sz val="11"/><color rgb="FF0563C1"/><name val="Calibri"/><family val="2"/><scheme val="minor"/></font>
<font><strike/><sz val="11"/><color theme="1"/><name val="Calibri"/><family val="2"/><scheme val="minor"/></font>
</fonts>
<fills count="5"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill><fill><patternFill patternType="solid"><fgColor rgb="FF1F4E79"/><bgColor indexed="64"/></patternFill></fill><fill><patternFill patternType="solid"><fgColor rgb="FFDDEBF7"/><bgColor indexed="64"/></patternFill></fill><fill><patternFill patternType="solid"><fgColor rgb="FFFFF2CC"/><bgColor indexed="64"/></patternFill></fill></fills>
<borders count="4"><border><left/><right/><top/><bottom/><diagonal/></border><border><left style="thin"><color auto="1"/></left><right style="thin"><color auto="1"/></right><top style="thin"><color auto="1"/></top><bottom style="thin"><color auto="1"/></bottom><diagonal/></border><border><left/><right/><top style="thin"><color rgb="FF1F4E79"/></top><bottom style="double"><color rgb="FF1F4E79"/></bottom><diagonal/></border><border><left/><right/><top/><bottom style="medium"><color rgb="FF000000"/></bottom><diagonal/></border></borders>
<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>
<cellXfs count="14">
<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>
<xf numFmtId="0" fontId="2" fillId="2" borderId="0" xfId="0" applyFont="1" applyFill="1" applyAlignment="1"><alignment horizontal="center" vertical="center"/></xf>
<xf numFmtId="0" fontId="1" fillId="3" borderId="3" xfId="0" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1"><alignment horizontal="center" vertical="center" wrapText="1"/></xf>
<xf numFmtId="0" fontId="0" fillId="0" borderId="1" xfId="0" applyBorder="1"/>
<xf numFmtId="164" fontId="0" fillId="0" borderId="1" xfId="0" applyNumberFormat="1" applyBorder="1"/>
<xf numFmtId="165" fontId="0" fillId="0" borderId="1" xfId="0" applyNumberFormat="1" applyBorder="1"/>
<xf numFmtId="166" fontId="0" fillId="0" borderId="1" xfId="0" applyNumberFormat="1" applyBorder="1" applyAlignment="1"><alignment horizontal="right"/></xf>
<xf numFmtId="164" fontId="1" fillId="4" borderId="2" xfId="0" applyNumberFormat="1" applyFont="1" applyFill="1" applyBorder="1"/>
<xf numFmtId="0" fontId="3" fillId="0" borderId="0" xfId="0" applyFont="1" applyAlignment="1"><alignment vertical="top" wrapText="1"/></xf>
<xf numFmtId="0" fontId="4" fillId="0" borderId="0" xfId="0" applyFont="1"/>
<xf numFmtId="167" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/>
<xf numFmtId="0" fontId="5" fillId="0" borderId="0" xfId="0" applyFont="1" applyAlignment="1"><alignment horizontal="right"/></xf>
<xf numFmtId="3" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/>
<xf numFmtId="0" fontId="1" fillId="0" borderId="0" xfId="0" applyFont="1"/>
</cellXfs>
<cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>
<dxfs count="1"><dxf><font><color rgb="FF9C0006"/></font><fill><patternFill><bgColor rgb="FFFFC7CE"/></patternFill></fill></dxf></dxfs>
<tableStyles count="0" defaultTableStyle="TableStyleMedium2" defaultPivotStyle="PivotStyleLight16"/>
</styleSheet>`;

function contentTypes(overrides, defaults = []) {
  const d = [['rels', 'application/vnd.openxmlformats-package.relationships+xml'], ['xml', 'application/xml'], ...defaults];
  return `${DECL}<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">${d.map(([e, t]) => `<Default Extension="${e}" ContentType="${t}"/>`).join('')}${overrides.map(([p, t]) => `<Override PartName="${p}" ContentType="${t}"/>`).join('')}</Types>`;
}
const rels = (list) => `${DECL}<Relationships xmlns="${PR}">${list.map(([id, type, target, ext]) => `<Relationship Id="${id}" Type="${type}" Target="${esc(target)}"${ext ? ' TargetMode="External"' : ''}/>`).join('')}</Relationships>`;
const CT = {
  wb: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml',
  ws: 'application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml',
  styles: 'application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml',
  sst: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml',
  comments: 'application/vnd.openxmlformats-officedocument.spreadsheetml.comments+xml',
  drawing: 'application/vnd.openxmlformats-officedocument.drawing+xml',
  chart: 'application/vnd.openxmlformats-officedocument.drawingml.chart+xml',
  core: 'application/vnd.openxmlformats-package.core-properties+xml',
  app: 'application/vnd.openxmlformats-officedocument.extended-properties+xml',
  threaded: 'application/vnd.ms-excel.threadedcomments+xml',
  person: 'application/vnd.ms-excel.person+xml',
  metadata: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheetMetadata+xml',
};
const CORE = `${DECL}<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"><dc:title>Fixture</dc:title><dc:creator>fixture</dc:creator><dcterms:modified xsi:type="dcterms:W3CDTF">2024-01-01T00:00:00Z</dcterms:modified></cp:coreProperties>`;
const APP = `${DECL}<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties"><Application>Microsoft Excel</Application></Properties>`;
const ROOT_RELS = rels([['rId1', `${REL}/officeDocument`, 'xl/workbook.xml'], ['rId2', 'http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties', 'docProps/core.xml'], ['rId3', `${REL}/extended-properties`, 'docProps/app.xml']]);

function save(name, files) {
  const zipped = zip(Object.fromEntries(Object.entries(files).map(([k, v]) => [k, typeof v === 'string' ? strToU8(v) : v])));
  writeFileSync(new URL(name, OUT), zipped);
  console.log('wrote', name, zipped.length, 'bytes');
}

// ================================================================================================
// kitchen-sink.xlsx: styles, formulas, merges, frozen pane, CF, DV, hyperlink, note, image, hidden sheet
{
  const shared = new Shared();
  const s1 = new SheetBuilder(shared);
  s1.s(1, 0, 'Quarterly Sales Report', 1);
  for (let c = 1; c < 6; c++) s1.raw(1, c, `<c r="${colName(c)}1" s="1"/>`);
  ['Region', 'Units', 'Unit price', 'Revenue', 'Growth', 'Since'].forEach((h, c) => s1.s(2, c, h, 2));
  const data = [['North', 120, 9.5, 0.125, 45000], ['South', 85, 12.25, -0.034, 45031], ['East', 240, 7.75, 0.21, 45062], ['West', 64, 15.0, 0.0, 45092], ['Central', 150, 10.0, 0.056, 45123]];
  data.forEach(([reg, units, price, growth, since], i) => {
    const r = 3 + i;
    s1.s(r, 0, reg, 3); s1.n(r, 1, units, 3); s1.n(r, 2, price, 4); s1.f(r, 3, `B${r}*C${r}`, units * price, 4); s1.n(r, 4, growth, 5); s1.n(r, 5, since, 6);
  });
  s1.s(8, 0, 'Total', 13);
  s1.f(8, 1, 'SUM(B3:B7)', 659, 12);
  s1.f(8, 3, 'SUM(D3:D7)', 6501.25, 7);
  s1.f(8, 4, 'AVERAGE(E3:E7)', 0.0714, 5);
  s1.s(10, 0, 'Biggest region'); s1.f(10, 1, 'INDEX(A3:A7,MATCH(MAX(D3:D7),D3:D7,0))', 'East', 0, 'str');
  s1.s(11, 0, 'Units status'); s1.f(11, 1, 'IF(B8>500,"On target","Below target")', 'On target', 0, 'str');
  s1.s(12, 0, 'From Data sheet'); s1.f(12, 1, 'SUM(Data!B2:B31)', 0, 12);
  s1.s(13, 0, 'Link'); s1.s(13, 1, 'Example site', 9);
  s1.s(14, 0, 'Notes'); s1.s(14, 1, 'This long note wraps inside its cell because wrap text is turned on for it.', 8);
  s1.s(15, 0, 'Bool'); s1.raw(15, 1, '<c r="B15" t="b"><v>1</v></c>');
  s1.s(16, 0, 'Financial'); s1.n(16, 1, -1234.5, 10);
  s1.s(17, 0, 'Struck'); s1.s(17, 1, 'old value', 11);
  s1.s(18, 0, 'Error'); s1.f(18, 1, '1/0', '#DIV/0!', 0, 'e');
  s1.s(19, 0, 'Named'); s1.f(19, 1, 'SUM(Revenue)', 6501.25, 4);
  s1.s(20, 0, 'Choose region');
  const sheet1 = s1.xml({
    views: '<sheetViews><sheetView tabSelected="1" workbookViewId="0"><pane ySplit="2" topLeftCell="A3" activePane="bottomLeft" state="frozen"/><selection pane="bottomLeft" activeCell="B3" sqref="B3"/></sheetView></sheetViews>',
    fmt: '<sheetFormatPr defaultRowHeight="15" x14ac:dyDescent="0.25"/>',
    cols: '<cols><col min="1" max="1" width="18" customWidth="1"/><col min="2" max="2" width="14" customWidth="1"/><col min="3" max="5" width="13" customWidth="1"/><col min="6" max="6" width="12" customWidth="1"/></cols>',
    heights: { 1: 30, 14: 45 },
    before: '<mergeCells count="1"><mergeCell ref="A1:F1"/></mergeCells><conditionalFormatting sqref="E3:E7"><cfRule type="cellIs" dxfId="0" priority="1" operator="lessThan"><formula>0</formula></cfRule></conditionalFormatting><dataValidations count="1"><dataValidation type="list" allowBlank="1" showInputMessage="1" showErrorMessage="1" sqref="B20"><formula1>"North,South,East,West,Central"</formula1></dataValidation></dataValidations><hyperlinks><hyperlink ref="B13" r:id="rId1" tooltip="Open the site"/></hyperlinks>',
    after: '<pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/><pageSetup orientation="landscape"/><drawing r:id="rId2"/><legacyDrawing r:id="rId3"/>',
  });
  const s2 = new SheetBuilder(shared);
  s2.s(1, 0, 'Day', 13); s2.s(1, 1, 'Visitors', 13); s2.s(1, 2, 'Share', 13);
  for (let i = 0; i < 30; i++) {
    s2.n(2 + i, 0, 45200 + i, 6);
    s2.n(2 + i, 1, 100 + (i * 37) % 90);
    s2.f(2 + i, 2, `B${2 + i}/SUM($B$2:$B$31)`, 0, 5);
  }
  const sheet2 = s2.xml({ cols: '<cols><col min="1" max="1" width="12" customWidth="1"/></cols>' });
  const s3 = new SheetBuilder(shared);
  s3.s(1, 0, 'Lookup'); s3.n(1, 1, 42);
  const sheet3 = s3.xml();
  const comments = `${DECL}<comments xmlns="${MAIN}"><authors><author>Reviewer</author></authors><commentList><comment ref="D8" authorId="0"><text><r><rPr><b/><sz val="9"/><color indexed="81"/><rFont val="Tahoma"/></rPr><t>Reviewer:</t></r><r><rPr><sz val="9"/><color indexed="81"/><rFont val="Tahoma"/></rPr><t xml:space="preserve">\nCheck the West region figure.</t></r></text></comment></commentList></comments>`;
  const vml = '<xml xmlns:v="urn:schemas-microsoft-com:vml" xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:x="urn:schemas-microsoft-com:office:excel"><o:shapelayout v:ext="edit"><o:idmap v:ext="edit" data="1"/></o:shapelayout><v:shapetype id="_x0000_t202" coordsize="21600,21600" o:spt="202" path="m,l,21600r21600,l21600,xe"><v:stroke joinstyle="miter"/><v:path gradientshapeok="t" o:connecttype="rect"/></v:shapetype><v:shape id="_x0000_s1025" type="#_x0000_t202" style="position:absolute;margin-left:300pt;margin-top:120pt;width:108pt;height:59.25pt;z-index:1;visibility:hidden" fillcolor="#ffffe1" o:insetmode="auto"><v:fill color2="#ffffe1"/><v:shadow on="t" color="black" obscured="t"/><v:path o:connecttype="none"/><v:textbox style="mso-direction-alt:auto"><div style="text-align:left"></div></v:textbox><x:ClientData ObjectType="Note"><x:MoveWithCells/><x:SizeWithCells/><x:Anchor>4, 15, 6, 10, 6, 15, 10, 4</x:Anchor><x:AutoFill>False</x:AutoFill><x:Row>7</x:Row><x:Column>3</x:Column></x:ClientData></v:shape></xml>';
  const drawing = `${DECL}<xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><xdr:twoCellAnchor editAs="oneCell"><xdr:from><xdr:col>7</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>2</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from><xdr:to><xdr:col>9</xdr:col><xdr:colOff>304800</xdr:colOff><xdr:row>9</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:to><xdr:pic><xdr:nvPicPr><xdr:cNvPr id="2" name="Picture 1" descr="Gradient"/><xdr:cNvPicPr><a:picLocks noChangeAspect="1"/></xdr:cNvPicPr></xdr:nvPicPr><xdr:blipFill><a:blip xmlns:r="${R}" r:embed="rId1"/><a:stretch><a:fillRect/></a:stretch></xdr:blipFill><xdr:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="1524000" cy="1333500"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></xdr:spPr></xdr:pic><xdr:clientData/></xdr:twoCellAnchor>
<xdr:twoCellAnchor><xdr:from><xdr:col>7</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>11</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from><xdr:to><xdr:col>12</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>24</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:to><xdr:graphicFrame macro=""><xdr:nvGraphicFramePr><xdr:cNvPr id="3" name="Chart 1"/><xdr:cNvGraphicFramePr/></xdr:nvGraphicFramePr><xdr:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/></xdr:xfrm><a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart"><c:chart xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:r="${R}" r:id="rId2"/></a:graphicData></a:graphic></xdr:graphicFrame><xdr:clientData/></xdr:twoCellAnchor></xdr:wsDr>`;
  const chart = `${DECL}<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="${R}"><c:chart><c:title><c:tx><c:rich><a:bodyPr/><a:p><a:r><a:t>Revenue by region</a:t></a:r></a:p></c:rich></c:tx><c:overlay val="0"/></c:title><c:plotArea><c:layout/><c:barChart><c:barDir val="col"/><c:grouping val="clustered"/><c:varyColors val="0"/><c:ser><c:idx val="0"/><c:order val="0"/><c:tx><c:strRef><c:f>Sales!$D$2</c:f></c:strRef></c:tx><c:cat><c:strRef><c:f>Sales!$A$3:$A$7</c:f></c:strRef></c:cat><c:val><c:numRef><c:f>Sales!$D$3:$D$7</c:f></c:numRef></c:val></c:ser><c:axId val="1"/><c:axId val="2"/></c:barChart><c:catAx><c:axId val="1"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:delete val="0"/><c:axPos val="b"/><c:crossAx val="2"/></c:catAx><c:valAx><c:axId val="2"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:delete val="0"/><c:axPos val="l"/><c:crossAx val="1"/></c:valAx></c:plotArea><c:plotVisOnly val="1"/></c:chart></c:chartSpace>`;
  const image = png(96, 84, (x, y) => [40 + x * 2, 90 + y, 200 - x]);
  save('kitchen-sink.xlsx', {
    '[Content_Types].xml': contentTypes([
      ['/xl/workbook.xml', CT.wb], ['/xl/worksheets/sheet1.xml', CT.ws], ['/xl/worksheets/sheet2.xml', CT.ws], ['/xl/worksheets/sheet3.xml', CT.ws],
      ['/xl/styles.xml', CT.styles], ['/xl/sharedStrings.xml', CT.sst], ['/xl/comments1.xml', CT.comments], ['/xl/drawings/drawing1.xml', CT.drawing],
      ['/xl/charts/chart1.xml', CT.chart], ['/docProps/core.xml', CT.core], ['/docProps/app.xml', CT.app],
    ], [['png', 'image/png'], ['vml', 'application/vnd.openxmlformats-officedocument.vmlDrawing']]),
    '_rels/.rels': ROOT_RELS,
    'docProps/core.xml': CORE,
    'docProps/app.xml': APP,
    'xl/workbook.xml': `${DECL}<workbook xmlns="${MAIN}" xmlns:r="${R}"><workbookPr defaultThemeVersion="166925"/><bookViews><workbookView activeTab="0"/></bookViews><sheets><sheet name="Sales" sheetId="1" r:id="rId1"/><sheet name="Data" sheetId="2" r:id="rId2"/><sheet name="Hidden" sheetId="3" state="hidden" r:id="rId3"/></sheets><definedNames><definedName name="Revenue">Sales!$D$3:$D$7</definedName></definedNames><calcPr calcId="191029"/></workbook>`,
    'xl/_rels/workbook.xml.rels': rels([['rId1', `${REL}/worksheet`, 'worksheets/sheet1.xml'], ['rId2', `${REL}/worksheet`, 'worksheets/sheet2.xml'], ['rId3', `${REL}/worksheet`, 'worksheets/sheet3.xml'], ['rId4', `${REL}/styles`, 'styles.xml'], ['rId5', `${REL}/sharedStrings`, 'sharedStrings.xml']]),
    'xl/styles.xml': STYLES,
    'xl/sharedStrings.xml': shared.xml(),
    'xl/worksheets/sheet1.xml': sheet1,
    'xl/worksheets/_rels/sheet1.xml.rels': rels([['rId1', `${REL}/hyperlink`, 'https://example.com/', true], ['rId2', `${REL}/drawing`, '../drawings/drawing1.xml'], ['rId3', `${REL}/vmlDrawing`, '../drawings/vmlDrawing1.vml'], ['rId4', `${REL}/comments`, '../comments1.xml']]),
    'xl/worksheets/sheet2.xml': sheet2,
    'xl/worksheets/sheet3.xml': sheet3,
    'xl/comments1.xml': comments,
    'xl/drawings/vmlDrawing1.vml': vml,
    'xl/drawings/drawing1.xml': drawing,
    'xl/drawings/_rels/drawing1.xml.rels': rels([['rId1', `${REL}/image`, '../media/image1.png'], ['rId2', `${REL}/chart`, '../charts/chart1.xml']]),
    'xl/charts/chart1.xml': chart,
    'xl/media/image1.png': image,
  });
}

// ================================================================================================
// modern.xlsx: threaded comments with replies, dynamic arrays, shared formulas, rich text, theme colors
{
  const shared = new Shared();
  const s = new SheetBuilder(shared);
  s.s(1, 0, 'Item'); s.s(1, 1, 'Qty'); s.s(1, 2, 'Double');
  const items = [['pear', 3], ['apple', 5], ['fig', 2], ['apple', 1]];
  items.forEach(([name, qty], i) => {
    s.s(2 + i, 0, name); s.n(2 + i, 1, qty);
    if (i === 0) s.raw(2, 2, '<c r="C2"><f t="shared" ref="C2:C5" si="0">B2*2</f><v>6</v></c>');
    else s.raw(2 + i, 2, `<c r="C${2 + i}"><f t="shared" si="0"/><v>${qty * 2}</v></c>`);
  });
  s.s(1, 4, 'Sorted unique');
  s.raw(2, 4, '<c r="E2" cm="1" t="str"><f t="array" ref="E2:E4">_xlfn._xlws.SORT(_xlfn.UNIQUE(A2:A5))</f><v>apple</v></c>');
  s.raw(3, 4, '<c r="E3" t="str"><v>fig</v></c>');
  s.raw(4, 4, '<c r="E4" t="str"><v>pear</v></c>');
  s.s(1, 6, 'Lookup');
  s.raw(2, 6, '<c r="G2"><f>_xlfn.XLOOKUP("fig",A2:A5,B2:B5)</f><v>2</v></c>');
  s.raw(3, 6, '<c r="G3"><f>SUM(_xlfn.ANCHORARRAY(E2))</f><v>0</v></c>');
  s.raw(4, 6, '<c r="G4" t="inlineStr"><is><r><rPr><b/><color rgb="FFFF0000"/></rPr><t>Rich</t></r><r><t xml:space="preserve"> text</t></r></is></c>');
  s.raw(5, 6, '<c r="G5" t="b"><f>ISFORMULA(G2)</f><v>1</v></c>');
  const sheet = s.xml();
  const threaded = `${DECL}<ThreadedComments xmlns="http://schemas.microsoft.com/office/spreadsheetml/2018/threadedcomments" xmlns:x="${MAIN}"><threadedComment ref="A2" dT="2024-03-01T10:00:00.00" personId="{11111111-1111-4111-8111-111111111111}" id="{AAAAAAAA-0000-4000-8000-000000000001}"><text>Is pear in stock?</text></threadedComment><threadedComment ref="A2" dT="2024-03-01T11:00:00.00" personId="{22222222-2222-4222-8222-222222222222}" id="{AAAAAAAA-0000-4000-8000-000000000002}" parentId="{AAAAAAAA-0000-4000-8000-000000000001}"><text>Yes, three left.</text></threadedComment><threadedComment ref="B3" dT="2024-03-02T09:30:00.00" personId="{11111111-1111-4111-8111-111111111111}" id="{AAAAAAAA-0000-4000-8000-000000000003}" done="1"><text>Resolved thread</text></threadedComment></ThreadedComments>`;
  const persons = `${DECL}<personList xmlns="http://schemas.microsoft.com/office/spreadsheetml/2018/threadedcomments" xmlns:x="${MAIN}"><person displayName="Ana Lee" id="{11111111-1111-4111-8111-111111111111}" userId="ana" providerId="None"/><person displayName="Bo Chen" id="{22222222-2222-4222-8222-222222222222}" userId="bo" providerId="None"/></personList>`;
  const legacy = `${DECL}<comments xmlns="${MAIN}"><authors><author>tc={AAAAAAAA-0000-4000-8000-000000000001}</author><author>tc={AAAAAAAA-0000-4000-8000-000000000003}</author></authors><commentList><comment ref="A2" authorId="0"><text><t>[Threaded comment] Is pear in stock?</t></text></comment><comment ref="B3" authorId="1"><text><t>[Threaded comment] Resolved thread</t></text></comment></commentList></comments>`;
  const metadata = `${DECL}<metadata xmlns="${MAIN}" xmlns:xda="http://schemas.microsoft.com/office/spreadsheetml/2017/dynamicarray"><metadataTypes count="1"><metadataType name="XLDAPR" minSupportedVersion="120000" copy="1" pasteAll="1" pasteValues="1" merge="1" splitFirst="1" rowColShift="1" clearFormats="1" clearComments="1" assign="1" coerce="1" cellMeta="1"/></metadataTypes><futureMetadata name="XLDAPR" count="1"><bk><extLst><ext uri="{bdbb8cdc-fa1e-496e-a857-3c3f30c029c3}"><xda:dynamicArrayProperties fDynamic="1" fCollapsed="0"/></ext></extLst></bk></futureMetadata><cellMetadata count="1"><bk><rc t="1" v="0"/></bk></cellMetadata></metadata>`;
  save('modern.xlsx', {
    '[Content_Types].xml': contentTypes([
      ['/xl/workbook.xml', CT.wb], ['/xl/worksheets/sheet1.xml', CT.ws], ['/xl/styles.xml', CT.styles], ['/xl/sharedStrings.xml', CT.sst],
      ['/xl/comments1.xml', CT.comments], ['/xl/threadedComments/threadedComment1.xml', CT.threaded], ['/xl/persons/person.xml', CT.person],
      ['/xl/metadata.xml', CT.metadata], ['/docProps/core.xml', CT.core], ['/docProps/app.xml', CT.app],
    ], [['vml', 'application/vnd.openxmlformats-officedocument.vmlDrawing']]),
    '_rels/.rels': ROOT_RELS,
    'docProps/core.xml': CORE,
    'docProps/app.xml': APP,
    'xl/workbook.xml': `${DECL}<workbook xmlns="${MAIN}" xmlns:r="${R}"><workbookPr/><bookViews><workbookView/></bookViews><sheets><sheet name="Stock" sheetId="1" r:id="rId1"/></sheets><calcPr calcId="191029"/></workbook>`,
    'xl/_rels/workbook.xml.rels': rels([['rId1', `${REL}/worksheet`, 'worksheets/sheet1.xml'], ['rId2', `${REL}/styles`, 'styles.xml'], ['rId3', `${REL}/sharedStrings`, 'sharedStrings.xml'], ['rId4', 'http://schemas.microsoft.com/office/2017/10/relationships/person', 'persons/person.xml'], ['rId5', `${REL}/sheetMetadata`, 'metadata.xml']]),
    'xl/styles.xml': STYLES,
    'xl/sharedStrings.xml': shared.xml(),
    'xl/metadata.xml': metadata,
    'xl/worksheets/sheet1.xml': sheet.replace('</sheetData>', '</sheetData><legacyDrawing r:id="rId3"/>'),
    'xl/worksheets/_rels/sheet1.xml.rels': rels([['rId1', `${REL}/comments`, '../comments1.xml'], ['rId2', 'http://schemas.microsoft.com/office/2017/10/relationships/threadedComment', '../threadedComments/threadedComment1.xml'], ['rId3', `${REL}/vmlDrawing`, '../drawings/vmlDrawing1.vml']]),
    'xl/comments1.xml': legacy,
    'xl/threadedComments/threadedComment1.xml': threaded,
    'xl/persons/person.xml': persons,
    'xl/drawings/vmlDrawing1.vml': '<xml xmlns:v="urn:schemas-microsoft-com:vml" xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:x="urn:schemas-microsoft-com:office:excel"></xml>',
  });
}

// ================================================================================================
// simple.csv
writeFileSync(new URL('simple.csv', OUT), 'Name,Amount,Date,Note\n"Smith, Jane",1200.50,2024-01-15,"Says ""hi"""\nLee,-35,2024-02-01,\nKim,1e3,3/4/2024,multi\n');
console.log('wrote simple.csv');
