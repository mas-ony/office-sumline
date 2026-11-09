/** Clipboard formats shared with other spreadsheet apps: tab separated text and HTML tables. */

export function toTsv(rows: string[][]): string {
  return rows
    .map((row) =>
      row
        .map((v) => (/[\t\n\r"]/.test(v) ? `"${v.replace(/"/g, '""')}"` : v))
        .join('\t'),
    )
    .join('\n');
}

/** Parses tab (or comma, for CSV) separated text with quoted fields. */
export function parseDelimited(text: string, delimiter = '\t'): string[][] {
  const rows: string[][] = [];
  let row: string[] = [];
  let field = '';
  let i = 0;
  let quoted = false;
  let fieldStart = true;
  const t = text.replace(/\r\n?/g, '\n');
  while (i < t.length) {
    const ch = t[i];
    if (quoted) {
      if (ch === '"') {
        if (t[i + 1] === '"') {
          field += '"';
          i += 2;
          continue;
        }
        quoted = false;
        i++;
        continue;
      }
      field += ch;
      i++;
      continue;
    }
    if (ch === '"' && fieldStart) {
      quoted = true;
      fieldStart = false;
      i++;
      continue;
    }
    if (ch === delimiter) {
      row.push(field);
      field = '';
      fieldStart = true;
      i++;
      continue;
    }
    if (ch === '\n') {
      row.push(field);
      rows.push(row);
      row = [];
      field = '';
      fieldStart = true;
      i++;
      continue;
    }
    field += ch;
    fieldStart = false;
    i++;
  }
  if (field !== '' || row.length || !fieldStart) {
    row.push(field);
    rows.push(row);
  }
  return rows;
}

/** Guesses the delimiter of CSV-like text. */
export function sniffDelimiter(text: string): string {
  const sample = text.split(/\r?\n/).slice(0, 10).join('\n');
  const candidates = [',', ';', '\t', '|'];
  let best = ',';
  let bestScore = -1;
  for (const d of candidates) {
    const counts = parseDelimited(sample, d).map((r) => r.length).filter((n) => n > 1);
    if (!counts.length) continue;
    const consistent = counts.every((n) => n === counts[0]) ? 2 : 1;
    const score = counts[0] * consistent;
    if (score > bestScore) {
      bestScore = score;
      best = d;
    }
  }
  return best;
}

export interface HtmlCellStyle {
  bold?: boolean;
  italic?: boolean;
  underline?: boolean;
  strike?: boolean;
  color?: string;
  background?: string;
  align?: 'left' | 'center' | 'right';
  fontSize?: number;
  fontFamily?: string;
}

export interface HtmlTable {
  rows: { text: string; style: HtmlCellStyle; colspan: number; rowspan: number }[][];
}

function escapeHtml(s: string): string {
  return s.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
}

export interface HtmlOutCell {
  text: string;
  style?: HtmlCellStyle;
  colspan?: number;
  rowspan?: number;
  /** Covered by a merged cell. */
  hidden?: boolean;
}

export function toHtml(rows: HtmlOutCell[][]): string {
  const body = rows
    .map((row) => {
      const cells = row
        .filter((c) => !c.hidden)
        .map((c) => {
          const st = c.style ?? {};
          const css: string[] = [];
          if (st.bold) css.push('font-weight:bold');
          if (st.italic) css.push('font-style:italic');
          const deco = [st.underline ? 'underline' : '', st.strike ? 'line-through' : ''].filter(Boolean).join(' ');
          if (deco) css.push(`text-decoration:${deco}`);
          if (st.color) css.push(`color:${st.color}`);
          if (st.background) css.push(`background-color:${st.background}`);
          if (st.align) css.push(`text-align:${st.align}`);
          if (st.fontSize) css.push(`font-size:${st.fontSize}pt`);
          if (st.fontFamily) css.push(`font-family:${st.fontFamily}`);
          const span = `${c.colspan && c.colspan > 1 ? ` colspan="${c.colspan}"` : ''}${c.rowspan && c.rowspan > 1 ? ` rowspan="${c.rowspan}"` : ''}`;
          return `<td${span}${css.length ? ` style="${css.join(';')}"` : ''}>${escapeHtml(c.text).replace(/\n/g, '<br>')}</td>`;
        })
        .join('');
      return `<tr>${cells}</tr>`;
    })
    .join('');
  return `<meta charset="utf-8"><table border="0" cellspacing="0" cellpadding="0" style="border-collapse:collapse">${body}</table>`;
}

function readStyle(el: Element): HtmlCellStyle {
  const style: HtmlCellStyle = {};
  const s = (el as HTMLElement).style;
  const weight = s?.fontWeight;
  if (weight === 'bold' || Number(weight) >= 600 || el.querySelector('b,strong')) style.bold = true;
  if (s?.fontStyle === 'italic' || el.querySelector('i,em')) style.italic = true;
  const deco = `${s?.textDecoration ?? ''} ${s?.textDecorationLine ?? ''}`;
  if (deco.includes('underline') || el.querySelector('u')) style.underline = true;
  if (deco.includes('line-through') || el.querySelector('s,strike,del')) style.strike = true;
  if (s?.color) style.color = s.color;
  const bg = s?.backgroundColor || el.getAttribute('bgcolor') || '';
  if (bg && bg !== 'transparent') style.background = bg;
  const align = (s?.textAlign || el.getAttribute('align') || '').toLowerCase();
  if (align === 'left' || align === 'center' || align === 'right') style.align = align;
  const size = s?.fontSize;
  if (size) {
    const m = /([\d.]+)(pt|px)/.exec(size);
    if (m) style.fontSize = m[2] === 'px' ? Math.round(Number(m[1]) * 0.75) : Number(m[1]);
  }
  if (s?.fontFamily) style.fontFamily = s.fontFamily.split(',')[0].replace(/["']/g, '').trim();
  return style;
}

/** Reads the first table of clipboard HTML (Excel, Google Sheets, web pages). */
export function parseHtmlTable(html: string): HtmlTable | null {
  if (typeof DOMParser === 'undefined') return null;
  const doc = new DOMParser().parseFromString(html, 'text/html');
  const table = doc.querySelector('table');
  if (!table) return null;
  const rows: HtmlTable['rows'] = [];
  for (const tr of Array.from(table.querySelectorAll('tr'))) {
    const row: HtmlTable['rows'][number] = [];
    for (const td of Array.from(tr.children)) {
      if (td.tagName !== 'TD' && td.tagName !== 'TH') continue;
      const clone = td.cloneNode(true) as HTMLElement;
      for (const br of Array.from(clone.querySelectorAll('br'))) br.replaceWith('\n');
      const text = (clone.textContent ?? '').replace(/ /g, ' ');
      row.push({
        text: text.replace(/^\n+|\n+$/g, ''),
        style: readStyle(td),
        colspan: Number(td.getAttribute('colspan') ?? 1) || 1,
        rowspan: Number(td.getAttribute('rowspan') ?? 1) || 1,
      });
    }
    rows.push(row);
  }
  return rows.length ? { rows } : null;
}
