/**
 * Cell references inside chart parts (`<c:f>`, and `<cx:f>`/`<c15:f>` in newer chart XML). The
 * editor keeps them in the workbook model so that renaming or deleting sheets and inserting or
 * deleting rows and columns update charts the way they update cell formulas.
 */
import { decodeEntities, escapeText } from '../xml/xml';
import type { OoxmlPackage } from './package';

const F_RE = /<((?:[A-Za-z][\w]*:)?f)(\s[^>]*)?>([^<]*)<\/\1>/g;

/** Formula texts of a chart part, in document order. */
export function readChartFormulas(xml: string): string[] {
  const out: string[] = [];
  for (const m of xml.matchAll(F_RE)) out.push(decodeEntities(m[3]));
  return out;
}

/** The chart part with its formulas replaced (in document order) by `formulas`. */
export function writeChartFormulas(xml: string, formulas: string[]): string {
  let i = 0;
  return xml.replace(F_RE, (whole, tag: string, attrs: string | undefined, text: string) => {
    const f = formulas[i++];
    if (f === undefined || f === decodeEntities(text)) return whole;
    return `<${tag}${attrs ?? ''}>${escapeText(f)}</${tag}>`;
  });
}

/** Chart parts of a package (classic charts and chartEx), keyed by part path. */
export function readPackageChartRefs(pkg: OoxmlPackage): Map<string, string[]> {
  const out = new Map<string, string[]>();
  for (const path of pkg.paths()) {
    if (!/^xl\/charts\/chart(Ex)?\d*\.xml$/i.test(path)) continue;
    const xml = pkg.getText(path);
    if (!xml) continue;
    const refs = readChartFormulas(xml);
    if (refs.length) out.set(path, refs);
  }
  return out;
}
