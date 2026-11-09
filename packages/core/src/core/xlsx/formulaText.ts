/**
 * Conversion between formulas as stored in files (with _xlfn. prefixes, _xlfn.ANCHORARRAY for
 * the spill operator and _xlfn.SINGLE for '@') and formulas as shown to users.
 */
import { tokenize, type Token } from '../formula/tokenizer';

/** Functions newer than Excel 2007 that must be saved with the _xlfn. prefix. */
const FUTURE = new Set(
  (
    'AGGREGATE ARABIC ARRAYTOTEXT BASE BETA.DIST BETA.INV BINOM.DIST BINOM.DIST.RANGE BINOM.INV BITAND BITLSHIFT BITOR ' +
    'BITRSHIFT BITXOR BYCOL BYROW CEILING.MATH CEILING.PRECISE CHISQ.DIST CHISQ.DIST.RT CHISQ.INV CHISQ.INV.RT CHISQ.TEST ' +
    'CHOOSECOLS CHOOSEROWS COMBINA CONCAT CONFIDENCE.NORM CONFIDENCE.T COT COTH COVARIANCE.P COVARIANCE.S CSC CSCH DAYS ' +
    'DECIMAL DROP ERF.PRECISE ERFC.PRECISE EXPAND EXPON.DIST F.DIST F.DIST.RT F.INV F.INV.RT F.TEST FILTERXML ' +
    'FLOOR.MATH FLOOR.PRECISE FORECAST.ETS FORECAST.ETS.CONFINT FORECAST.ETS.SEASONALITY FORECAST.ETS.STAT FORECAST.LINEAR ' +
    'FORMULATEXT GAMMA GAMMA.DIST GAMMA.INV GAMMALN.PRECISE GAUSS HSTACK HYPGEOM.DIST IFNA IFS IMAGE IMCOSH IMCOT IMCSC ' +
    'IMCSCH IMSEC IMSECH IMSINH IMTAN ISFORMULA ISO.CEILING ISOMITTED ISOWEEKNUM LAMBDA LET LOGNORM.DIST LOGNORM.INV ' +
    'MAKEARRAY MAP MAXIFS MINIFS MODE.MULT MODE.SNGL MUNIT NEGBINOM.DIST NETWORKDAYS.INTL NORM.DIST NORM.INV NORM.S.DIST ' +
    'NORM.S.INV NUMBERVALUE PDURATION PERCENTILE.EXC PERCENTILE.INC PERCENTRANK.EXC PERCENTRANK.INC PERMUTATIONA PHI ' +
    'POISSON.DIST QUARTILE.EXC QUARTILE.INC RANDARRAY RANK.AVG RANK.EQ REDUCE REGEXEXTRACT REGEXREPLACE REGEXTEST RRI ' +
    'SCAN SEC SECH SEQUENCE SHEET SHEETS SKEW.P SORTBY STDEV.P STDEV.S SWITCH T.DIST T.DIST.2T T.DIST.RT T.INV T.INV.2T ' +
    'T.TEST TAKE TEXTAFTER TEXTBEFORE TEXTJOIN TEXTSPLIT TOCOL TOROW UNICHAR UNICODE UNIQUE VALUETOTEXT VAR.P VAR.S VSTACK ' +
    'WEBSERVICE WEIBULL.DIST WORKDAY.INTL WRAPCOLS WRAPROWS XLOOKUP XMATCH XOR Z.TEST ERROR.TYPE'
  ).split(' '),
);
/** Functions saved with the worksheet namespace prefix (_xlfn._xlws.). */
const WORKSHEET_NS = new Set(['FILTER', 'SORT']);

function stripPrefix(name: string): string {
  return name.replace(/^_xlfn\./i, '').replace(/^_xlws\./i, '').replace(/^_xll\./i, '');
}

/** Index of the token closing the parenthesis opened right after tokens[start] (a func token). */
function closingParen(tokens: Token[], openIndex: number): number {
  let depth = 0;
  for (let i = openIndex; i < tokens.length; i++) {
    if (tokens[i].type === 'open') depth++;
    else if (tokens[i].type === 'close') {
      depth--;
      if (depth === 0) return i;
    }
  }
  return -1;
}

/** Formula from a file -> formula for display and editing. */
export function formulaFromFile(text: string): string {
  if (!/_xl/i.test(text)) return text;
  let tokens: Token[];
  try {
    tokens = tokenize(text);
  } catch {
    return text;
  }
  let out = '';
  for (let i = 0; i < tokens.length; i++) {
    const t = tokens[i];
    if (t.type === 'func') {
      const name = stripPrefix(String(t.value));
      const upper = name.toUpperCase();
      if (upper === 'ANCHORARRAY' || upper === 'SINGLE') {
        const close = closingParen(tokens, i + 1);
        if (close > i) {
          const inner = tokens.slice(i + 2, close).map((x) => x.text).join('');
          const innerTokens = tokens.slice(i + 2, close).filter((x) => x.type !== 'ws');
          const simple = innerTokens.length === 1;
          out += upper === 'ANCHORARRAY' ? `${formulaFromFile(inner)}#` : simple ? `@${formulaFromFile(inner)}` : `@(${formulaFromFile(inner)})`;
          i = close;
          continue;
        }
      }
      out += name;
      continue;
    }
    out += t.text;
  }
  return out;
}

const FN_CALL = /([A-Za-z_][A-Za-z0-9_.]*)\s*\(/g;

/** Cheap check: only spill/@ operators and newer functions need rewriting (false positives are fine). */
function mayNeedRewrite(text: string): boolean {
  if (text.includes('#') || text.includes('@')) return true;
  for (const m of text.matchAll(FN_CALL)) {
    const upper = m[1].toUpperCase();
    if (FUTURE.has(upper) || WORKSHEET_NS.has(upper)) return true;
  }
  return false;
}

/** Formula as typed by the user -> formula as stored in a file. */
export function formulaToFile(text: string): string {
  if (!mayNeedRewrite(text)) return text;
  let tokens: Token[];
  try {
    tokens = tokenize(text);
  } catch {
    return text;
  }
  let out = '';
  for (let i = 0; i < tokens.length; i++) {
    const t = tokens[i];
    if (t.type === 'func') {
      const raw = String(t.value);
      if (/^_xl/i.test(raw)) {
        out += raw;
        continue;
      }
      const upper = raw.toUpperCase();
      if (WORKSHEET_NS.has(upper)) out += `_xlfn._xlws.${upper}`;
      else if (FUTURE.has(upper)) out += `_xlfn.${upper}`;
      else out += raw;
      continue;
    }
    if (t.type === 'spill') {
      // The preceding reference becomes _xlfn.ANCHORARRAY(ref).
      const prev = out.lastIndexOf(tokens[i - 1]?.text ?? '');
      if (prev >= 0 && tokens[i - 1]) {
        out = `${out.slice(0, prev)}_xlfn.ANCHORARRAY(${tokens[i - 1].text})`;
        continue;
      }
    }
    if (t.type === 'at') {
      const next = tokens[i + 1];
      if (next && (next.type === 'ref' || next.type === 'name')) {
        out += `_xlfn.SINGLE(${next.text})`;
        i++;
        continue;
      }
      if (next && (next.type === 'func' || next.type === 'open')) {
        const openIdx = next.type === 'func' ? i + 2 : i + 1;
        const close = closingParen(tokens, openIdx);
        if (close > 0) {
          const inner = formulaToFile(tokens.slice(i + 1, close + 1).map((x) => x.text).join(''));
          out += `_xlfn.SINGLE(${inner})`;
          i = close;
          continue;
        }
      }
    }
    out += t.text;
  }
  return out;
}

/** True when the formula uses a function that only exists in Excel 365 (dynamic arrays). */
export function usesDynamicFunctions(text: string): boolean {
  return /\b(FILTER|SORT|SORTBY|UNIQUE|SEQUENCE|RANDARRAY|XLOOKUP|XMATCH|LET|TEXTSPLIT|VSTACK|HSTACK|TOCOL|TOROW|WRAPROWS|WRAPCOLS|TAKE|DROP|CHOOSEROWS|CHOOSECOLS|EXPAND)\s*\(/i.test(text) || /#/.test(text);
}
