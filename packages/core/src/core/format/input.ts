/** Interprets what a user typed into a cell, the way Excel does. */
import { parseValue } from 'numfmt';
import { parseError, type Scalar } from '../values';

export interface ParsedInput {
  value: Scalar;
  /** Formula without '=' when the input is a formula. */
  formula?: string;
  /** Number format implied by the input ("10%" -> 0%, "$1,000" -> currency). */
  numFmt?: string;
  /** Input started with an apostrophe: always text. */
  quotePrefix?: boolean;
}

export function parseInput(text: string, opts: { textFormat?: boolean; date1904?: boolean } = {}): ParsedInput {
  if (text === '') return { value: null };
  if (opts.textFormat) return { value: text };
  if (text.startsWith("'")) return { value: text.slice(1), quotePrefix: true };
  if (text.length > 1 && text.startsWith('=')) return { value: null, formula: text.slice(1) };
  if (text.length > 1 && (text.startsWith('+') || text.startsWith('-')) && /^[+-]\s*[A-Za-z(]/.test(text) && !/^[+-]\s*(TRUE|FALSE)$/i.test(text)) {
    return { value: null, formula: text };
  }
  const trimmed = text.trim();
  const upper = trimmed.toUpperCase();
  if (upper === 'TRUE') return { value: true };
  if (upper === 'FALSE') return { value: false };
  const err = parseError(trimmed);
  if (err) return { value: err };
  // Numbers, percentages, currency, dates and times.
  if (/[0-9]/.test(trimmed)) {
    const parsed = parseValue(trimmed);
    if (parsed && typeof parsed.v === 'number' && Number.isFinite(parsed.v)) {
      let v = parsed.v;
      if (opts.date1904 && parsed.z && /[dmy]/i.test(parsed.z.replace(/"[^"]*"/g, ''))) v -= 1462;
      return { value: v, numFmt: parsed.z };
    }
  }
  return { value: text };
}
