/** Criteria strings of COUNTIF, SUMIF, AVERAGEIF, the *IFS functions and the database functions. */
import { CellError, parseError, type Scalar } from '../values';
import { textToNumber } from './runtime';

export type Matcher = (v: Scalar) => boolean;

/** Converts an Excel wildcard pattern (* ? ~) into a case-insensitive regular expression. */
export function wildcardToRegExp(pattern: string, whole = true): RegExp {
  let re = '';
  for (let i = 0; i < pattern.length; i++) {
    const ch = pattern[i];
    if (ch === '~' && i + 1 < pattern.length && '*?~'.includes(pattern[i + 1])) {
      re += escapeRe(pattern[++i]);
    } else if (ch === '*') re += '[\\s\\S]*';
    else if (ch === '?') re += '[\\s\\S]';
    else re += escapeRe(ch);
  }
  return new RegExp(whole ? `^${re}$` : re, 'i');
}

function escapeRe(s: string): string {
  return s.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
}

export function hasWildcards(s: string): boolean {
  return /(^|[^~])[*?]/.test(s);
}

function asNumber(v: Scalar): number | null {
  if (typeof v === 'number') return v;
  if (typeof v === 'string') return textToNumber(v);
  return null;
}

export function makeCriteria(crit: Scalar): Matcher {
  if (crit === null) return (v) => v === null || v === '';
  if (typeof crit === 'number') {
    return (v) => {
      if (typeof v === 'number') return v === crit;
      if (typeof v === 'string') return textToNumber(v) === crit;
      return false;
    };
  }
  if (typeof crit === 'boolean') return (v) => v === crit;
  if (crit instanceof CellError) return (v) => v === crit;

  const m = /^(<=|>=|<>|=|<|>)?([\s\S]*)$/.exec(crit)!;
  const op = m[1] ?? '=';
  const rhs = m[2];

  if (rhs === '') {
    if (op === '=') return (v) => v === null || (m[1] === undefined && v === '');
    if (op === '<>') return (v) => v !== null && v !== '';
    return () => false;
  }

  const err = parseError(rhs);
  if (err) {
    if (op === '=') return (v) => v === err;
    if (op === '<>') return (v) => v !== err;
    return () => false;
  }

  const upper = rhs.toUpperCase();
  if (upper === 'TRUE' || upper === 'FALSE') {
    const b = upper === 'TRUE';
    if (op === '=') return (v) => v === b;
    if (op === '<>') return (v) => v !== b;
  }

  const n = textToNumber(rhs);
  if (n !== null) {
    switch (op) {
      case '=':
        return (v) => asNumber(v) === n && v !== null;
      case '<>':
        return (v) => asNumber(v) !== n || v === null;
      case '<':
        return (v) => typeof v === 'number' && v < n;
      case '>':
        return (v) => typeof v === 'number' && v > n;
      case '<=':
        return (v) => typeof v === 'number' && v <= n;
      case '>=':
        return (v) => typeof v === 'number' && v >= n;
    }
  }

  const lower = rhs.toLowerCase();
  if (op === '=' || op === '<>') {
    const re = hasWildcards(rhs) ? wildcardToRegExp(rhs) : null;
    const test = (v: Scalar) => {
      if (typeof v !== 'string') return false;
      return re ? re.test(v) : v.toLowerCase() === lower.replace(/~([*?~])/g, '$1');
    };
    return op === '=' ? test : (v) => !test(v);
  }
  return (v) => {
    if (typeof v !== 'string') return false;
    const x = v.toLowerCase();
    switch (op) {
      case '<':
        return x < lower;
      case '>':
        return x > lower;
      case '<=':
        return x <= lower;
      default:
        return x >= lower;
    }
  };
}
