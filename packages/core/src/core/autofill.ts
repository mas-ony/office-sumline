/**
 * Fill handle series: numbers and dates continue their trend, text ending in a number counts
 * up, day and month names cycle, formulas are copied with relative references shifted, and
 * anything else repeats.
 */
import type { Cell } from './model/sheet';

const NAME_LISTS = [
  ['sunday', 'monday', 'tuesday', 'wednesday', 'thursday', 'friday', 'saturday'],
  ['sun', 'mon', 'tue', 'wed', 'thu', 'fri', 'sat'],
  ['january', 'february', 'march', 'april', 'may', 'june', 'july', 'august', 'september', 'october', 'november', 'december'],
  ['jan', 'feb', 'mar', 'apr', 'may', 'jun', 'jul', 'aug', 'sep', 'oct', 'nov', 'dec'],
];

function matchCase(sample: string, word: string): string {
  if (sample.length > 1 && sample === sample.toUpperCase()) return word.toUpperCase();
  if (sample[0] === sample[0].toUpperCase()) return word[0].toUpperCase() + word.slice(1);
  return word;
}

function listIndex(text: string): { list: string[]; index: number } | null {
  const lower = text.trim().toLowerCase();
  for (const list of NAME_LISTS) {
    const i = list.indexOf(lower);
    if (i >= 0) return { list, index: i };
  }
  return null;
}

const TRAILING_NUMBER = /^(.*?)(\d+)(\D*)$/;

export interface FillOptions {
  /** Copy instead of creating a series (Ctrl while dragging). */
  copy?: boolean;
  /** True when filling up or left. */
  reverse?: boolean;
  /** Whether a cell holds a date (dates step by one day even when alone). */
  isDate?: (cell: Cell) => boolean;
  /** Shifts a formula by `distance` cells along the fill direction. */
  translate: (formula: string, distance: number) => string;
}

/**
 * Cells for `count` targets continuing `source` (ordered along the fill axis). Target k (1-based)
 * is the k-th cell after the source block, or before it for reverse fills.
 */
export function fillSeries(source: (Cell | undefined)[], count: number, opts: FillOptions): (Cell | undefined)[] {
  const n = source.length;
  const out: (Cell | undefined)[] = [];
  if (!n) return out;
  const plain = source.every((c) => c !== undefined && c.f === undefined && typeof c.v === 'number');
  const nums = plain ? source.map((c) => c!.v as number) : [];
  const singleDate = n === 1 && plain && !!opts.isDate?.(source[0]!);

  for (let k = 1; k <= count; k++) {
    const pos = opts.reverse ? -k : n - 1 + k;
    const srcIndex = ((pos % n) + n) % n;
    const cycles = Math.floor(pos / n);
    const src = source[srcIndex];
    if (!src) {
      out.push(undefined);
      continue;
    }
    if (src.f !== undefined) {
      out.push({ s: src.s, v: null, f: opts.translate(src.f, pos - srcIndex), dyn: src.dyn });
      continue;
    }
    if (opts.copy) {
      out.push({ ...src });
      continue;
    }
    if (plain && (n > 1 || singleDate)) {
      const delta = n > 1 ? (nums[n - 1] - nums[0]) / (n - 1) : 1;
      const v = opts.reverse ? nums[0] - delta * k : nums[n - 1] + delta * k;
      out.push({ ...src, v: Number(v.toPrecision(15)) });
      continue;
    }
    const v = src.v;
    if (typeof v === 'string') {
      const li = listIndex(v);
      if (li) {
        const len = li.list.length;
        const idx = (((li.index + cycles * n) % len) + len) % len;
        out.push({ ...src, v: matchCase(v, li.list[idx]), rt: undefined });
        continue;
      }
      const m = TRAILING_NUMBER.exec(v);
      if (m) {
        const num = Math.max(0, Number(m[2]) + cycles * n);
        out.push({ ...src, v: `${m[1]}${String(num).padStart(m[2].length, '0')}${m[3]}`, rt: undefined });
        continue;
      }
    }
    out.push({ ...src });
  }
  return out;
}
