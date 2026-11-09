import { CellError, Errors, type Scalar } from '../../values';
import { ArrayValue, RefValue, toNumber, toText, type Value } from '../runtime';
import { makeCriteria } from '../criteria';
import { define, toArray, toScalar, type FnContext } from './helpers';
import { wrap } from './fjs';
import { mean, stdev, sum, variance } from './numeric';

for (const name of ['BIN2DEC', 'BIN2HEX', 'BIN2OCT', 'DEC2BIN', 'DEC2HEX', 'DEC2OCT', 'HEX2BIN', 'HEX2DEC', 'HEX2OCT', 'OCT2BIN', 'OCT2DEC', 'OCT2HEX']) {
  wrap(name, name, name.endsWith('DEC') ? 'v' : 'v,n?');
}
for (const name of ['BITAND', 'BITLSHIFT', 'BITOR', 'BITRSHIFT', 'BITXOR']) wrap(name, name, 'n,n');
wrap('COMPLEX', 'COMPLEX', 'n,n,s?');
for (const name of [
  'IMABS', 'IMAGINARY', 'IMARGUMENT', 'IMCONJUGATE', 'IMCOS', 'IMCOSH', 'IMCOT', 'IMCSC', 'IMCSCH', 'IMEXP', 'IMLN', 'IMLOG10',
  'IMLOG2', 'IMREAL', 'IMSEC', 'IMSECH', 'IMSIN', 'IMSINH', 'IMSQRT', 'IMTAN',
]) wrap(name, name, 'v');
wrap('IMDIV', 'IMDIV', 'v,v');
wrap('IMSUB', 'IMSUB', 'v,v');
wrap('IMPOWER', 'IMPOWER', 'v,n');
wrap('IMPRODUCT', 'IMPRODUCT', 'v*');
wrap('IMSUM', 'IMSUM', 'v*');
wrap('CONVERT', 'CONVERT', 'n,s,s');
wrap('DELTA', 'DELTA', 'n,n?');
wrap('GESTEP', 'GESTEP', 'n,n?');
wrap('BESSELI', 'BESSELI', 'n,n');
wrap('BESSELJ', 'BESSELJ', 'n,n');
wrap('BESSELK', 'BESSELK', 'n,n');
wrap('BESSELY', 'BESSELY', 'n,n');

// -----------------------------------------------------------------------------------------------
// Database functions

/** Rows (as value arrays) of `database` that satisfy the criteria table; the header row is excluded. */
function matchingRecords(dbArg: Value, critArg: Value, ctx: FnContext): { rows: Scalar[][]; headers: string[] } | CellError {
  const db = toArray(dbArg, ctx);
  const crit = toArray(critArg, ctx);
  if (db.rows < 1 || crit.rows < 1) return Errors.VALUE;
  const headers = db.data[0].map((h) => String(toText(h ?? '')).toLowerCase());
  const critHeaders = crit.data[0].map((h) => String(toText(h ?? '')).toLowerCase());
  const conditions = crit.data.slice(1).map((row) =>
    row
      .map((v, i) => ({ col: headers.indexOf(critHeaders[i]), v }))
      .filter((c) => c.v !== null && c.v !== '')
      .map((c) => ({ col: c.col, match: makeCriteria(c.v) })),
  );
  const rows = db.data.slice(1).filter((row) => {
    if (!conditions.length) return true;
    return conditions.some((conds) => conds.every((c) => c.col >= 0 && c.match(row[c.col])));
  });
  return { rows, headers };
}

function fieldIndex(field: Value, headers: string[], ctx: FnContext): number | CellError {
  if (field === null) return -1;
  const v = toScalar(field, ctx);
  if (typeof v === 'number') {
    const k = Math.trunc(v);
    return k >= 1 && k <= headers.length ? k - 1 : Errors.VALUE;
  }
  const name = String(toText(v ?? '')).toLowerCase();
  const i = headers.indexOf(name);
  return i < 0 ? Errors.VALUE : i;
}

function databaseFn(name: string, reduce: (values: Scalar[]) => Value, needsField = true): void {
  define(name, {
    min: needsField ? 3 : 2,
    max: 3,
    impl: (args, ctx) => {
      const rec = matchingRecords(args[0], args[2] ?? args[1], ctx);
      if (rec instanceof CellError) return rec;
      if (args[1] === null && !needsField) return reduce(rec.rows.map(() => 1));
      const idx = fieldIndex(args[1], rec.headers, ctx);
      if (idx instanceof CellError) return idx;
      if (idx < 0) return reduce(rec.rows.map(() => 1));
      return reduce(rec.rows.map((r) => r[idx]));
    },
  });
}

const numbers = (vs: Scalar[]) => vs.filter((v): v is number => typeof v === 'number');
databaseFn('DSUM', (vs) => sum(numbers(vs)));
databaseFn('DAVERAGE', (vs) => mean(numbers(vs)));
databaseFn('DCOUNT', (vs) => numbers(vs).length, false);
databaseFn('DCOUNTA', (vs) => vs.filter((v) => v !== null && v !== '').length, false);
databaseFn('DMAX', (vs) => (numbers(vs).length ? Math.max(...numbers(vs)) : 0));
databaseFn('DMIN', (vs) => (numbers(vs).length ? Math.min(...numbers(vs)) : 0));
databaseFn('DPRODUCT', (vs) => (numbers(vs).length ? numbers(vs).reduce((a, b) => a * b, 1) : 0));
databaseFn('DSTDEV', (vs) => stdev(numbers(vs), true));
databaseFn('DSTDEVP', (vs) => stdev(numbers(vs), false));
databaseFn('DVAR', (vs) => variance(numbers(vs), true));
databaseFn('DVARP', (vs) => variance(numbers(vs), false));
databaseFn('DGET', (vs) => (vs.length === 0 ? Errors.VALUE : vs.length > 1 ? Errors.NUM : (vs[0] ?? 0)));

void ArrayValue;
void RefValue;
void toNumber;
