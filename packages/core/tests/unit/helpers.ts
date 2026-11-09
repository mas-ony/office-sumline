import { readFileSync } from 'node:fs';
import { dirname, resolve } from 'node:path';
import { fileURLToPath } from 'node:url';
import { Workbook } from '../../src/core/model/workbook';
import { CalcEngine } from '../../src/core/formula/engine';
import { parseInput } from '../../src/core/format/input';
import { parseCellRef, splitSheetRef } from '../../src/core/refs';
import type { Scalar } from '../../src/core/values';
import { CellError } from '../../src/core/values';

/** A file of the repository's test fixtures (tests/fixtures at the root). */
export function readFixture(name: string): Uint8Array {
  return new Uint8Array(readFileSync(resolve(dirname(fileURLToPath(import.meta.url)), '../../../../tests/fixtures', name)));
}

export interface Book {
  wb: Workbook;
  engine: CalcEngine;
  /** Sets cells from a map of 'A1' -> input text (formulas start with '='). */
  set(cells: Record<string, string | number | boolean | null>, sheetName?: string): void;
  get(ref: string): Scalar;
  recalc(): void;
}

export function makeBook(sheetNames: string[] = ['Sheet1']): Book {
  const wb = new Workbook();
  for (const n of sheetNames) wb.addSheet(n);
  const engine = new CalcEngine(wb);
  engine.rebuild();
  const book: Book = {
    wb,
    engine,
    set(cells, sheetName) {
      for (const [ref, input] of Object.entries(cells)) {
        const { sheet: qual, ref: cellRef } = splitSheetRef(ref);
        const sheet = wb.sheetByName(qual ?? sheetName ?? sheetNames[0])!;
        const a = parseCellRef(cellRef)!;
        if (input === null) sheet.setCell(a.r, a.c, undefined);
        else if (typeof input === 'string') {
          const p = parseInput(input);
          sheet.setCell(a.r, a.c, p.formula !== undefined ? { v: null, f: p.formula, dyn: true } : { v: p.value });
        } else sheet.setCell(a.r, a.c, { v: input });
        engine.cellChanged(sheet, a.r, a.c);
      }
      engine.recalc();
    },
    get(ref) {
      const { sheet: qual, ref: cellRef } = splitSheetRef(ref);
      const sheet = wb.sheetByName(qual ?? sheetNames[0])!;
      const a = parseCellRef(cellRef)!;
      const v = sheet.getCell(a.r, a.c)?.v ?? null;
      return v;
    },
    recalc() {
      engine.recalc();
    },
  };
  return book;
}

/** Value of a formula evaluated in a fresh one-cell workbook (plus optional data). */
export function evalFormula(formula: string, data: Record<string, string | number | boolean | null> = {}): Scalar {
  const b = makeBook();
  b.set(data);
  b.set({ Z100: formula.startsWith('=') ? formula : `=${formula}` });
  return b.get('Z100');
}

export function errCode(v: Scalar): string | null {
  return v instanceof CellError ? v.code : null;
}
