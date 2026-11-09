/**
 * The calculation engine: keeps a dependency graph of formula cells, recalculates what changed,
 * spills dynamic arrays and distributes legacy array formulas.
 */
import { MAX_COLS, MAX_ROWS, cellKey, keyCol, keyRow, parseRange, rangeContains, splitSheetRef, type Range } from '../refs';
import { CellError, Errors, type Scalar } from '../values';
import type { Cell, Sheet } from '../model/sheet';
import type { Workbook } from '../model/workbook';
import { collectRefs, parseFormula, type Node } from './parser';
import { ArrayValue, RefValue, type Area, type Value } from './runtime';
import { evaluate, resultScalar, type EvalContext } from './evaluator';
import { registry, toArray, type EngineApi, type FnContext } from './functions/helpers';
import { setDate1904 } from './dates';
import './functions';

const SHEET_FACTOR = 2 ** 34;
const DEPTH_LIMIT = 300;
const FULL_RECALC_THRESHOLD = 5000;

export function cellUid(sheet: Sheet, r: number, c: number): number {
  return sheet.uid * SHEET_FACTOR + r * MAX_COLS + c;
}

function uidSheet(uid: number): number {
  return Math.floor(uid / SHEET_FACTOR);
}

function uidKey(uid: number): number {
  return uid % SHEET_FACTOR;
}

interface FormulaEntry {
  sheet: Sheet;
  r: number;
  c: number;
  text: string;
  ast: Node | null;
  /** Legacy formula: references in scalar positions use implicit intersection. */
  legacy: boolean;
  /** Dynamic array formula: array results spill. */
  dynamic: boolean;
  /** Fixed range of a legacy (CSE) array formula. */
  arrayRange?: Range;
  volatile: boolean;
  cellDeps: number[];
  rangeDeps: { sheetUid: number; range: Range }[];
}

class DeepRecursion {
  constructor(readonly uid: number) {}
}

interface RangeBucket {
  range: Range;
  deps: Set<number>;
}

export class CalcEngine implements EngineApi {
  private formulas = new Map<number, FormulaEntry>();
  private cellDependents = new Map<number, Set<number>>();
  private rangeDependents = new Map<number, Map<string, RangeBucket>>();
  private volatileSet = new Set<number>();
  private dirty = new Set<number>();
  private computing = new Set<number>();
  private depth = 0;
  private spills = new Map<number, Range>();
  /** Ranges a dynamic array wanted to spill into but could not (cleared cells unblock them). */
  private blockedSpills = new Map<number, Range>();
  private nameCache = new Map<string, Node | null>();
  private astCache = new Map<string, Node | null>();
  /** Cells whose value changed during the last recalculation. */
  changed = new Set<number>();
  /** Formula cells found in a circular reference during the last recalculation. */
  circular = new Set<number>();
  /** Set while a full recalculation is pending. */
  private fullPending = false;

  constructor(readonly workbook: Workbook) {}

  // ---------------------------------------------------------------------------------------------
  // Registration

  /** Re-registers every formula of the workbook and marks all of them dirty. */
  rebuild(): void {
    this.formulas.clear();
    this.cellDependents.clear();
    this.rangeDependents.clear();
    this.volatileSet.clear();
    this.dirty.clear();
    this.spills.clear();
    this.blockedSpills.clear();
    this.nameCache.clear();
    for (const sheet of this.workbook.sheets) {
      sheet.forEachCell((r, c, cell) => {
        if (cell.f !== undefined) this.register(sheet, r, c, cell);
        if (cell.arr && cell.f !== undefined) {
          const range = parseRange(cell.arr);
          if (range) this.spills.set(cellUid(sheet, r, c), range);
        }
      });
    }
    for (const uid of this.formulas.keys()) this.dirty.add(uid);
    this.fullPending = true;
  }

  private parse(text: string): Node | null {
    let ast = this.astCache.get(text);
    if (ast === undefined) {
      try {
        ast = parseFormula(text);
      } catch {
        ast = null;
      }
      if (this.astCache.size > 20000) this.astCache.clear();
      this.astCache.set(text, ast);
    }
    return ast;
  }

  private register(sheet: Sheet, r: number, c: number, cell: Cell): void {
    const uid = cellUid(sheet, r, c);
    this.unregister(uid);
    const text = cell.f!;
    const ast = this.parse(text);
    const arrayRange = cell.arr && !cell.dyn ? parseRange(cell.arr) ?? undefined : undefined;
    const entry: FormulaEntry = {
      sheet,
      r,
      c,
      text,
      ast,
      legacy: !cell.dyn && !cell.arr,
      dynamic: !!cell.dyn,
      arrayRange,
      volatile: false,
      cellDeps: [],
      rangeDeps: [],
    };
    if (ast) {
      const { refs, names, funcs } = collectRefs(ast);
      for (const f of funcs) {
        const def = registry.get(f);
        if (def?.volatile || f === 'INDIRECT' || f === 'OFFSET') entry.volatile = true;
      }
      const addRef = (targetSheet: Sheet, range: Range) => {
        if (range.r0 === range.r1 && range.c0 === range.c1) entry.cellDeps.push(cellUid(targetSheet, range.r0, range.c0));
        else entry.rangeDeps.push({ sheetUid: targetSheet.uid, range });
      };
      for (const ref of refs) {
        if (ref.kind === 'error' || ref.book !== undefined) continue;
        const range = { r0: Math.min(ref.r0, ref.r1), c0: Math.min(ref.c0, ref.c1), r1: Math.max(ref.r0, ref.r1), c1: Math.max(ref.c0, ref.c1) };
        for (const target of this.sheetsOf(sheet, ref.sheet, ref.sheet2)) addRef(target, range);
      }
      for (const { name, sheet: q } of names) {
        const resolved = this.nameRefs(name, q, sheet);
        if (resolved === 'dynamic') entry.volatile = true;
        else for (const a of resolved) addRef(a.sheet, a.range);
      }
      // Spill references (A1#) and structured references are resolved at run time.
      if (/#|\[/.test(text)) entry.volatile = true;
    }
    this.formulas.set(uid, entry);
    for (const d of entry.cellDeps) {
      let set = this.cellDependents.get(d);
      if (!set) this.cellDependents.set(d, (set = new Set()));
      set.add(uid);
    }
    for (const { sheetUid, range } of entry.rangeDeps) {
      let buckets = this.rangeDependents.get(sheetUid);
      if (!buckets) this.rangeDependents.set(sheetUid, (buckets = new Map()));
      const key = `${range.r0},${range.c0},${range.r1},${range.c1}`;
      let bucket = buckets.get(key);
      if (!bucket) buckets.set(key, (bucket = { range, deps: new Set() }));
      bucket.deps.add(uid);
    }
    if (entry.volatile) this.volatileSet.add(uid);
  }

  private unregister(uid: number): void {
    const entry = this.formulas.get(uid);
    if (!entry) return;
    for (const d of entry.cellDeps) {
      const set = this.cellDependents.get(d);
      set?.delete(uid);
      if (set && !set.size) this.cellDependents.delete(d);
    }
    for (const { sheetUid, range } of entry.rangeDeps) {
      const buckets = this.rangeDependents.get(sheetUid);
      const key = `${range.r0},${range.c0},${range.r1},${range.c1}`;
      const bucket = buckets?.get(key);
      bucket?.deps.delete(uid);
      if (bucket && !bucket.deps.size) buckets!.delete(key);
    }
    this.volatileSet.delete(uid);
    this.formulas.delete(uid);
  }

  private sheetsOf(current: Sheet, name?: string, name2?: string): Sheet[] {
    if (name === undefined) return [current];
    const first = this.workbook.sheetByName(name);
    if (!first) return [];
    if (name2 === undefined) return [first];
    const last = this.workbook.sheetByName(name2);
    if (!last) return [first];
    const a = this.workbook.sheets.indexOf(first);
    const b = this.workbook.sheets.indexOf(last);
    return this.workbook.sheets.slice(Math.min(a, b), Math.max(a, b) + 1);
  }

  /** Static areas referenced by a defined name, or 'dynamic' when they cannot be known upfront. */
  private nameRefs(name: string, qualifier: string | undefined, sheet: Sheet): Area[] | 'dynamic' {
    const def = this.lookupName(name, qualifier, sheet);
    if (!def) return [];
    const ast = this.nameAst(def.formula);
    if (!ast) return [];
    const { refs, funcs, names } = collectRefs(ast);
    if (names.length || [...funcs].some((f) => f === 'INDIRECT' || f === 'OFFSET')) return 'dynamic';
    const out: Area[] = [];
    for (const ref of refs) {
      if (ref.kind === 'error' || ref.book !== undefined) continue;
      const range = { r0: Math.min(ref.r0, ref.r1), c0: Math.min(ref.c0, ref.c1), r1: Math.max(ref.r0, ref.r1), c1: Math.max(ref.c0, ref.c1) };
      for (const s of this.sheetsOf(sheet, ref.sheet, ref.sheet2)) out.push({ sheet: s, range });
    }
    return out;
  }

  private lookupName(name: string, qualifier: string | undefined, sheet: Sheet) {
    if (qualifier !== undefined) {
      const target = this.workbook.sheetByName(qualifier);
      if (!target) return undefined;
      const idx = this.workbook.sheets.indexOf(target);
      return this.workbook.names.find((n) => n.localSheet === idx && n.name.toUpperCase() === name.toUpperCase());
    }
    return this.workbook.nameLookup(name, this.workbook.sheets.indexOf(sheet));
  }

  private nameAst(formula: string): Node | null {
    let ast = this.nameCache.get(formula);
    if (ast === undefined) {
      try {
        ast = parseFormula(formula);
      } catch {
        ast = null;
      }
      this.nameCache.set(formula, ast);
    }
    return ast;
  }

  /**
   * Must be called after the content (value or formula) of a cell changed. Updates the
   * dependency graph and marks the cell's dependents dirty.
   */
  cellChanged(sheet: Sheet, r: number, c: number): void {
    const uid = cellUid(sheet, r, c);
    const cell = sheet.getCell(r, c);
    const hadFormula = this.formulas.has(uid);
    if (cell?.f !== undefined) {
      this.register(sheet, r, c, cell);
      this.dirty.add(uid);
    } else if (hadFormula) {
      this.unregister(uid);
      this.clearSpill(uid);
      this.blockedSpills.delete(uid);
    }
    // A value typed into a spill area blocks the spill.
    for (const map of [this.spills, this.blockedSpills]) {
      for (const [anchor, range] of map) {
        if (uidSheet(anchor) === sheet.uid && anchor !== uid && rangeContains(range, r, c)) this.dirty.add(anchor);
      }
    }
    this.pendingChanges.push(uid);
  }

  private pendingChanges: number[] = [];

  /** Marks all formulas dirty (after structural changes, renames or name edits). */
  invalidateAll(): void {
    this.rebuild();
  }

  // ---------------------------------------------------------------------------------------------
  // Recalculation

  private dependentsOf(uid: number, out: number[]): void {
    const direct = this.cellDependents.get(uid);
    if (direct) for (const d of direct) out.push(d);
    const buckets = this.rangeDependents.get(uidSheet(uid));
    if (!buckets) return;
    const key = uidKey(uid);
    const r = keyRow(key);
    const c = keyCol(key);
    for (const b of buckets.values()) {
      if (r >= b.range.r0 && r <= b.range.r1 && c >= b.range.c0 && c <= b.range.c1) for (const d of b.deps) out.push(d);
    }
  }

  private propagate(start: number[]): void {
    const queue = [...start];
    while (queue.length) {
      const x = queue.pop()!;
      const deps: number[] = [];
      this.dependentsOf(x, deps);
      for (const d of deps) {
        if (!this.dirty.has(d)) {
          this.dirty.add(d);
          queue.push(d);
          // Spill cells of a dirty dynamic formula change too.
          const spill = this.spills.get(d);
          if (spill) {
            const sheet = this.workbook.sheetByUid(uidSheet(d));
            if (sheet) for (let rr = spill.r0; rr <= spill.r1; rr++) for (let cc = spill.c0; cc <= spill.c1; cc++) queue.push(cellUid(sheet, rr, cc));
          }
        }
      }
    }
  }

  /**
   * Recalculates dirty formulas and their dependents. Returns the uids of cells whose value
   * changed (including spill cells).
   */
  recalc(): Set<number> {
    setDate1904(this.workbook.date1904);
    this.changed = new Set();
    this.circular = new Set();
    const start = this.pendingChanges;
    this.pendingChanges = [];
    if (this.fullPending || start.length > FULL_RECALC_THRESHOLD) {
      for (const uid of this.formulas.keys()) this.dirty.add(uid);
      this.fullPending = false;
    } else {
      const seeds = [...start, ...this.dirty];
      this.propagate(seeds);
    }
    for (const v of this.volatileSet) this.dirty.add(v);
    this.propagate([...this.volatileSet]);
    let guard = 0;
    while (this.dirty.size && guard++ < 10) {
      const order = [...this.dirty];
      for (const uid of order) if (this.dirty.has(uid)) this.drive(uid);
    }
    this.dirty.clear();
    return this.changed;
  }

  /** Computes a formula cell iteratively so that long dependency chains do not overflow the stack. */
  private drive(uid: number): void {
    const stack = [uid];
    let steps = 0;
    while (stack.length) {
      const top = stack[stack.length - 1];
      if (!this.dirty.has(top) || !this.formulas.has(top)) {
        this.dirty.delete(top);
        stack.pop();
        continue;
      }
      try {
        this.depth = 0;
        this.compute(top);
        stack.pop();
      } catch (e) {
        if (!(e instanceof DeepRecursion)) throw e;
        if (stack.includes(e.uid) || ++steps > 1e7) {
          // Circular through the explicit stack.
          this.circular.add(e.uid);
          this.writeValue(e.uid, 0);
          this.dirty.delete(e.uid);
        } else stack.push(e.uid);
      }
    }
  }

  private compute(uid: number): void {
    const entry = this.formulas.get(uid);
    if (!entry) {
      this.dirty.delete(uid);
      return;
    }
    if (this.computing.has(uid)) return;
    this.computing.add(uid);
    let done = false;
    try {
      this.depth++;
      if (this.depth > DEPTH_LIMIT) throw new DeepRecursion(uid);
      const ctx = this.context(entry);
      let value: Value;
      if (!entry.ast) value = Errors.NAME;
      else {
        try {
          value = evaluate(entry.ast, ctx);
        } catch (e) {
          if (e instanceof DeepRecursion) throw e;
          value = Errors.VALUE;
        }
      }
      this.store(uid, entry, value, ctx);
      done = true;
    } finally {
      this.depth--;
      this.computing.delete(uid);
      if (done) this.dirty.delete(uid);
    }
  }

  private context(entry: FormulaEntry): EvalContext {
    return {
      engine: this,
      sheet: entry.sheet,
      row: entry.r,
      col: entry.c,
      legacy: entry.legacy,
      resolveName: (name, qualifier, ctx) => this.resolveName(name, qualifier, ctx),
      spillRange: (sheet, r, c) => this.spillRange(sheet, r, c),
    };
  }

  private store(uid: number, entry: FormulaEntry, value: Value, ctx: FnContext): void {
    if (entry.arrayRange) {
      this.distributeArray(uid, entry, value, ctx, entry.arrayRange);
      return;
    }
    const multi = value instanceof ArrayValue ? value.rows * value.cols > 1 : value instanceof RefValue ? !value.isCell() : false;
    if (entry.dynamic && multi) {
      this.spill(uid, entry, value, ctx);
      return;
    }
    if (this.spills.has(uid)) this.clearSpill(uid);
    if (entry.dynamic && value instanceof ArrayValue) {
      this.writeValue(uid, normalize(value.get(0, 0)));
      return;
    }
    this.writeValue(uid, normalize(resultScalar(value, ctx)));
  }

  private spill(uid: number, entry: FormulaEntry, value: Value, ctx: FnContext): void {
    const arr = toArray(value, ctx);
    const target: Range = { r0: entry.r, c0: entry.c, r1: entry.r + arr.rows - 1, c1: entry.c + arr.cols - 1 };
    const sheet = entry.sheet;
    const anchorKey = cellKey(entry.r, entry.c);
    let blocked = target.r1 >= MAX_ROWS || target.c1 >= MAX_COLS;
    if (!blocked) {
      sheet.forEachCellIn(target, (r, c, cell) => {
        if (blocked || (r === entry.r && c === entry.c)) return;
        if (cell.spill === anchorKey) return;
        if (cell.v !== null || cell.f !== undefined) blocked = true;
      });
      if (!blocked && sheet.mergesIntersecting(target).length) blocked = true;
    }
    if (blocked) {
      this.clearSpill(uid);
      this.blockedSpills.set(uid, target);
      this.writeValue(uid, Errors.SPILL);
      return;
    }
    this.blockedSpills.delete(uid);
    const old = this.spills.get(uid);
    if (old) {
      for (let r = old.r0; r <= old.r1; r++) {
        for (let c = old.c0; c <= old.c1; c++) {
          if ((r === entry.r && c === entry.c) || rangeContains(target, r, c)) continue;
          this.clearSpillCell(sheet, r, c, anchorKey);
        }
      }
    }
    this.spills.set(uid, target);
    for (let r = 0; r < arr.rows; r++) {
      for (let c = 0; c < arr.cols; c++) {
        const v = normalize(arr.get(r, c));
        if (r === 0 && c === 0) this.writeValue(uid, v);
        else this.writeSpillCell(sheet, entry.r + r, entry.c + c, v, anchorKey);
      }
    }
    // Keep the array range on the anchor so it can be saved as a dynamic array formula.
    const anchor = sheet.getCell(entry.r, entry.c);
    const ref = `${rangeLabel(target)}`;
    if (anchor && anchor.arr !== ref) sheet.setCell(entry.r, entry.c, { ...anchor, arr: ref });
  }

  private distributeArray(uid: number, entry: FormulaEntry, value: Value, ctx: FnContext, range: Range): void {
    const sheet = entry.sheet;
    const anchorKey = cellKey(entry.r, entry.c);
    const isMulti = value instanceof ArrayValue || (value instanceof RefValue && !value.isCell());
    const arr = isMulti ? toArray(value, ctx) : null;
    const scalar = isMulti ? null : normalize(resultScalar(value, ctx));
    for (let r = range.r0; r <= range.r1; r++) {
      for (let c = range.c0; c <= range.c1; c++) {
        let v: Scalar;
        if (!arr) v = scalar;
        else {
          const rr = arr.rows === 1 ? 0 : r - range.r0;
          const cc = arr.cols === 1 ? 0 : c - range.c0;
          v = rr < arr.rows && cc < arr.cols ? normalize(arr.get(rr, cc)) : Errors.NA;
        }
        if (r === entry.r && c === entry.c) this.writeValue(uid, v);
        else this.writeSpillCell(sheet, r, c, v, anchorKey);
      }
    }
    this.spills.set(uid, range);
  }

  private writeSpillCell(sheet: Sheet, r: number, c: number, v: Scalar, anchorKey: number): void {
    const cell = sheet.getCell(r, c);
    if (cell && cell.v === v && cell.spill === anchorKey) return;
    sheet.setCell(r, c, { ...(cell ?? { v: null }), v, spill: anchorKey });
    this.changed.add(cellUid(sheet, r, c));
    this.markReaders(cellUid(sheet, r, c));
  }

  private clearSpillCell(sheet: Sheet, r: number, c: number, anchorKey: number): void {
    const cell = sheet.getCell(r, c);
    if (!cell || cell.spill !== anchorKey) return;
    const { spill: _spill, ...rest } = cell;
    void _spill;
    sheet.setCell(r, c, { ...rest, v: null });
    this.changed.add(cellUid(sheet, r, c));
    this.markReaders(cellUid(sheet, r, c));
  }

  /** Formulas reading a cell whose value changed during this recalculation must run (again). */
  private markReaders(uid: number): void {
    const deps: number[] = [];
    this.dependentsOf(uid, deps);
    for (const d of deps) if (!this.computing.has(d)) this.dirty.add(d);
  }

  clearSpill(uid: number): void {
    const range = this.spills.get(uid);
    if (!range) return;
    const sheet = this.workbook.sheetByUid(uidSheet(uid));
    this.spills.delete(uid);
    if (!sheet) return;
    const key = uidKey(uid);
    for (let r = range.r0; r <= range.r1; r++) {
      for (let c = range.c0; c <= range.c1; c++) {
        if (cellKey(r, c) === key) continue;
        this.clearSpillCell(sheet, r, c, key);
      }
    }
    const anchor = sheet.getCell(keyRow(key), keyCol(key));
    if (anchor?.arr && anchor.dyn) sheet.setCell(keyRow(key), keyCol(key), { ...anchor, arr: undefined });
  }

  private writeValue(uid: number, v: Scalar): void {
    const sheet = this.workbook.sheetByUid(uidSheet(uid));
    if (!sheet) return;
    const key = uidKey(uid);
    const r = keyRow(key);
    const c = keyCol(key);
    const cell = sheet.getCell(r, c);
    if (!cell) return;
    if (cell.v === v) return;
    sheet.setCell(r, c, { ...cell, v, rt: undefined });
    this.changed.add(uid);
  }

  spillRange(sheet: Sheet, r: number, c: number): Range | null {
    const uid = cellUid(sheet, r, c);
    if (this.dirty.has(uid)) this.ensure(uid);
    return this.spills.get(uid) ?? null;
  }

  /** Spill range of the formula anchored at a cell, if it spills. */
  spillOf(sheet: Sheet, r: number, c: number): Range | undefined {
    return this.spills.get(cellUid(sheet, r, c));
  }

  isFormula(sheet: Sheet, r: number, c: number): boolean {
    return this.formulas.has(cellUid(sheet, r, c));
  }

  // ---------------------------------------------------------------------------------------------
  // EngineApi

  private ensure(uid: number): void {
    if (!this.dirty.has(uid)) return;
    if (this.computing.has(uid)) {
      this.circular.add(uid);
      return;
    }
    this.compute(uid);
  }

  cellValue(sheet: Sheet, r: number, c: number): Scalar {
    const cell = sheet.getCell(r, c);
    if (!cell) return null;
    if (cell.f !== undefined) {
      const uid = cellUid(sheet, r, c);
      if (this.computing.has(uid)) {
        this.circular.add(uid);
        return 0;
      }
      this.ensure(uid);
      return sheet.getCell(r, c)?.v ?? null;
    }
    if (cell.spill !== undefined) {
      const anchorUid = sheet.uid * SHEET_FACTOR + cell.spill;
      if (this.dirty.has(anchorUid) && !this.computing.has(anchorUid)) this.ensure(anchorUid);
      return sheet.getCell(r, c)?.v ?? null;
    }
    return cell.v;
  }

  private clip(area: Area): Range | null {
    const { sheet, range } = area;
    const ext = sheet.dataExtent();
    const r1 = Math.min(range.r1, Math.max(ext.rows, range.r0));
    const c1 = Math.min(range.c1, Math.max(ext.cols, range.c0));
    if (r1 < range.r0 || c1 < range.c0) return null;
    return { r0: range.r0, c0: range.c0, r1, c1 };
  }

  forEachValue(area: Area, fn: (v: Scalar, r: number, c: number) => void): void {
    const { sheet, range } = area;
    sheet.forEachCellIn(range, (r, c) => {
      const v = this.cellValue(sheet, r, c);
      fn(v, r, c);
    });
  }

  areaToArray(area: Area): ArrayValue {
    const range = this.clip(area) ?? { ...area.range, r1: area.range.r0, c1: area.range.c0 };
    const rows = range.r1 - range.r0 + 1;
    const cols = range.c1 - range.c0 + 1;
    const data: Scalar[][] = Array.from({ length: rows }, () => new Array<Scalar>(cols).fill(null));
    area.sheet.forEachCellIn(range, (r, c) => {
      data[r - range.r0][c - range.c0] = this.cellValue(area.sheet, r, c);
    });
    return new ArrayValue(data);
  }

  sheetByName(name: string): Sheet | undefined {
    return this.workbook.sheetByName(name);
  }

  formulaText(sheet: Sheet, r: number, c: number): string | undefined {
    return sheet.getCell(r, c)?.f;
  }

  resolveReferenceText(text: string, ctx: FnContext): RefValue | CellError {
    const { sheet: sheetName, ref } = splitSheetRef(text);
    const sheet = sheetName !== undefined ? this.workbook.sheetByName(sheetName) : ctx.sheet;
    if (!sheet) return Errors.REF;
    const range = parseRange(ref);
    if (range) return new RefValue([{ sheet, range }]);
    const v = this.resolveName(ref, sheetName, { ...ctx, resolveName: (n, q, c) => this.resolveName(n, q, c), spillRange: (s, r, c) => this.spillRange(s, r, c) } as EvalContext);
    return v instanceof RefValue ? v : Errors.REF;
  }

  resolveName(name: string, qualifier: string | undefined, ctx: EvalContext): Value {
    const def = this.lookupName(name, qualifier, ctx.sheet);
    if (!def) return Errors.NAME;
    const ast = this.nameAst(def.formula);
    if (!ast) return Errors.NAME;
    return evaluate(ast, ctx);
  }

  /** Evaluates a formula text in the context of a cell without storing anything (used by the UI). */
  evaluateText(text: string, sheet: Sheet, r: number, c: number): Value {
    const ast = this.parse(text);
    if (!ast) return Errors.NAME;
    const entry: FormulaEntry = { sheet, r, c, text, ast, legacy: false, dynamic: true, volatile: false, cellDeps: [], rangeDeps: [] };
    return evaluate(ast, this.context(entry));
  }
}

function normalize(v: Scalar): Scalar {
  if (typeof v === 'number') {
    if (!Number.isFinite(v)) return Errors.NUM;
    if (Object.is(v, -0)) return 0;
  }
  return v;
}

function rangeLabel(r: Range): string {
  const name = (row: number, col: number) => {
    let s = '';
    let n = col + 1;
    while (n > 0) {
      const rem = (n - 1) % 26;
      s = String.fromCharCode(65 + rem) + s;
      n = Math.floor((n - 1) / 26);
    }
    return `${s}${row + 1}`;
  };
  return r.r0 === r.r1 && r.c0 === r.c1 ? name(r.r0, r.c0) : `${name(r.r0, r.c0)}:${name(r.r1, r.c1)}`;
}
