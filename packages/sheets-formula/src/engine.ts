// The recalculation engine: keeps formulas, their cached results and who depends on whom, and when a
// cell changes recomputes only what is affected, in dependency order. Constants (typed values) stay
// with the host and are read through `getConstant`, so a chunked sheet never has to be fully loaded.
//
// Order is found with an iterative topological sort (no recursion, so a 100,000-long chain of formulas
// is fine). Cells that depend on themselves, directly or through others, become #CIRC!.
import { evaluateCell, type Context } from "./evaluate";
import { parse } from "./parser";
import { print } from "./print";
import { moveOne, references, renameSheet as renameInTree, rewriteForEdit, type Edit } from "./rewrite";
import { FormulaError, ParseError, type Node, type Scalar } from "./types";

export type CellValue = Scalar | FormulaError;

export interface EngineOptions {
  /** A typed (non-formula) cell's value; undefined or null for an empty cell. */
  getConstant(sheet: string, row: number, col: number): CellValue | undefined;
  extent?(sheet: string): { rows: number; cols: number };
  sheetNames?: Context["sheetNames"];
  resolveName?: Context["resolveName"];
  now?: Context["now"];
  random?: Context["random"];
  maxRangeCells?: number;
}

export interface Change {
  sheet: string;
  row: number;
  col: number;
  value: CellValue;
}

interface Rect {
  sheet: string; // lower-cased
  r1: number;
  c1: number;
  r2: number; // Infinity for whole-column references
  c2: number;
}

interface FormulaCell {
  sheet: string; // as written
  row: number;
  col: number;
  node: Node | null; // null when the text did not parse
  value: CellValue;
  cells: { sheet: string; row: number; col: number }[]; // single-cell precedents (lower-cased sheet)
  rects: Rect[]; // range precedents
  volatile: boolean;
}

const MAX_COLS = 1 << 20;
const VOLATILE = new Set(["NOW", "TODAY", "RAND", "RANDBETWEEN", "OFFSET", "INDIRECT"]);

const lower = (sheet: string) => sheet.toLowerCase();
const posKey = (row: number, col: number) => row * MAX_COLS + col;

export function sameValue(a: CellValue | undefined, b: CellValue | undefined): boolean {
  if (a instanceof FormulaError || b instanceof FormulaError) {
    return a instanceof FormulaError && b instanceof FormulaError && a.code === b.code;
  }
  return (a ?? null) === (b ?? null);
}

function isVolatile(node: Node): boolean {
  switch (node.t) {
    case "call":
      return VOLATILE.has(node.name.toUpperCase()) || node.args.some(isVolatile);
    case "unary":
    case "pct":
      return isVolatile(node.arg);
    case "bin":
      return isVolatile(node.l) || isVolatile(node.r);
    case "array":
      return node.rows.some((row) => row.some(isVolatile));
    default:
      return false;
  }
}

export class Engine {
  private readonly bySheet = new Map<string, Map<number, FormulaCell>>();
  private readonly dirty = new Set<FormulaCell>();
  /** Single-cell subscriptions: "sheet|row|col" -> formulas that read exactly that cell. */
  private readonly cellSubs = new Map<string, Set<FormulaCell>>();
  /** Formulas with range precedents, by sheet; scanned when a cell changes. */
  private readonly rangeSubs = new Map<string, Set<FormulaCell>>();
  private readonly pending: Change[] = [];

  constructor(private readonly options: EngineOptions) {}

  get formulaCount(): number {
    let n = 0;
    for (const sheet of this.bySheet.values()) n += sheet.size;
    return n;
  }

  // ---- reading -------------------------------------------------------------------------------

  hasFormula(sheet: string, row: number, col: number): boolean {
    return this.at(sheet, row, col) !== undefined;
  }

  /** The formula text with its leading "=", or undefined when the cell has none (or it did not parse). */
  formulaText(sheet: string, row: number, col: number): string | undefined {
    const cell = this.at(sheet, row, col);
    return cell?.node ? print(cell.node) : undefined;
  }

  /** A formula cell's result (recalculating first if anything is out of date), or undefined if not a formula. */
  value(sheet: string, row: number, col: number): CellValue | undefined {
    const cell = this.at(sheet, row, col);
    if (!cell) return undefined;
    if (this.dirty.size > 0) this.recalculate();
    return cell.value;
  }

  /** Every formula cell, for saving or listing. */
  *formulas(): IterableIterator<{ sheet: string; row: number; col: number; text: string | undefined; value: CellValue }> {
    for (const map of this.bySheet.values()) {
      for (const cell of map.values()) {
        yield { sheet: cell.sheet, row: cell.row, col: cell.col, text: cell.node ? print(cell.node) : undefined, value: cell.value };
      }
    }
  }

  // ---- changing ------------------------------------------------------------------------------

  /** Put a formula in a cell ("=A1+1"). Text that does not parse is kept and shows #NAME?. */
  setFormula(sheet: string, row: number, col: number, text: string): void {
    let node: Node | null = null;
    try {
      node = parse(text);
    } catch (error) {
      if (!(error instanceof ParseError)) throw error;
    }
    this.install(sheet, row, col, node);
  }

  /** Put an already-parsed formula in a cell. */
  setNode(sheet: string, row: number, col: number, node: Node): void {
    this.install(sheet, row, col, node);
  }

  /** The cell is no longer a formula (the host now stores a typed value, or nothing, there). */
  clearFormula(sheet: string, row: number, col: number): void {
    const cell = this.at(sheet, row, col);
    if (!cell) return;
    this.remove(cell);
    this.markDependents(lower(sheet), row, col, row, col);
  }

  /** The host changed a typed value; everything that reads it will be recomputed. */
  notifyChanged(sheet: string, row: number, col: number): void {
    this.markDependents(lower(sheet), row, col, row, col);
  }

  /** The host changed a block of typed values (paste, clear, fill, import). */
  notifyRange(sheet: string, r1: number, c1: number, r2: number, c2: number): void {
    this.markDependents(lower(sheet), r1, c1, r2, c2);
  }

  /** Recompute every formula that uses NOW, TODAY, RAND or the like. */
  recalculateVolatile(): void {
    for (const map of this.bySheet.values()) for (const cell of map.values()) if (cell.volatile) this.dirty.add(cell);
  }

  /** Recompute everything from scratch. */
  recalculateAll(): void {
    for (const map of this.bySheet.values()) for (const cell of map.values()) this.dirty.add(cell);
  }

  /**
   * Rows or columns were inserted or deleted. Formulas are rewritten and move with their cells; formulas
   * in deleted rows disappear. The host moves its own typed values. Everything is recomputed afterwards.
   */
  applyEdit(edit: Edit): void {
    const all = [...this.bySheet.values()].flatMap((map) => [...map.values()]);
    const target = lower(edit.sheet);
    for (const cell of all) this.unsubscribe(cell);
    this.bySheet.clear();
    this.dirty.clear();
    for (const cell of all) {
      let { row, col } = cell;
      if (lower(cell.sheet) === target) {
        const moved = edit.axis === "row" ? moveOne(row, edit) : moveOne(col, edit);
        if (moved === null) continue; // the cell itself was deleted
        if (edit.axis === "row") row = moved;
        else col = moved;
      }
      const node = cell.node ? rewriteForEdit(cell.node, edit, cell.sheet) : null;
      this.place({ ...cell, row, col, node });
    }
    this.recalculateAll();
  }

  /** A sheet was renamed: references and formula locations follow. */
  renameSheet(from: string, to: string): void {
    const all = [...this.bySheet.values()].flatMap((map) => [...map.values()]);
    for (const cell of all) this.unsubscribe(cell);
    this.bySheet.clear();
    for (const cell of all) {
      const node = cell.node ? renameInTree(cell.node, from, to) : null;
      this.place({ ...cell, sheet: lower(cell.sheet) === lower(from) ? to : cell.sheet, node });
    }
    this.recalculateAll();
  }

  // ---- recalculation -------------------------------------------------------------------------

  /** Recompute everything that is out of date. Returns the cells whose result actually changed. */
  recalculate(): Change[] {
    if (this.dirty.size > 0) this.run();
    return this.pending.splice(0);
  }

  // ---- internals -----------------------------------------------------------------------------

  private at(sheet: string, row: number, col: number): FormulaCell | undefined {
    return this.bySheet.get(lower(sheet))?.get(posKey(row, col));
  }

  private install(sheet: string, row: number, col: number, node: Node | null): void {
    const existing = this.at(sheet, row, col);
    if (existing) this.remove(existing);
    this.place({ sheet, row, col, node, value: node ? null : new FormulaError("#NAME?"), cells: [], rects: [], volatile: false });
    // Whatever read this cell as a typed value now sees a formula result.
    this.markDependents(lower(sheet), row, col, row, col);
  }

  /** Index a formula cell and register its precedents. */
  private place(cell: FormulaCell): void {
    const sheetKey = lower(cell.sheet);
    let map = this.bySheet.get(sheetKey);
    if (!map) this.bySheet.set(sheetKey, (map = new Map()));
    map.set(posKey(cell.row, cell.col), cell);

    cell.cells = [];
    cell.rects = [];
    cell.volatile = false;
    if (cell.node) {
      cell.volatile = isVolatile(cell.node);
      for (const ref of references(cell.node)) {
        const sheetOf = lower(ref.sheet ?? cell.sheet);
        if (ref.kind === "cell" && ref.cell) {
          cell.cells.push({ sheet: sheetOf, row: ref.cell.row, col: ref.cell.col });
          const key = `${sheetOf}|${ref.cell.row}|${ref.cell.col}`;
          let subs = this.cellSubs.get(key);
          if (!subs) this.cellSubs.set(key, (subs = new Set()));
          subs.add(cell);
        } else if (ref.range) {
          const r = ref.range;
          cell.rects.push({ sheet: sheetOf, r1: r.r1, c1: r.c1, r2: r.r2 ?? Infinity, c2: r.c2 ?? Infinity });
        }
      }
      if (cell.rects.length > 0) {
        for (const rect of cell.rects) {
          let subs = this.rangeSubs.get(rect.sheet);
          if (!subs) this.rangeSubs.set(rect.sheet, (subs = new Set()));
          subs.add(cell);
        }
      }
    }
    this.dirty.add(cell);
  }

  private unsubscribe(cell: FormulaCell): void {
    for (const ref of cell.cells) {
      const key = `${ref.sheet}|${ref.row}|${ref.col}`;
      const subs = this.cellSubs.get(key);
      subs?.delete(cell);
      if (subs?.size === 0) this.cellSubs.delete(key);
    }
    for (const rect of cell.rects) this.rangeSubs.get(rect.sheet)?.delete(cell);
    cell.cells = [];
    cell.rects = [];
  }

  private remove(cell: FormulaCell): void {
    this.unsubscribe(cell);
    this.bySheet.get(lower(cell.sheet))?.delete(posKey(cell.row, cell.col));
    this.dirty.delete(cell);
  }

  /** Mark every formula that reads anything in the rectangle, and everything downstream of those, as dirty. */
  private markDependents(sheet: string, r1: number, c1: number, r2: number, c2: number): void {
    const queue: FormulaCell[] = [];
    const add = (cell: FormulaCell) => {
      if (!this.dirty.has(cell)) {
        this.dirty.add(cell);
        queue.push(cell);
      }
    };
    this.collect(sheet, r1, c1, r2, c2, add);
    // Each newly dirty formula also dirties whatever reads its own cell.
    for (let i = 0; i < queue.length; i++) {
      const cell = queue[i];
      this.collect(lower(cell.sheet), cell.row, cell.col, cell.row, cell.col, add);
    }
  }

  private collect(sheet: string, r1: number, c1: number, r2: number, c2: number, add: (cell: FormulaCell) => void): void {
    const area = (r2 - r1 + 1) * (c2 - c1 + 1);
    if (area <= 64) {
      for (let r = r1; r <= r2; r++) {
        for (let c = c1; c <= c2; c++) this.cellSubs.get(`${sheet}|${r}|${c}`)?.forEach(add);
      }
    } else {
      // Big block: walk the subscriptions instead of every cell in it.
      for (const [key, subs] of this.cellSubs) {
        const [s, r, c] = key.split("|");
        if (s === sheet && +r >= r1 && +r <= r2 && +c >= c1 && +c <= c2) subs.forEach(add);
      }
    }
    const ranged = this.rangeSubs.get(sheet);
    if (!ranged) return;
    for (const cell of ranged) {
      for (const rect of cell.rects) {
        if (rect.sheet === sheet && rect.r1 <= r2 && rect.r2 >= r1 && rect.c1 <= c2 && rect.c2 >= c1) {
          add(cell);
          break;
        }
      }
    }
  }

  private run(): void {
    const todo = [...this.dirty];
    this.dirty.clear();
    const perSheet = new Map<string, Map<number, FormulaCell>>();
    for (const cell of todo) {
      const key = lower(cell.sheet);
      let map = perSheet.get(key);
      if (!map) perSheet.set(key, (map = new Map()));
      map.set(posKey(cell.row, cell.col), cell);
    }

    // Edges among the cells being recomputed: precedent -> dependents.
    const waiting = new Map<FormulaCell, number>();
    const dependents = new Map<FormulaCell, FormulaCell[]>();
    for (const cell of todo) {
      const precedents = new Set<FormulaCell>();
      for (const ref of cell.cells) {
        const found = perSheet.get(ref.sheet)?.get(posKey(ref.row, ref.col));
        if (found) precedents.add(found);
      }
      for (const rect of cell.rects) {
        const sheetCells = perSheet.get(rect.sheet);
        if (!sheetCells) continue;
        const area = (Math.min(rect.r2, 1e9) - rect.r1 + 1) * (Math.min(rect.c2, MAX_COLS) - rect.c1 + 1);
        if (area <= sheetCells.size) {
          for (let r = rect.r1; r <= rect.r2 && r <= 1e9; r++) {
            for (let c = rect.c1; c <= rect.c2 && c < MAX_COLS; c++) {
              const found = sheetCells.get(posKey(r, c));
              if (found) precedents.add(found);
            }
          }
        } else {
          for (const found of sheetCells.values()) {
            if (found.row >= rect.r1 && found.row <= rect.r2 && found.col >= rect.c1 && found.col <= rect.c2) precedents.add(found);
          }
        }
      }
      waiting.set(cell, precedents.size);
      for (const p of precedents) {
        let list = dependents.get(p);
        if (!list) dependents.set(p, (list = []));
        list.push(cell);
      }
    }

    const ready: FormulaCell[] = todo.filter((cell) => waiting.get(cell) === 0);
    const done = new Set<FormulaCell>();
    const finish = (cell: FormulaCell) => {
      done.add(cell);
      for (const next of dependents.get(cell) ?? []) {
        const left = (waiting.get(next) ?? 0) - 1;
        waiting.set(next, left);
        if (left === 0) ready.push(next);
      }
    };

    while (done.size < todo.length) {
      while (ready.length > 0) {
        const cell = ready.pop() as FormulaCell;
        this.compute(cell);
        finish(cell);
      }
      if (done.size === todo.length) break;
      // Nothing is ready but cells remain: they are in a cycle or hang off one. Cells in a cycle are
      // found by repeatedly dropping remaining cells that nothing remaining depends on.
      const remaining = new Set(todo.filter((cell) => !done.has(cell)));
      const outgoing = new Map<FormulaCell, number>();
      for (const cell of remaining) {
        outgoing.set(cell, (dependents.get(cell) ?? []).filter((d) => remaining.has(d)).length);
      }
      const reverse = new Map<FormulaCell, FormulaCell[]>();
      for (const cell of remaining) {
        for (const next of dependents.get(cell) ?? []) {
          if (!remaining.has(next)) continue;
          let list = reverse.get(next);
          if (!list) reverse.set(next, (list = []));
          list.push(cell);
        }
      }
      const prune = [...remaining].filter((cell) => outgoing.get(cell) === 0);
      for (let i = 0; i < prune.length; i++) {
        remaining.delete(prune[i]);
        for (const p of reverse.get(prune[i]) ?? []) {
          const left = (outgoing.get(p) ?? 0) - 1;
          outgoing.set(p, left);
          if (left === 0) prune.push(p);
        }
      }
      // What is left sits on a cycle or between two cycles: all of it is circular.
      for (const cell of remaining) {
        this.store(cell, new FormulaError("#CIRC!"));
        finish(cell);
      }
    }
  }

  private context(cell: FormulaCell): Context {
    return {
      sheet: cell.sheet,
      row: cell.row,
      col: cell.col,
      getCell: (sheet, row, col) => {
        const formula = this.at(sheet, row, col);
        return formula ? formula.value : this.options.getConstant(sheet, row, col);
      },
      extent: this.options.extent,
      hasFormula: (sheet, row, col) => this.at(sheet, row, col) !== undefined,
      sheetNames: this.options.sheetNames,
      resolveName: this.options.resolveName,
      now: this.options.now,
      random: this.options.random,
      maxRangeCells: this.options.maxRangeCells,
    };
  }

  private compute(cell: FormulaCell): void {
    const value = cell.node ? evaluateCell(cell.node, this.context(cell)) : new FormulaError("#NAME?");
    this.store(cell, value);
  }

  private store(cell: FormulaCell, value: CellValue): void {
    const before = cell.value;
    cell.value = value;
    if (!sameValue(before, value)) this.pending.push({ sheet: cell.sheet, row: cell.row, col: cell.col, value });
  }
}
