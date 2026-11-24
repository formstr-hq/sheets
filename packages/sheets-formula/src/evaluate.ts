// Evaluating a syntax tree. Cells are read through callbacks, so the engine never needs the whole
// sheet in memory: a formula only asks for the cells it names, and a chunked store can answer from
// whichever chunks are loaded.
import { compare, isArray, isError, toBoolean, toNumber, toText, VALUE_ERROR } from "./coerce";
import { callFunction, LAZY_FUNCTIONS, type LazyContext } from "./functions";
import { parse } from "./parser";
import { FormulaError, ParseError, type BinaryOp, type Node, type RangeRef, type Scalar, type Value } from "./types";

export interface Context {
  /** The sheet the formula lives on (what a reference without a sheet name means). */
  sheet: string;
  /** The formula's own cell, for ROW() and COLUMN(). */
  row?: number;
  col?: number;
  /** A cell's value: undefined or null for an empty cell. */
  getCell(sheet: string, row: number, col: number): Scalar | FormulaError | undefined;
  /** How many rows and columns hold data, to size whole-column (A:A) and whole-row (1:1) references. */
  extent?(sheet: string): { rows: number; cols: number };
  /** Whether a cell holds a formula (for ISFORMULA). */
  hasFormula?(sheet: string, row: number, col: number): boolean;
  /** Sheet names in tab order (for SHEET and SHEETS). */
  sheetNames?(): string[];
  /** A defined name. */
  resolveName?(name: string): Value | undefined;
  /** Clock and randomness, injectable so results are testable. */
  now?(): Date;
  random?(): number;
  /** Largest block a range may expand to (default 5,000,000 cells) before the formula errors. */
  maxRangeCells?: number;
}

const DEFAULT_MAX_RANGE_CELLS = 5_000_000;

/** The top-left value of an array, or the value itself (Excel's implicit intersection). */
export function scalarOf(value: Value): Scalar | FormulaError {
  return isArray(value) ? (value[0]?.[0] === undefined ? null : scalarOf(value[0][0])) : value;
}

function readRange(ref: RangeRef, ctx: Context): Value {
  const sheet = ref.sheet ?? ctx.sheet;
  const extent = ctx.extent?.(sheet) ?? { rows: 0, cols: 0 };
  const r2 = ref.r2 ?? extent.rows - 1;
  const c2 = ref.c2 ?? extent.cols - 1;
  const rows = Math.max(0, r2 - ref.r1 + 1);
  const cols = Math.max(0, c2 - ref.c1 + 1);
  if (rows * cols > (ctx.maxRangeCells ?? DEFAULT_MAX_RANGE_CELLS)) return new FormulaError("#VALUE!");
  const block: Value[][] = [];
  for (let r = 0; r < rows; r++) {
    const line: Value[] = [];
    for (let c = 0; c < cols; c++) line.push(ctx.getCell(sheet, ref.r1 + r, ref.c1 + c) ?? null);
    block.push(line);
  }
  return block;
}

/** Applies fn to every pair of values, broadcasting a scalar or a one-row/column array like Excel. */
function lift(l: Value, r: Value, fn: (a: Scalar | FormulaError, b: Scalar | FormulaError) => Scalar | FormulaError): Value {
  if (!isArray(l) && !isArray(r)) return fn(l, r);
  const a = isArray(l) ? l : [[l]];
  const b = isArray(r) ? r : [[r]];
  const rows = Math.max(a.length, b.length);
  const cols = Math.max(a[0]?.length ?? 0, b[0]?.length ?? 0);
  const out: Value[][] = [];
  for (let i = 0; i < rows; i++) {
    const line: Value[] = [];
    for (let j = 0; j < cols; j++) {
      const x = a.length === 1 ? a[0][j] : a[i]?.[j];
      const y = b.length === 1 ? b[0][j] : b[i]?.[j];
      const xv = a[0].length === 1 && a.length > 1 ? a[i]?.[0] : x;
      const yv = b[0].length === 1 && b.length > 1 ? b[i]?.[0] : y;
      line.push(xv === undefined || yv === undefined ? new FormulaError("#N/A") : fn(scalarOf(xv), scalarOf(yv)));
    }
    out.push(line);
  }
  return out;
}

function binary(op: BinaryOp, a: Scalar | FormulaError, b: Scalar | FormulaError): Scalar | FormulaError {
  if (isError(a)) return a;
  if (isError(b)) return b;
  switch (op) {
    case "&": {
      return `${toText(a)}${toText(b)}`;
    }
    case "=":
      return compare(a, b) === 0;
    case "<>":
      return compare(a, b) !== 0;
    case "<":
      return compare(a, b) < 0;
    case ">":
      return compare(a, b) > 0;
    case "<=":
      return compare(a, b) <= 0;
    case ">=":
      return compare(a, b) >= 0;
  }
  const x = toNumber(a);
  if (isError(x)) return x;
  const y = toNumber(b);
  if (isError(y)) return y;
  let result: number;
  switch (op) {
    case "+":
      result = x + y;
      break;
    case "-":
      result = x - y;
      break;
    case "*":
      result = x * y;
      break;
    case "/":
      if (y === 0) return new FormulaError("#DIV/0!");
      result = x / y;
      break;
    case "^":
      if (x === 0 && y === 0) return new FormulaError("#NUM!");
      result = x ** y;
      break;
  }
  return Number.isFinite(result) ? result : new FormulaError("#NUM!");
}

export function evaluate(node: Node, ctx: Context): Value {
  switch (node.t) {
    case "num":
    case "str":
    case "bool":
      return node.v;
    case "err":
      return new FormulaError(node.v);
    case "missing":
      return null;
    case "ref":
      return ctx.getCell(node.ref.sheet ?? ctx.sheet, node.ref.row, node.ref.col) ?? null;
    case "range":
      return readRange(node.ref, ctx);
    case "name": {
      const value = ctx.resolveName?.(node.name);
      return value === undefined ? new FormulaError("#NAME?") : value;
    }
    case "array":
      return node.rows.map((row) => row.map((item) => scalarOf(evaluate(item, ctx))));
    case "unary": {
      const value = evaluate(node.arg, ctx);
      if (node.op === "+") return value;
      return lift(value, 0, (a) => {
        const n = toNumber(a);
        return isError(n) ? n : -n;
      });
    }
    case "pct":
      return lift(evaluate(node.arg, ctx), 0, (a) => {
        const n = toNumber(a);
        return isError(n) ? n : n / 100;
      });
    case "bin":
      return lift(evaluate(node.l, ctx), evaluate(node.r, ctx), (a, b) => binary(node.op, a, b));
    case "call": {
      const lazy = LAZY_FUNCTIONS[node.name];
      const helper: LazyContext = { ctx, evaluate: (n) => evaluate(n, ctx) };
      if (lazy) return lazy(node.args, helper);
      const args = node.args.map((arg) => (arg.t === "missing" ? undefined : evaluate(arg, ctx)));
      for (const arg of args) if (isError(arg)) return arg;
      return callFunction(node.name, args, ctx);
    }
  }
}

/** The value a cell shows: an array result collapses to its top-left value. */
export function evaluateCell(node: Node, ctx: Context): Scalar | FormulaError {
  return scalarOf(evaluate(node, ctx));
}

/** Parse and evaluate formula text. Text that does not parse gives #NAME?. */
export function calculate(formula: string, ctx: Context): Scalar | FormulaError {
  try {
    return evaluateCell(parse(formula), ctx);
  } catch (error) {
    if (error instanceof ParseError) return new FormulaError("#NAME?");
    throw error;
  }
}

export { VALUE_ERROR, toBoolean };
