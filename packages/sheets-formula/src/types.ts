// Values and syntax tree of a formula. Rows and columns are 0-based everywhere.

export type ErrorCode = "#NULL!" | "#DIV/0!" | "#VALUE!" | "#REF!" | "#NAME?" | "#NUM!" | "#N/A" | "#CIRC!";

/** An Excel error value. Errors are values: they flow through formulas and IFERROR can catch them. */
export class FormulaError {
  constructor(readonly code: ErrorCode) {}
  toString(): string {
    return this.code;
  }
}

export const ERROR_CODES: readonly ErrorCode[] = ["#NULL!", "#DIV/0!", "#VALUE!", "#REF!", "#NAME?", "#NUM!", "#N/A", "#CIRC!"];

export type Scalar = number | string | boolean | null;
/** A cell value, or a 2D block of them (a range or an array literal). */
export type Value = Scalar | FormulaError | Value[][];

export interface CellRef {
  sheet?: string;
  row: number;
  col: number;
  rowAbs: boolean;
  colAbs: boolean;
}

/**
 * A rectangle. `r2: null` runs to the last row (a whole-column reference like A:A), `c2: null` to
 * the last column (a whole-row reference like 1:1).
 */
export interface RangeRef {
  sheet?: string;
  r1: number;
  c1: number;
  r2: number | null;
  c2: number | null;
  r1Abs: boolean;
  c1Abs: boolean;
  r2Abs: boolean;
  c2Abs: boolean;
}

export type BinaryOp = "+" | "-" | "*" | "/" | "^" | "&" | "=" | "<>" | "<" | ">" | "<=" | ">=";

export type Node =
  | { t: "num"; v: number }
  | { t: "str"; v: string }
  | { t: "bool"; v: boolean }
  | { t: "err"; v: ErrorCode }
  | { t: "missing" }
  | { t: "ref"; ref: CellRef }
  | { t: "range"; ref: RangeRef }
  | { t: "name"; name: string }
  | { t: "call"; name: string; args: Node[] }
  | { t: "unary"; op: "-" | "+"; arg: Node }
  | { t: "pct"; arg: Node }
  | { t: "bin"; op: BinaryOp; l: Node; r: Node }
  | { t: "array"; rows: Node[][] };

export class ParseError extends Error {
  constructor(
    message: string,
    readonly position: number,
  ) {
    super(message);
    this.name = "ParseError";
  }
}
