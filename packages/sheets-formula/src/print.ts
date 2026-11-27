// Syntax tree -> formula text (with the leading "="). parse(print(x)) gives back the same tree.
import { columnName, quoteSheet } from "./refs";
import type { CellRef, Node, RangeRef } from "./types";

const PRECEDENCE: Record<string, number> = {
  "=": 1, "<>": 1, "<": 1, ">": 1, "<=": 1, ">=": 1,
  "&": 2,
  "+": 3, "-": 3,
  "*": 4, "/": 4,
  "^": 5,
};

const prefix = (sheet: string | undefined) => (sheet === undefined ? "" : `${quoteSheet(sheet)}!`);

export function printCell(ref: CellRef): string {
  return `${prefix(ref.sheet)}${ref.colAbs ? "$" : ""}${columnName(ref.col)}${ref.rowAbs ? "$" : ""}${ref.row + 1}`;
}

export function printRange(ref: RangeRef): string {
  const sheet = prefix(ref.sheet);
  if (ref.r2 === null) {
    return `${sheet}${ref.c1Abs ? "$" : ""}${columnName(ref.c1)}:${ref.c2Abs ? "$" : ""}${columnName(ref.c2 as number)}`;
  }
  if (ref.c2 === null) {
    return `${sheet}${ref.r1Abs ? "$" : ""}${ref.r1 + 1}:${ref.r2Abs ? "$" : ""}${ref.r2 + 1}`;
  }
  const a = `${ref.c1Abs ? "$" : ""}${columnName(ref.c1)}${ref.r1Abs ? "$" : ""}${ref.r1 + 1}`;
  const b = `${ref.c2Abs ? "$" : ""}${columnName(ref.c2)}${ref.r2Abs ? "$" : ""}${ref.r2 + 1}`;
  return `${sheet}${a}:${b}`;
}

function inner(node: Node, parent = 0, right = false): string {
  switch (node.t) {
    case "num":
      return String(node.v);
    case "str":
      return `"${node.v.replace(/"/g, '""')}"`;
    case "bool":
      return node.v ? "TRUE" : "FALSE";
    case "err":
      return node.v;
    case "missing":
      return "";
    case "ref":
      return printCell(node.ref);
    case "range":
      return printRange(node.ref);
    case "name":
      return node.name;
    case "call":
      return `${node.name}(${node.args.map((arg) => inner(arg)).join(",")})`;
    case "unary":
      return `${node.op}${inner(node.arg, 6)}`;
    case "pct":
      return `${inner(node.arg, 7)}%`;
    case "array":
      return `{${node.rows.map((row) => row.map((item) => inner(item)).join(",")).join(";")}}`;
    case "bin": {
      const level = PRECEDENCE[node.op];
      const text = `${inner(node.l, level)}${node.op}${inner(node.r, level, true)}`;
      // Parentheses are needed when a looser operation sits inside a tighter one, or on the right of an equal one.
      return level < parent || (right && level === parent) ? `(${text})` : text;
    }
  }
}

export function print(node: Node): string {
  return `=${inner(node)}`;
}
