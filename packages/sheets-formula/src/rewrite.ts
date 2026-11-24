// Walking and rewriting syntax trees: what a formula points at, how it changes when rows or columns
// are inserted or deleted, when it is copied elsewhere, and when a sheet is renamed.
import type { CellRef, Node, RangeRef } from "./types";

export type Edit = {
  kind: "insert" | "delete";
  axis: "row" | "col";
  /** The sheet whose rows or columns change. */
  sheet: string;
  /** First affected row or column (0-based) and how many. */
  index: number;
  count: number;
};

export interface Reference {
  kind: "cell" | "range";
  /** The sheet it names, or undefined for "the formula's own sheet". */
  sheet?: string;
  cell?: CellRef;
  range?: RangeRef;
}

const same = (a: string | undefined, b: string) => (a ?? "").toLowerCase() === b.toLowerCase();

/** Maps a node's children, leaving the structure otherwise intact. */
function mapChildren(node: Node, fn: (child: Node) => Node): Node {
  switch (node.t) {
    case "call":
      return { ...node, args: node.args.map(fn) };
    case "unary":
    case "pct":
      return { ...node, arg: fn(node.arg) };
    case "bin":
      return { ...node, l: fn(node.l), r: fn(node.r) };
    case "array":
      return { ...node, rows: node.rows.map((row) => row.map(fn)) };
    default:
      return node;
  }
}

/** Every cell and range a formula names, in order of appearance (the edges of the dependency graph). */
export function references(node: Node): Reference[] {
  const found: Reference[] = [];
  const visit = (n: Node): Node => {
    if (n.t === "ref") found.push({ kind: "cell", sheet: n.ref.sheet, cell: n.ref });
    else if (n.t === "range") found.push({ kind: "range", sheet: n.ref.sheet, range: n.ref });
    else mapChildren(n, visit);
    return n;
  };
  visit(node);
  return found;
}

/** Where a coordinate lands after an edit, or null if it was deleted. */
export function moveOne(value: number, edit: Edit): number | null {
  if (edit.kind === "insert") return value >= edit.index ? value + edit.count : value;
  if (value < edit.index) return value;
  if (value >= edit.index + edit.count) return value - edit.count;
  return null;
}

/** A range's [start, end] after an edit; null when the whole range was deleted. */
function moveSpan(start: number, end: number, edit: Edit): [number, number] | null {
  if (edit.kind === "insert") {
    return [start >= edit.index ? start + edit.count : start, end >= edit.index ? end + edit.count : end];
  }
  const last = edit.index + edit.count - 1;
  const newStart = start < edit.index ? start : start > last ? start - edit.count : edit.index;
  const newEnd = end < edit.index ? end : end > last ? end - edit.count : edit.index - 1;
  return newEnd < newStart ? null : [newStart, newEnd];
}

const REF_ERROR: Node = { t: "err", v: "#REF!" };

/**
 * The formula after rows or columns were inserted or deleted. `formulaSheet` is the sheet the formula
 * lives on: a reference without a sheet name means that sheet. Deleted targets become #REF!.
 */
export function rewriteForEdit(node: Node, edit: Edit, formulaSheet: string): Node {
  const affects = (sheet: string | undefined) => same(sheet ?? formulaSheet, edit.sheet);
  const visit = (n: Node): Node => {
    if (n.t === "ref") {
      if (!affects(n.ref.sheet)) return n;
      if (edit.axis === "row") {
        const row = moveOne(n.ref.row, edit);
        return row === null ? REF_ERROR : { t: "ref", ref: { ...n.ref, row } };
      }
      const col = moveOne(n.ref.col, edit);
      return col === null ? REF_ERROR : { t: "ref", ref: { ...n.ref, col } };
    }
    if (n.t === "range") {
      if (!affects(n.ref.sheet)) return n;
      const ref = n.ref;
      if (edit.axis === "row") {
        if (ref.r2 === null) return n; // whole columns are unaffected by row edits
        const span = moveSpan(ref.r1, ref.r2, edit);
        return span ? { t: "range", ref: { ...ref, r1: span[0], r2: span[1] } } : REF_ERROR;
      }
      if (ref.c2 === null) return n;
      const span = moveSpan(ref.c1, ref.c2, edit);
      return span ? { t: "range", ref: { ...ref, c1: span[0], c2: span[1] } } : REF_ERROR;
    }
    return mapChildren(n, visit);
  };
  return visit(node);
}

/**
 * The formula as copied or filled by (rowDelta, colDelta): relative parts move, $absolute parts stay.
 * A reference pushed off the sheet becomes #REF!.
 */
export function shiftReferences(node: Node, rowDelta: number, colDelta: number): Node {
  const visit = (n: Node): Node => {
    if (n.t === "ref") {
      const row = n.ref.rowAbs ? n.ref.row : n.ref.row + rowDelta;
      const col = n.ref.colAbs ? n.ref.col : n.ref.col + colDelta;
      return row < 0 || col < 0 ? REF_ERROR : { t: "ref", ref: { ...n.ref, row, col } };
    }
    if (n.t === "range") {
      const r = n.ref;
      const r1 = r.r1Abs ? r.r1 : r.r1 + rowDelta;
      const r2 = r.r2 === null ? null : r.r2Abs ? r.r2 : r.r2 + rowDelta;
      const c1 = r.c1Abs ? r.c1 : r.c1 + colDelta;
      const c2 = r.c2 === null ? null : r.c2Abs ? r.c2 : r.c2 + colDelta;
      if (r1 < 0 || c1 < 0 || (r2 !== null && r2 < 0) || (c2 !== null && c2 < 0)) return REF_ERROR;
      return { t: "range", ref: { ...r, r1, r2, c1, c2 } };
    }
    return mapChildren(n, visit);
  };
  return visit(node);
}

/** The formula after a sheet was renamed: references that name `from` now name `to`. */
export function renameSheet(node: Node, from: string, to: string): Node {
  const visit = (n: Node): Node => {
    if (n.t === "ref" && n.ref.sheet !== undefined && same(n.ref.sheet, from)) {
      return { t: "ref", ref: { ...n.ref, sheet: to } };
    }
    if (n.t === "range" && n.ref.sheet !== undefined && same(n.ref.sheet, from)) {
      return { t: "range", ref: { ...n.ref, sheet: to } };
    }
    return mapChildren(n, visit);
  };
  return visit(node);
}
