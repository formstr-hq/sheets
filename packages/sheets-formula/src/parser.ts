// Formula text -> syntax tree. Excel's precedence, from loosest to tightest:
//   = <> < > <= >=   then   &   then   + -   then   * /   then   ^   then   unary - +   then   %
// so -2^2 is 4, as in Excel, and 2^-1 works.
import { ERROR_CODES, ParseError, type BinaryOp, type CellRef, type ErrorCode, type Node, type RangeRef } from "./types";
import { columnIndex } from "./refs";

const SHEET = String.raw`((?:'(?:[^']|'')+'|[A-Za-z_][A-Za-z0-9_.]*)!)?`;
const CELL = String.raw`(\$?)([A-Za-z]{1,3})(\$?)(\d+)`;
const COL = String.raw`(\$?)([A-Za-z]{1,3})`;
const ROW = String.raw`(\$?)(\d+)`;
const END = String.raw`(?![A-Za-z0-9_(])`;

const RANGE_RE = new RegExp(`^${SHEET}${CELL}:${CELL}${END}`);
const COL_RANGE_RE = new RegExp(`^${SHEET}${COL}:${COL}${END}`);
const ROW_RANGE_RE = new RegExp(`^${SHEET}${ROW}:${ROW}${END}`);
const CELL_RE = new RegExp(`^${SHEET}${CELL}${END}`);
const NUMBER_RE = /^(?:\d+\.?\d*|\.\d+)(?:[eE][+-]?\d+)?/;
const NAME_RE = /^[A-Za-z_\\][A-Za-z0-9_.]*/;
const COMPARISONS = ["<>", "<=", ">=", "=", "<", ">"] as const;

function sheetOf(raw: string | undefined): string | undefined {
  if (!raw) return undefined;
  const body = raw.slice(0, -1);
  return body.startsWith("'") ? body.slice(1, -1).replace(/''/g, "'") : body;
}

function cellOf(sheet: string | undefined, colAbs: string, letters: string, rowAbs: string, digits: string): CellRef {
  return {
    sheet: sheetOf(sheet),
    colAbs: colAbs === "$",
    col: columnIndex(letters),
    rowAbs: rowAbs === "$",
    row: Number(digits) - 1,
  };
}

function rangeOf(a: CellRef, b: CellRef): RangeRef {
  const rowsFlipped = a.row > b.row;
  const colsFlipped = a.col > b.col;
  return {
    sheet: a.sheet,
    r1: Math.min(a.row, b.row),
    c1: Math.min(a.col, b.col),
    r2: Math.max(a.row, b.row),
    c2: Math.max(a.col, b.col),
    r1Abs: rowsFlipped ? b.rowAbs : a.rowAbs,
    c1Abs: colsFlipped ? b.colAbs : a.colAbs,
    r2Abs: rowsFlipped ? a.rowAbs : b.rowAbs,
    c2Abs: colsFlipped ? a.colAbs : b.colAbs,
  };
}

export function parse(source: string): Node {
  const text = source.startsWith("=") ? source.slice(1) : source;
  let pos = 0;

  const fail = (message: string): never => {
    throw new ParseError(message, pos);
  };
  const skip = () => {
    while (pos < text.length && /\s/.test(text[pos])) pos++;
  };
  const peek = (): string | undefined => {
    skip();
    return text[pos];
  };
  const eat = (token: string): boolean => {
    skip();
    if (text.startsWith(token, pos)) {
      pos += token.length;
      return true;
    }
    return false;
  };

  const comparison = (): Node => {
    let left = concat();
    for (;;) {
      skip();
      const op = COMPARISONS.find((candidate) => text.startsWith(candidate, pos));
      if (!op) return left;
      pos += op.length;
      left = { t: "bin", op: op as BinaryOp, l: left, r: concat() };
    }
  };
  const concat = (): Node => {
    let left = additive();
    while (peek() === "&") {
      pos++;
      left = { t: "bin", op: "&", l: left, r: additive() };
    }
    return left;
  };
  const additive = (): Node => {
    let left = multiplicative();
    for (;;) {
      const ch = peek();
      if (ch !== "+" && ch !== "-") return left;
      pos++;
      left = { t: "bin", op: ch, l: left, r: multiplicative() };
    }
  };
  const multiplicative = (): Node => {
    let left = power();
    for (;;) {
      const ch = peek();
      if (ch !== "*" && ch !== "/") return left;
      pos++;
      left = { t: "bin", op: ch, l: left, r: power() };
    }
  };
  const power = (): Node => {
    let left = unary();
    while (peek() === "^") {
      pos++;
      left = { t: "bin", op: "^", l: left, r: unary() };
    }
    return left;
  };
  const unary = (): Node => {
    const ch = peek();
    if (ch === "-" || ch === "+") {
      pos++;
      return { t: "unary", op: ch, arg: unary() };
    }
    let node = primary();
    while (peek() === "%") {
      pos++;
      node = { t: "pct", arg: node };
    }
    return node;
  };

  const primary = (): Node => {
    skip();
    const rest = text.slice(pos);
    const ch = text[pos];
    if (ch === undefined) return fail("Unexpected end of formula");

    if (ch === "(") {
      pos++;
      const inner = comparison();
      if (!eat(")")) fail("Missing )");
      return inner;
    }

    if (ch === '"') {
      let value = "";
      pos++;
      for (;;) {
        if (pos >= text.length) fail("Unterminated string");
        if (text[pos] === '"') {
          if (text[pos + 1] === '"') {
            value += '"';
            pos += 2;
            continue;
          }
          pos++;
          return { t: "str", v: value };
        }
        value += text[pos++];
      }
    }

    if (ch === "{") {
      pos++;
      const rows: Node[][] = [[]];
      for (;;) {
        rows[rows.length - 1].push(comparison());
        if (eat(",")) continue;
        if (eat(";")) {
          rows.push([]);
          continue;
        }
        if (eat("}")) break;
        fail("Bad array literal");
      }
      return { t: "array", rows };
    }

    if (ch === "#") {
      const upper = rest.toUpperCase();
      const code = ERROR_CODES.find((candidate) => upper.startsWith(candidate));
      if (!code) return fail("Unknown error value");
      pos += code.length;
      return { t: "err", v: code as ErrorCode };
    }

    // "1:1" is a whole-row reference, not the number 1.
    const number = NUMBER_RE.exec(rest);
    if (number && !ROW_RANGE_RE.test(rest)) {
      pos += number[0].length;
      return { t: "num", v: Number(number[0]) };
    }

    // NAME( ... )
    const name = NAME_RE.exec(rest);
    if (name && text[pos + name[0].length] === "(") {
      pos += name[0].length + 1;
      const args: Node[] = [];
      if (peek() === ")") {
        pos++;
      } else {
        for (;;) {
          // An empty argument, as in SUM(1,,2), is an omitted value.
          args.push(peek() === "," || peek() === ")" ? { t: "missing" } : comparison());
          if (eat(",")) continue;
          if (eat(")")) break;
          fail("Expected , or )");
        }
      }
      return { t: "call", name: name[0].toUpperCase(), args };
    }

    let m = RANGE_RE.exec(rest);
    if (m) {
      pos += m[0].length;
      const a = cellOf(m[1], m[2], m[3], m[4], m[5]);
      const b = cellOf(undefined, m[6], m[7], m[8], m[9]);
      return { t: "range", ref: rangeOf(a, b) };
    }
    m = COL_RANGE_RE.exec(rest);
    if (m) {
      pos += m[0].length;
      const first = columnIndex(m[3]);
      const second = columnIndex(m[5]);
      return {
        t: "range",
        ref: {
          sheet: sheetOf(m[1]),
          r1: 0,
          r2: null,
          c1: Math.min(first, second),
          c2: Math.max(first, second),
          r1Abs: true,
          r2Abs: true,
          c1Abs: (first <= second ? m[2] : m[4]) === "$",
          c2Abs: (first <= second ? m[4] : m[2]) === "$",
        },
      };
    }
    m = ROW_RANGE_RE.exec(rest);
    if (m) {
      pos += m[0].length;
      const first = Number(m[3]) - 1;
      const second = Number(m[5]) - 1;
      return {
        t: "range",
        ref: {
          sheet: sheetOf(m[1]),
          r1: Math.min(first, second),
          r2: Math.max(first, second),
          c1: 0,
          c2: null,
          r1Abs: (first <= second ? m[2] : m[4]) === "$",
          r2Abs: (first <= second ? m[4] : m[2]) === "$",
          c1Abs: true,
          c2Abs: true,
        },
      };
    }
    m = CELL_RE.exec(rest);
    if (m) {
      pos += m[0].length;
      return { t: "ref", ref: cellOf(m[1], m[2], m[3], m[4], m[5]) };
    }

    if (name) {
      pos += name[0].length;
      const upper = name[0].toUpperCase();
      if (upper === "TRUE") return { t: "bool", v: true };
      if (upper === "FALSE") return { t: "bool", v: false };
      return { t: "name", name: name[0] };
    }
    return fail(`Unexpected "${ch}"`);
  };

  const result = comparison();
  skip();
  if (pos < text.length) fail(`Unexpected "${text[pos]}"`);
  return result;
}
