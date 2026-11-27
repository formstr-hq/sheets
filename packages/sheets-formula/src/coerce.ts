// Excel's rules for turning values into numbers, text and booleans, and for comparing them.
import { FormulaError, type Scalar, type Value } from "./types";

export const isError = (value: unknown): value is FormulaError => value instanceof FormulaError;
export const isArray = (value: Value): value is Value[][] => Array.isArray(value);

export const VALUE_ERROR = () => new FormulaError("#VALUE!");

/** Number(0.1+0.2) prints as 0.30000000000000004; Excel shows 15 significant digits. */
export function numberToText(n: number): string {
  if (!Number.isFinite(n)) return "#NUM!";
  return String(Number(n.toPrecision(15)));
}

export function toNumber(value: Scalar | FormulaError): number | FormulaError {
  if (isError(value)) return value;
  if (typeof value === "number") return value;
  if (typeof value === "boolean") return value ? 1 : 0;
  if (value === null) return 0;
  const text = value.trim();
  if (text === "") return VALUE_ERROR();
  const percent = text.endsWith("%");
  const n = Number((percent ? text.slice(0, -1) : text).replace(/,/g, ""));
  return Number.isNaN(n) ? VALUE_ERROR() : percent ? n / 100 : n;
}

export function toText(value: Scalar | FormulaError): string | FormulaError {
  if (isError(value)) return value;
  if (typeof value === "string") return value;
  if (typeof value === "number") return numberToText(value);
  if (typeof value === "boolean") return value ? "TRUE" : "FALSE";
  return "";
}

export function toBoolean(value: Scalar | FormulaError): boolean | FormulaError {
  if (isError(value)) return value;
  if (typeof value === "boolean") return value;
  if (typeof value === "number") return value !== 0;
  if (value === null) return false;
  const upper = value.trim().toUpperCase();
  if (upper === "TRUE") return true;
  if (upper === "FALSE") return false;
  return VALUE_ERROR();
}

const rank = (value: number | string | boolean) => (typeof value === "number" ? 1 : typeof value === "string" ? 2 : 3);

/** Negative, zero or positive, comparing like Excel: numbers < text < booleans, text without regard to case. */
export function compare(a: Scalar, b: Scalar): number {
  // An empty cell equals 0, "" or FALSE, whichever the other side is.
  const fill = (empty: Scalar, other: Scalar) =>
    empty !== null ? empty : typeof other === "string" ? "" : typeof other === "boolean" ? false : 0;
  const x = fill(a, b) as number | string | boolean;
  const y = fill(b, a) as number | string | boolean;
  const rx = rank(x);
  const ry = rank(y);
  if (rx !== ry) return rx - ry;
  if (typeof x === "string") {
    const lx = x.toLowerCase();
    const ly = (y as string).toLowerCase();
    return lx < ly ? -1 : lx > ly ? 1 : 0;
  }
  return Number(x) - Number(y);
}
