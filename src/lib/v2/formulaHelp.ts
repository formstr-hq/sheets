// Help while typing a formula: which function the cursor is inside, its arguments with the current
// one marked, and a coloured view of the formula's parts. Text only, no DOM.

/** Signatures of the functions people type most; others still show their name. */
export const SIGNATURES: Record<string, string> = {
  SUM: "number1, [number2], …",
  AVERAGE: "number1, [number2], …",
  COUNT: "value1, [value2], …",
  COUNTA: "value1, [value2], …",
  MAX: "number1, [number2], …",
  MIN: "number1, [number2], …",
  IF: "condition, value_if_true, [value_if_false]",
  IFERROR: "value, value_if_error",
  IFS: "condition1, value1, [condition2, value2], …",
  AND: "logical1, [logical2], …",
  OR: "logical1, [logical2], …",
  NOT: "logical",
  SUMIF: "range, criteria, [sum_range]",
  SUMIFS: "sum_range, criteria_range1, criteria1, …",
  COUNTIF: "range, criteria",
  COUNTIFS: "criteria_range1, criteria1, …",
  AVERAGEIF: "range, criteria, [average_range]",
  VLOOKUP: "search_key, range, index, [is_sorted]",
  HLOOKUP: "search_key, range, index, [is_sorted]",
  XLOOKUP: "search_key, lookup_range, result_range, [if_not_found]",
  INDEX: "range, [row], [column]",
  MATCH: "search_key, range, [search_type]",
  LOOKUP: "search_key, search_range, [result_range]",
  CONCATENATE: "text1, [text2], …",
  CONCAT: "text1, [text2], …",
  TEXTJOIN: "delimiter, ignore_empty, text1, [text2], …",
  LEFT: "text, [number_of_characters]",
  RIGHT: "text, [number_of_characters]",
  MID: "text, start, length",
  LEN: "text",
  TRIM: "text",
  UPPER: "text",
  LOWER: "text",
  PROPER: "text",
  SUBSTITUTE: "text, search_for, replace_with, [occurrence]",
  REPLACE: "text, position, length, new_text",
  FIND: "search_for, text, [start]",
  SEARCH: "search_for, text, [start]",
  TEXT: "number, format",
  VALUE: "text",
  ROUND: "number, [places]",
  ROUNDUP: "number, [places]",
  ROUNDDOWN: "number, [places]",
  INT: "number",
  ABS: "number",
  MOD: "number, divisor",
  POWER: "base, exponent",
  SQRT: "number",
  PRODUCT: "number1, [number2], …",
  RAND: "",
  RANDBETWEEN: "low, high",
  TODAY: "",
  NOW: "",
  DATE: "year, month, day",
  YEAR: "date",
  MONTH: "date",
  DAY: "date",
  WEEKDAY: "date, [type]",
  DATEDIF: "start_date, end_date, unit",
  EDATE: "start_date, months",
  EOMONTH: "start_date, months",
  NETWORKDAYS: "start_date, end_date, [holidays]",
  PMT: "rate, number_of_periods, present_value, [future_value], [end_or_beginning]",
  FV: "rate, number_of_periods, payment, [present_value], [end_or_beginning]",
  PV: "rate, number_of_periods, payment, [future_value], [end_or_beginning]",
  NPV: "discount, cashflow1, [cashflow2], …",
  IRR: "cashflow_amounts, [rate_guess]",
  ISBLANK: "value",
  ISNUMBER: "value",
  ISTEXT: "value",
  ISERROR: "value",
  MEDIAN: "number1, [number2], …",
  STDEV: "value1, [value2], …",
  LARGE: "data, n",
  SMALL: "data, n",
  RANK: "value, data, [is_ascending]",
  CHOOSE: "index, choice1, [choice2], …",
  SWITCH: "expression, case1, value1, [default or case2, value2], …",
  UNIQUE: "range",
  SORT: "range, sort_column, is_ascending",
  FILTER: "range, condition1, [condition2], …",
  TRANSPOSE: "range",
};

export interface FormulaHelp {
  name: string;
  /** The arguments as written in the signature, e.g. ["range", "criteria", "[sum_range]"]. */
  args: string[];
  /** Which argument the cursor is in (0-based), clamped to the last when the function takes a list. */
  current: number;
}

/** The function call the end of `text` is inside (the innermost open bracket after a name), or null. */
export function helpFor(text: string): FormulaHelp | null {
  if (!text.startsWith("=")) return null;
  const stack: { name: string; arg: number }[] = [];
  let quoted = false;
  let word = "";
  for (let i = 1; i < text.length; i++) {
    const ch = text[i];
    if (ch === '"') {
      quoted = !quoted;
      word = "";
      continue;
    }
    if (quoted) continue;
    if (/[A-Za-z0-9_.]/.test(ch)) {
      word += ch;
      continue;
    }
    if (ch === "(") stack.push({ name: word.toUpperCase(), arg: 0 });
    else if (ch === ")") stack.pop();
    else if ((ch === "," || ch === ";") && stack.length > 0) stack[stack.length - 1].arg++;
    word = "";
  }
  // The innermost call that has a name (a bare bracket like =(1+2) has none).
  for (let i = stack.length - 1; i >= 0; i--) {
    const { name, arg } = stack[i];
    if (!name) continue;
    const signature = SIGNATURES[name];
    const args = signature === undefined ? ["…"] : signature === "" ? [] : signature.split(/,\s*/);
    // A trailing "…" means the last argument repeats.
    const repeats = args.length > 0 && args[args.length - 1] === "…";
    const max = repeats ? Math.max(0, args.length - 2) : args.length - 1;
    const current = repeats ? Math.min(arg, args.length - 1) : Math.min(arg, Math.max(0, max));
    return { name, args, current: repeats && arg > max ? Math.max(0, args.length - 2) : current };
  }
  return null;
}

export type TokenKind = "ref" | "function" | "string" | "number" | "operator" | "text";
export interface Token {
  text: string;
  kind: TokenKind;
  /** For references: a colour number (0..N) that repeats for the same reference. */
  color?: number;
}

const TOKEN =
  /("(?:[^"]|"")*")|((?:'[^']+'|[A-Za-z_][A-Za-z0-9_.]*)!)?(\$?[A-Za-z]{1,3}\$?\d+(?::\$?[A-Za-z]{1,3}\$?\d+)?|\$?[A-Za-z]{1,3}:\$?[A-Za-z]{1,3})(?![A-Za-z0-9_(])|([A-Za-z][A-Za-z0-9_.]*)(?=\()|(\d+(?:\.\d+)?(?:[eE][-+]?\d+)?%?)|([-+*/^&=<>:,;()%])/g;

/** Splits a formula into parts to colour: references (each reference gets its own colour), functions, text, numbers. */
export function tokenizeFormula(formula: string): Token[] {
  const tokens: Token[] = [];
  const colors = new Map<string, number>();
  let at = 0;
  for (const match of formula.matchAll(TOKEN)) {
    const index = match.index ?? 0;
    if (index > at) tokens.push({ text: formula.slice(at, index), kind: "text" });
    const text = match[0];
    if (match[1] !== undefined) tokens.push({ text, kind: "string" });
    else if (match[3] !== undefined) {
      const key = text.replace(/\$/g, "").toUpperCase();
      if (!colors.has(key)) colors.set(key, colors.size);
      tokens.push({ text, kind: "ref", color: colors.get(key) });
    } else if (match[4] !== undefined) tokens.push({ text, kind: "function" });
    else if (match[5] !== undefined) tokens.push({ text, kind: "number" });
    else tokens.push({ text, kind: "operator" });
    at = index + text.length;
  }
  if (at < formula.length) tokens.push({ text: formula.slice(at), kind: "text" });
  return tokens;
}
