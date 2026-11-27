// Functions we implement ourselves because formulajs gets them wrong or does not have them: ROUND
// with extreme digits, TEXT (number and date formats), date/time text parsing, regular expressions,
// XLOOKUP, and the small operator-named functions (ADD, GT, ...). Arguments arrive already evaluated.
import { compare, isArray, isError, numberToText, toNumber, toText } from "./coerce";
import { columnName } from "./refs";
import { FormulaError, type Scalar, type Value } from "./types";

type Args = (Value | undefined)[];
type Fn = (args: Args) => Value;

const EPOCH = Date.UTC(1899, 11, 30);
const DAY = 86_400_000;

const first = (value: Value | undefined): Scalar | FormulaError =>
  value === undefined ? null : isArray(value) ? (value[0]?.[0] === undefined ? null : first(value[0][0])) : value;

const num = (value: Value | undefined): number | FormulaError => toNumber(first(value));
const text = (value: Value | undefined): string | FormulaError => toText(first(value));
const VALUE = () => new FormulaError("#VALUE!");
const NUM = () => new FormulaError("#NUM!");

function flatten(value: Value | undefined, out: Scalar | FormulaError[] = []): (Scalar | FormulaError)[] {
  if (value === undefined) return out as (Scalar | FormulaError)[];
  if (isArray(value)) for (const row of value) for (const item of row) flatten(item, out as never);
  else (out as (Scalar | FormulaError)[]).push(value);
  return out as (Scalar | FormulaError)[];
}

// ---- rounding -----------------------------------------------------------------------------------

function scaled(x: number, digits: number): { value: number; back: (v: number) => number } | null {
  if (digits >= 0) {
    const factor = 10 ** Math.min(digits, 308);
    const value = Number((Math.abs(x) * factor).toPrecision(15));
    // Past 2^53 every double is already a whole number at this scale: nothing to round.
    return Number.isFinite(value) && value < 2 ** 53 ? { value, back: (v) => v / factor } : null;
  }
  const factor = 10 ** Math.min(-digits, 308);
  return { value: Number((Math.abs(x) / factor).toPrecision(15)), back: (v) => v * factor };
}

function rounder(mode: "half" | "up" | "down"): Fn {
  return (args) => {
    const x = num(args[0]);
    const d = num(args[1] ?? 0);
    if (isError(x)) return x;
    if (isError(d)) return d;
    const digits = Math.trunc(d);
    const s = scaled(x, digits);
    if (!s) return x;
    const whole = mode === "half" ? Math.floor(s.value + 0.5) : mode === "up" ? Math.ceil(s.value) : Math.floor(s.value);
    const out = s.back(whole);
    return Math.sign(x) * out === 0 ? 0 : Math.sign(x) * out;
  };
}

// ---- TEXT: number and date formats ------------------------------------------------------------

const MONTHS = ["January", "February", "March", "April", "May", "June", "July", "August", "September", "October", "November", "December"];
const WEEKDAYS = ["Sunday", "Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday"];

interface Piece {
  kind: "lit" | "tok";
  text: string;
}

/** Splits a format section into quoted/escaped literals and other characters. */
function pieces(section: string): Piece[] {
  const out: Piece[] = [];
  for (let i = 0; i < section.length; i++) {
    const ch = section[i];
    if (ch === '"') {
      const end = section.indexOf('"', i + 1);
      out.push({ kind: "lit", text: section.slice(i + 1, end < 0 ? undefined : end) });
      i = end < 0 ? section.length : end;
    } else if (ch === "\\" && i + 1 < section.length) {
      out.push({ kind: "lit", text: section[++i] });
    } else if (ch === "_" && i + 1 < section.length) {
      i++;
      out.push({ kind: "lit", text: " " });
    } else if (ch === "*" && i + 1 < section.length) {
      i++;
    } else if (ch === "[") {
      const end = section.indexOf("]", i);
      const body = end < 0 ? "" : section.slice(i + 1, end);
      // [h], [mm], [ss] are elapsed-time counters; colours and conditions are dropped.
      if (/^(h+|m+|s+)$/i.test(body)) out.push({ kind: "tok", text: `[${body.toLowerCase()}]` });
      // [$€-407] is a currency symbol (the part after the dash is the locale).
      else if (body.startsWith("$")) {
        const symbol = body.slice(1).split("-")[0];
        if (symbol) out.push({ kind: "lit", text: symbol });
      }
      i = end < 0 ? section.length : end;
    } else {
      out.push({ kind: "tok", text: ch });
    }
  }
  return out;
}

function splitSections(format: string): string[] {
  const sections: string[] = [];
  let current = "";
  let quoted = false;
  for (const ch of format) {
    if (ch === '"') quoted = !quoted;
    if (ch === ";" && !quoted) {
      sections.push(current);
      current = "";
    } else current += ch;
  }
  sections.push(current);
  return sections;
}

const isDateSection = (parts: Piece[]) => {
  const joined = parts.filter((p) => p.kind === "tok").map((p) => p.text).join("");
  return /[ymdhs]/i.test(joined.replace(/E[+-]/gi, "")) && !/[0#?]/.test(joined.replace(/\.0+/g, ""));
};

function formatDate(parts: Piece[], serial: number): string {
  const total = Math.round(serial * 86_400) * 1000; // whole seconds, like Excel
  const date = new Date(EPOCH + total);
  const y = date.getUTCFullYear();
  const mo = date.getUTCMonth();
  const d = date.getUTCDate();
  const wd = date.getUTCDay();
  const h = date.getUTCHours();
  const mi = date.getUTCMinutes();
  const s = date.getUTCSeconds();
  const twelve = parts.some((p) => /^(am\/pm|a\/p)$/i.test(p.text)) || /am\/pm|a\/p/i.test(parts.map((p) => p.text).join(""));
  const raw = parts.map((p) => (p.kind === "lit" ? `\u0000${p.text}\u0000` : p.text)).join("");
  let out = "";
  for (let i = 0; i < raw.length; ) {
    if (raw[i] === "\u0000") {
      const end = raw.indexOf("\u0000", i + 1);
      out += raw.slice(i + 1, end);
      i = end + 1;
      continue;
    }
    const rest = raw.slice(i);
    let m: RegExpMatchArray | null;
    if ((m = rest.match(/^(yyyy|yy)/i))) {
      out += m[1].length === 4 ? String(y).padStart(4, "0") : String(y % 100).padStart(2, "0");
    } else if ((m = rest.match(/^(mmmmm|mmmm|mmm|mm|m)/i))) {
      // m means minutes right after h or right before s; otherwise it is the month.
      const before = out.length > 0 && /h/i.test(raw.slice(0, i).replace(/\u0000[^\u0000]*\u0000/g, "").replace(/[^a-z]/gi, "").slice(-1));
      const after = /^m+(?:[^a-z0-9]|\u0000[^\u0000]*\u0000)*s/i.test(rest);
      if (before || after) out += m[1].length >= 2 ? String(mi).padStart(2, "0") : String(mi);
      else if (m[1].length === 5) out += MONTHS[mo][0];
      else if (m[1].length === 4) out += MONTHS[mo];
      else if (m[1].length === 3) out += MONTHS[mo].slice(0, 3);
      else out += m[1].length === 2 ? String(mo + 1).padStart(2, "0") : String(mo + 1);
    } else if ((m = rest.match(/^(dddd|ddd|dd|d)/i))) {
      out += m[1].length === 4 ? WEEKDAYS[wd] : m[1].length === 3 ? WEEKDAYS[wd].slice(0, 3) : m[1].length === 2 ? String(d).padStart(2, "0") : String(d);
    } else if ((m = rest.match(/^(hh|h)/i))) {
      const hour = twelve ? h % 12 || 12 : h;
      out += m[1].length === 2 ? String(hour).padStart(2, "0") : String(hour);
    } else if ((m = rest.match(/^(ss|s)/i))) {
      out += m[1].length === 2 ? String(s).padStart(2, "0") : String(s);
    } else if ((m = rest.match(/^am\/pm/i))) {
      out += h < 12 ? "AM" : "PM";
    } else if ((m = rest.match(/^a\/p/i))) {
      out += h < 12 ? "A" : "P";
    } else if ((m = rest.match(/^\[(h+)\]/i))) {
      out += String(Math.floor(serial * 24)).padStart(m[1].length, "0");
    } else if ((m = rest.match(/^\[(m+)\]/i))) {
      out += String(Math.floor(serial * 1440)).padStart(m[1].length, "0");
    } else if ((m = rest.match(/^\[(s+)\]/i))) {
      out += String(Math.floor(serial * 86400)).padStart(m[1].length, "0");
    } else {
      out += raw[i];
      m = null;
    }
    i += m ? m[0].length : 1;
  }
  return out;
}

/** Rounds half away from zero in decimal, not in binary (1.005 -> 1.01). */
function toFixedHalfUp(n: number, digits: number): string {
  const shifted = Math.round(Number(`${n}e${digits}`));
  if (!Number.isFinite(shifted)) return n.toFixed(Math.min(digits, 100));
  return Number(`${shifted}e-${digits}`).toFixed(digits);
}

function formatNumber(parts: Piece[], value: number, negativeSign: boolean): string {
  type Tok = { lit: string; ch?: never } | { ch: string; lit?: never };
  const tokens: Tok[] = parts.map((p): Tok => (p.kind === "lit" ? { lit: p.text } : { ch: p.text }));
  const isCh = (t: Tok | undefined, test: RegExp): boolean => t !== undefined && t.ch !== undefined && test.test(t.ch);
  const percents = tokens.filter((t) => "ch" in t && t.ch === "%").length;
  // Find the run of digit placeholders (0 # ? , .) that holds the number.
  let start = -1;
  let end = -1;
  tokens.forEach((t, i) => {
    if (isCh(t, /[0#?,.]/)) {
      if (start < 0) {
        start = i;
        end = i;
      } else if (end === i - 1) end = i;
    }
  });
  // Scientific notation: ...E+00 or ...E-00 right after the run.
  let sci: { plus: boolean; digits: number } | null = null;
  let runEnd = end;
  if (end >= 0 && end + 3 < tokens.length + 1) {
    const e = tokens[end + 1];
    const sign = tokens[end + 2];
    if (isCh(e, /^e$/i) && isCh(sign, /[+-]/)) {
      let k = end + 3;
      let digits = 0;
      while (isCh(tokens[k], /^0$/)) {
        digits++;
        k++;
      }
      if (digits > 0) {
        sci = { plus: sign?.ch === "+", digits };
        runEnd = k - 1;
      }
    }
  }
  const prefix = tokens.slice(0, Math.max(start, 0)).map((t) => ("lit" in t ? t.lit : t.ch === "%" ? "%" : t.ch)).join("");
  const run = start < 0 ? [] : tokens.slice(start, end + 1).map((t) => (t as { ch: string }).ch);
  const suffixTokens = start < 0 ? [] : tokens.slice(runEnd + 1);
  const suffix = suffixTokens.map((t) => ("lit" in t ? t.lit : t.ch)).join("");
  if (start < 0) return (negativeSign ? "-" : "") + tokens.map((t) => ("lit" in t ? t.lit : t.ch)).join("");

  const runText = run.join("");
  const dot = runText.indexOf(".");
  const intPattern = dot < 0 ? runText : runText.slice(0, dot);
  const fracPattern = dot < 0 ? "" : runText.slice(dot + 1);
  const grouping = /[0#?],[0#?]/.test(intPattern);
  const scale = intPattern.match(/,+$/)?.[0].length ?? 0;
  const cleanInt = intPattern.replace(/,/g, "");
  const decimals = fracPattern.replace(/,/g, "").length;
  const n = (Math.abs(value) * 100 ** percents) / 1000 ** scale;

  let body: string;
  let exponent = "";
  if (sci) {
    const exp = n === 0 ? 0 : Math.floor(Math.log10(n));
    const mantissa = n / 10 ** exp;
    body = toFixedHalfUp(mantissa, decimals);
    const e = Math.abs(exp).toString().padStart(sci.digits, "0");
    exponent = `E${exp < 0 ? "-" : sci.plus ? "+" : ""}${e}`;
  } else {
    body = toFixedHalfUp(n, decimals);
  }
  let [intPart, fracPart = ""] = body.split(".");
  const minInt = (cleanInt.match(/0/g) ?? []).length;
  if (intPart === "0" && minInt === 0) intPart = "";
  intPart = intPart.padStart(minInt, "0");
  if (grouping) intPart = intPart.replace(/\B(?=(\d{3})+(?!\d))/g, ",");
  // '#' and '?' after the point drop trailing zeros; '0' keeps them.
  const keep = fracPattern.replace(/,/g, "").lastIndexOf("0") + 1;
  while (fracPart.length > keep && fracPart.endsWith("0")) fracPart = fracPart.slice(0, -1);
  const frac = fracPart.length > 0 ? `.${fracPart}` : decimals > 0 && fracPattern.includes("0") ? `.${"0".repeat(decimals)}` : "";
  const sign = negativeSign && (Number(body) !== 0 || sci) ? "-" : "";
  return `${sign}${prefix}${intPart}${frac}${exponent}${suffix}`;
}

/** Excel's TEXT(value, format) for numbers, dates, times and text. */
export function formatValue(value: Scalar, format: string): string {
  const sections = splitSections(format);
  if (typeof value === "string") {
    const section = sections[3] ?? sections[0];
    const parts = pieces(section);
    if (!parts.some((p) => p.kind === "tok" && p.text === "@")) return sections.length >= 4 || /@/.test(format) ? parts.map((p) => p.text).join("") : value;
    return parts.map((p) => (p.kind === "tok" && p.text === "@" ? value : p.text)).join("");
  }
  const n = typeof value === "boolean" ? (value ? 1 : 0) : (value ?? 0);
  let index = 0;
  if (sections.length > 1 && n < 0) index = 1;
  if (sections.length > 2 && n === 0) index = 2;
  const parts = pieces(sections[index] ?? sections[0]);
  if (parts.length === 0) return "";
  if (parts.length === 1 && parts[0].kind === "tok" && /^general$/i.test(parts[0].text)) return numberToText(n);
  const joined = parts.map((p) => p.text).join("");
  if (/^general$/i.test(joined)) return numberToText(n);
  if (isDateSection(parts)) return formatDate(parts, Math.abs(n) === n || index === 1 ? Math.abs(n) : n);
  // A negative number in the first of several sections uses the second with its own sign handling.
  const negativeSign = n < 0 && index === 0 && sections.length === 1;
  return formatNumber(parts, n, negativeSign);
}

// ---- date and time text -----------------------------------------------------------------------

const MONTH_INDEX: Record<string, number> = Object.fromEntries(MONTHS.flatMap((name, i) => [[name.toLowerCase(), i], [name.slice(0, 3).toLowerCase(), i]]));

function parseDate(source: string): number | null {
  const s = source.trim();
  let m = s.match(/^(\d{4})[-/](\d{1,2})[-/](\d{1,2})(?:[T\s].*)?$/);
  if (m) return Date.UTC(+m[1], +m[2] - 1, +m[3]) / DAY - EPOCH / DAY;
  m = s.match(/^(\d{1,2})[-/](\d{1,2})[-/](\d{2,4})(?:\s.*)?$/);
  if (m) {
    const year = +m[3] < 30 && m[3].length <= 2 ? 2000 + +m[3] : +m[3] < 100 ? 1900 + +m[3] : +m[3];
    return Date.UTC(year, +m[1] - 1, +m[2]) / DAY - EPOCH / DAY;
  }
  m = s.match(/^(\d{1,2})[-\s]([A-Za-z]{3,9})[-,\s]+(\d{2,4})$/);
  if (m && MONTH_INDEX[m[2].toLowerCase()] !== undefined) return Date.UTC(+m[3], MONTH_INDEX[m[2].toLowerCase()], +m[1]) / DAY - EPOCH / DAY;
  m = s.match(/^([A-Za-z]{3,9})\.?\s+(\d{1,2}),?\s+(\d{4})$/);
  if (m && MONTH_INDEX[m[1].toLowerCase()] !== undefined) return Date.UTC(+m[3], MONTH_INDEX[m[1].toLowerCase()], +m[2]) / DAY - EPOCH / DAY;
  return null;
}

function parseTime(source: string): number | null {
  const m = source.match(/(\d{1,2}):(\d{2})(?::(\d{2}))?\s*(AM|PM)?/i);
  if (!m) return null;
  let hours = +m[1];
  if (m[4]) {
    if (hours < 1 || hours > 12) return null;
    hours = (hours % 12) + (m[4].toUpperCase() === "PM" ? 12 : 0);
  }
  if (hours > 23 || +m[2] > 59 || +(m[3] ?? 0) > 59) return null;
  return (hours * 3600 + +m[2] * 60 + +(m[3] ?? 0)) / 86_400;
}

// ---- lookup ------------------------------------------------------------------------------------

const wildcard = (pattern: string) =>
  new RegExp(`^${pattern.replace(/[.+^${}()|[\]\\]/g, "\\$&").replace(/~?\*/g, (m) => (m === "~*" ? "\\*" : ".*")).replace(/~?\?/g, (m) => (m === "~?" ? "\\?" : "."))}$`, "i");

function xmatch(needle: Scalar | FormulaError, haystack: (Scalar | FormulaError)[], mode: number, reverse: boolean): number {
  if (isError(needle)) return -1;
  const order = haystack.map((_, i) => i);
  if (reverse) order.reverse();
  let best = -1;
  for (const i of order) {
    const candidate = haystack[i];
    if (isError(candidate)) continue;
    if (mode === 2 && typeof needle === "string") {
      if (typeof candidate === "string" && wildcard(needle).test(candidate)) return i;
      continue;
    }
    const same = (typeof candidate === typeof needle || candidate === null) && compare(candidate, needle) === 0;
    if (same) return i;
    if (mode === -1 && typeof candidate === typeof needle && compare(candidate, needle) < 0) {
      if (best < 0 || compare(candidate, haystack[best] as Scalar) > 0) best = i;
    }
    if (mode === 1 && typeof candidate === typeof needle && compare(candidate, needle) > 0) {
      if (best < 0 || compare(candidate, haystack[best] as Scalar) < 0) best = i;
    }
  }
  return mode === 0 || mode === 2 ? -1 : best;
}

function lookupVector(value: Value | undefined): (Scalar | FormulaError)[] {
  return flatten(value);
}

// ---- the table ---------------------------------------------------------------------------------

const numeric2 = (fn: (a: number, b: number) => Value): Fn => (args) => {
  const a = num(args[0]);
  const b = num(args[1]);
  if (isError(a)) return a;
  if (isError(b)) return b;
  return fn(a, b);
};

const compare2 = (test: (c: number) => boolean): Fn => (args) => {
  const a = first(args[0]);
  const b = first(args[1]);
  if (isError(a)) return a;
  if (isError(b)) return b;
  return test(compare(a, b));
};

/** A square block of numbers, or an error. */
function squareMatrix(input: Value | undefined): number[][] | FormulaError {
  if (input === undefined || !isArray(input) || input.length === 0) return new FormulaError("#VALUE!");
  const value = input;
  const size = value.length;
  const out: number[][] = [];
  for (const row of value) {
    if (row.length !== size) return new FormulaError("#VALUE!");
    const line: number[] = [];
    for (const item of row) {
      if (isError(item)) return item;
      if (typeof item !== "number") return new FormulaError("#VALUE!");
      line.push(item);
    }
    out.push(line);
  }
  return out;
}

/** Gauss-Jordan on [A | I]: the determinant and (when it is not 0) the inverse. */
function eliminate(matrix: number[][]): { det: number; inverse: number[][] | null } {
  const n = matrix.length;
  const a = matrix.map((row, i) => [...row, ...Array.from({ length: n }, (_, j) => (i === j ? 1 : 0))]);
  let det = 1;
  for (let col = 0; col < n; col++) {
    let pivot = col;
    for (let r = col + 1; r < n; r++) if (Math.abs(a[r][col]) > Math.abs(a[pivot][col])) pivot = r;
    if (Math.abs(a[pivot][col]) < 1e-12) return { det: 0, inverse: null };
    if (pivot !== col) {
      [a[pivot], a[col]] = [a[col], a[pivot]];
      det = -det;
    }
    det *= a[col][col];
    const scale = a[col][col];
    for (let c = 0; c < 2 * n; c++) a[col][c] /= scale;
    for (let r = 0; r < n; r++) {
      if (r === col) continue;
      const factor = a[r][col];
      if (factor !== 0) for (let c = 0; c < 2 * n; c++) a[r][c] -= factor * a[col][c];
    }
  }
  return { det, inverse: a.map((row) => row.slice(n)) };
}

export const EXTRA: Record<string, Fn> = {
  /** ASC / DBCS convert between half- and full-width characters; a single-byte locale leaves text as it is. */
  ASC(args) {
    return text(args[0]);
  },
  DBCS(args) {
    return text(args[0]);
  },
  /** INFO(type): facts about the environment. */
  INFO(args) {
    const kind = text(args[0]);
    if (isError(kind)) return kind;
    const known: Record<string, string> = { recalc: "Automatic", system: "pcdos", osversion: "Web", release: "1.0", directory: "", origin: "$A:$A$1", numfile: "1" };
    return known[kind.toLowerCase()] ?? VALUE();
  },
  /** MDETERM(array): determinant of a square matrix. */
  MDETERM(args) {
    const matrix = squareMatrix(args[0]);
    if (isError(matrix)) return matrix;
    return Math.round(eliminate(matrix).det * 1e10) / 1e10;
  },
  /** MINVERSE(array): inverse of a square matrix (#NUM! when it has none). */
  MINVERSE(args) {
    const matrix = squareMatrix(args[0]);
    if (isError(matrix)) return matrix;
    const { inverse } = eliminate(matrix);
    return inverse ? inverse.map((row) => row.map((v) => Math.round(v * 1e12) / 1e12)) : new FormulaError("#NUM!");
  },

  /** ADDRESS(row, column, [absolute 1-4], [A1 style], [sheet]): the address of a cell as text. */
  ADDRESS(args) {
    const row = num(args[0]);
    const col = num(args[1]);
    if (isError(row)) return row;
    if (isError(col)) return col;
    if (row < 1 || col < 1) return VALUE();
    const kind = args[2] === undefined ? 1 : num(args[2]);
    if (isError(kind)) return kind;
    if (kind < 1 || kind > 4) return VALUE();
    const a1 = args[3] === undefined ? true : first(args[3]) !== false && first(args[3]) !== 0;
    const rowAbs = kind === 1 || kind === 2;
    const colAbs = kind === 1 || kind === 3;
    const r = Math.trunc(row);
    const c = Math.trunc(col);
    const cell = a1
      ? `${colAbs ? "$" : ""}${columnName(c - 1)}${rowAbs ? "$" : ""}${r}`
      : `R${rowAbs ? r : `[${r}]`}C${colAbs ? c : `[${c}]`}`;
    if (args[4] === undefined) return cell;
    const sheet = text(args[4]);
    if (isError(sheet)) return sheet;
    return `${/^[A-Za-z_][A-Za-z0-9_.]*$/.test(sheet) ? sheet : `'${sheet.replace(/'/g, "''")}'`}!${cell}`;
  },

  /** HYPERLINK(url, [label]): shows the label (or the address); the cell's link is set separately. */
  HYPERLINK(args) {
    const url = text(args[0]);
    if (isError(url)) return url;
    if (args[1] === undefined) return url;
    const label = first(args[1]);
    return label === null ? "" : label;
  },

  /** LENB / FINDB count characters (a single-byte locale has no double-byte characters). */
  LENB(args) {
    const value = text(args[0]);
    return isError(value) ? value : value.length;
  },
  FINDB(args) {
    const needle = text(args[0]);
    const hay = text(args[1]);
    const start = args[2] === undefined ? 1 : num(args[2]);
    if (isError(needle)) return needle;
    if (isError(hay)) return hay;
    if (isError(start)) return start;
    if (start < 1 || start > hay.length + 1) return VALUE();
    const at = hay.indexOf(needle, Math.trunc(start) - 1);
    return at < 0 ? VALUE() : at + 1;
  },

  // Google Sheets conversions: the value stays, only how it is shown differs.
  TO_TEXT(args) {
    const value = first(args[0]);
    return isError(value) ? value : toText(value);
  },
  TO_PURE_NUMBER(args) {
    return num(args[0]);
  },
  TO_PERCENT(args) {
    return num(args[0]);
  },
  TO_DOLLARS(args) {
    return num(args[0]);
  },
  TO_DATE(args) {
    const value = num(args[0]);
    return isError(value) ? value : Math.floor(value);
  },
  UNARY_PERCENT(args) {
    const value = num(args[0]);
    return isError(value) ? value : value / 100;
  },

  ROUND: rounder("half"),
  ROUNDUP: rounder("up"),
  ROUNDDOWN: rounder("down"),

  TEXT(args) {
    const format = text(args[1]);
    if (isError(format)) return format;
    const raw = first(args[0]);
    if (isError(raw)) return raw;
    let value: Scalar = raw;
    if (typeof raw === "string") {
      const asNumber = toNumber(raw);
      if (!isError(asNumber) && raw.trim() !== "") value = asNumber;
    }
    return formatValue(value, format);
  },

  // Text to number the way a cell would read it; text that is no number, date or time is #VALUE!.
  VALUE(args) {
    const raw = first(args[0]);
    if (isError(raw)) return raw;
    if (typeof raw === "number") return raw;
    if (typeof raw === "boolean" || raw === null) return VALUE();
    const asNumber = toNumber(raw);
    if (!isError(asNumber)) return asNumber;
    const time = parseTime(raw);
    if (time !== null) return time;
    const date = parseDate(raw);
    return date === null ? VALUE() : date;
  },

  DATEVALUE(args) {
    const source = text(args[0]);
    if (isError(source)) return source;
    const serial = parseDate(source);
    return serial === null ? VALUE() : serial;
  },
  TIMEVALUE(args) {
    const source = text(args[0]);
    if (isError(source)) return source;
    const t = parseTime(source);
    if (t !== null) return t;
    return parseDate(source) !== null ? 0 : VALUE();
  },

  REGEXMATCH(args) {
    const subject = text(args[0]);
    const pattern = text(args[1]);
    if (isError(subject)) return subject;
    if (isError(pattern)) return pattern;
    try {
      return new RegExp(pattern).test(subject);
    } catch {
      return VALUE();
    }
  },
  REGEXEXTRACT(args) {
    const subject = text(args[0]);
    const pattern = text(args[1]);
    if (isError(subject)) return subject;
    if (isError(pattern)) return pattern;
    try {
      const match = new RegExp(pattern).exec(subject);
      return match ? (match[1] ?? match[0]) : new FormulaError("#N/A");
    } catch {
      return VALUE();
    }
  },
  REGEXREPLACE(args) {
    const subject = text(args[0]);
    const pattern = text(args[1]);
    const replacement = text(args[2]);
    if (isError(subject)) return subject;
    if (isError(pattern)) return pattern;
    if (isError(replacement)) return replacement;
    try {
      return subject.replace(new RegExp(pattern, "g"), replacement);
    } catch {
      return VALUE();
    }
  },
  COUNTUNIQUE(args) {
    const seen = new Set<string>();
    for (const arg of args) {
      for (const item of flatten(arg)) {
        if (item === null || item === "") continue;
        seen.add(isError(item) ? item.code : `${typeof item}:${typeof item === "string" ? item.toLowerCase() : item}`);
      }
    }
    return seen.size;
  },

  XMATCH(args) {
    const mode = args[2] === undefined ? 0 : num(args[2]);
    const search = args[3] === undefined ? 1 : num(args[3]);
    if (isError(mode)) return mode;
    if (isError(search)) return search;
    const at = xmatch(first(args[0]), lookupVector(args[1]), mode, search < 0);
    return at < 0 ? new FormulaError("#N/A") : at + 1;
  },
  XLOOKUP(args) {
    const mode = args[4] === undefined ? 0 : num(args[4]);
    const search = args[5] === undefined ? 1 : num(args[5]);
    if (isError(mode)) return mode;
    if (isError(search)) return search;
    const at = xmatch(first(args[0]), lookupVector(args[1]), mode, search < 0);
    if (at < 0) return args[3] !== undefined ? args[3] : new FormulaError("#N/A");
    const result = args[2];
    if (!result || !isArray(result)) return result ?? null;
    // A vertical lookup returns the matching row, a horizontal one the matching column.
    const lookup = args[1] ?? null;
    const vertical = !isArray(lookup) || lookup.length >= (lookup[0]?.length ?? 0);
    return vertical ? [result[at] ?? []] : result.map((row) => [row[at]]);
  },

  ISBINARY(args) {
    const value = text(args[0]);
    return isError(value) ? value : /^[01]{1,10}$/.test(value);
  },
  SPLIT(args) {
    const source = text(args[0]);
    const delimiter = text(args[1]);
    if (isError(source)) return source;
    if (isError(delimiter)) return delimiter;
    const byEach = args[2] === undefined ? true : first(args[2]) !== false;
    const removeEmpty = args[3] === undefined ? true : first(args[3]) !== false;
    if (delimiter === "") return [[source]];
    const pattern = byEach ? new RegExp(`[${[...delimiter].map((c) => c.replace(/[\]^-]/g, "\  ISOCEILING: (args) => {")).join("")}]`) : null;
    let parts = pattern ? source.split(pattern) : source.split(delimiter);
    if (removeEmpty) parts = parts.filter((part) => part !== "");
    return [parts];
  },
  JOIN(args) {
    const delimiter = text(args[0]);
    if (isError(delimiter)) return delimiter;
    const items: string[] = [];
    for (const arg of args.slice(1)) {
      for (const item of flatten(arg)) {
        if (isError(item)) return item;
        if (item !== null) items.push(toText(item) as string);
      }
    }
    return items.join(delimiter);
  },
  NUMBERS(args) {
    const out: Scalar[] = [];
    for (const arg of args) for (const item of flatten(arg)) if (typeof item === "number") out.push(item);
    return [out];
  },
  FLATTEN(args) {
    const out: Value[][] = [];
    for (const arg of args) for (const item of flatten(arg)) out.push([item]);
    return out.length > 0 ? out : [[null]];
  },
  ARGS2ARRAY: (args) => [args.map((arg) => first(arg))],
  COUNTIN(args) {
    const needle = first(args[1]);
    if (isError(needle)) return needle;
    return flatten(args[0]).filter((item) => !isError(item) && compare(item, needle) === 0 && typeof item === typeof needle).length;
  },

  ISOCEILING: (args) => {
    const x = num(args[0]);
    const sig = args[1] === undefined ? 1 : num(args[1]);
    if (isError(x)) return x;
    if (isError(sig)) return sig;
    if (sig === 0) return 0;
    return Math.ceil(x / Math.abs(sig)) * Math.abs(sig);
  },

  ADD: numeric2((a, b) => a + b),
  MINUS: numeric2((a, b) => a - b),
  MULTIPLY: numeric2((a, b) => a * b),
  DIVIDE: numeric2((a, b) => (b === 0 ? new FormulaError("#DIV/0!") : a / b)),
  POW: numeric2((a, b) => (Number.isFinite(a ** b) ? a ** b : NUM())),
  EQ: compare2((c) => c === 0),
  NE: compare2((c) => c !== 0),
  GT: compare2((c) => c > 0),
  GTE: compare2((c) => c >= 0),
  LT: compare2((c) => c < 0),
  LTE: compare2((c) => c <= 0),
};

/** Results to upper-case: formulajs returns hex and base digits in lower case, Excel in upper case. */
export const UPPERCASE_RESULTS = new Set(["BASE", "DEC2HEX", "BIN2HEX", "OCT2HEX"]);
