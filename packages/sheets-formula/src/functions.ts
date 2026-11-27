// The function table. Most functions come from @formulajs/formulajs (MIT) through a thin wrapper that
// converts values both ways; the ones whose behaviour must be exact, lazy or timezone-independent are
// ours: IF and friends (arguments that must not be evaluated unless chosen), the IS* tests, AND/OR,
// DATE/TODAY/NOW (formulajs returns JS Dates in local time; we return serial numbers).
import * as formulajs from "@formulajs/formulajs";
import { isArray, isError, toBoolean, toNumber } from "./coerce";
import type { Context } from "./evaluate";
import { BOND_FUNCTIONS } from "./bonds";
import { EXTRA, UPPERCASE_RESULTS } from "./extras";
import { ERROR_CODES, FormulaError, type Node, type Scalar, type Value } from "./types";

export interface LazyContext {
  ctx: Context;
  evaluate(node: Node): Value;
}
type LazyFunction = (args: Node[], helper: LazyContext) => Value;

const EPOCH = Date.UTC(1899, 11, 30);
const DAY = 86_400_000;

/** A JS Date (as formulajs returns, in local time) to a spreadsheet serial number. */
export function dateToSerial(date: Date): number {
  const wall = Date.UTC(date.getFullYear(), date.getMonth(), date.getDate(), date.getHours(), date.getMinutes(), date.getSeconds(), date.getMilliseconds());
  return (wall - EPOCH) / DAY;
}

function scalar(value: Value): Scalar | FormulaError {
  return isArray(value) ? (value[0]?.[0] === undefined ? null : scalar(value[0][0])) : value;
}

const bool = (value: Value): boolean | FormulaError => toBoolean(scalar(value));

function ifFunction(args: Node[], { evaluate }: LazyContext): Value {
  const condition = bool(evaluate(args[0] ?? { t: "missing" }));
  if (isError(condition)) return condition;
  const branch = condition ? args[1] : args[2];
  if (!branch) return condition ? 0 : false;
  return branch.t === "missing" ? 0 : evaluate(branch);
}

export const LAZY_FUNCTIONS: Record<string, LazyFunction> = {
  IF: ifFunction,
  IFERROR(args, { evaluate }) {
    const value = evaluate(args[0] ?? { t: "missing" });
    const first = scalar(value);
    return isError(first) ? evaluate(args[1] ?? { t: "missing" }) : value;
  },
  IFNA(args, { evaluate }) {
    const value = evaluate(args[0] ?? { t: "missing" });
    const first = scalar(value);
    return isError(first) && first.code === "#N/A" ? evaluate(args[1] ?? { t: "missing" }) : value;
  },
  IFS(args, { evaluate }) {
    for (let i = 0; i + 1 < args.length; i += 2) {
      const condition = bool(evaluate(args[i]));
      if (isError(condition)) return condition;
      if (condition) return evaluate(args[i + 1]);
    }
    return new FormulaError("#N/A");
  },
  CHOOSE(args, { evaluate }) {
    const index = toNumber(scalar(evaluate(args[0] ?? { t: "missing" })));
    if (isError(index)) return index;
    const pick = args[Math.trunc(index)];
    return Math.trunc(index) < 1 || !pick ? new FormulaError("#VALUE!") : evaluate(pick);
  },
  SWITCH(args, { evaluate }) {
    const subject = scalar(evaluate(args[0] ?? { t: "missing" }));
    if (isError(subject)) return subject;
    let i = 1;
    for (; i + 1 < args.length; i += 2) {
      const candidate = scalar(evaluate(args[i]));
      if (!isError(candidate) && String(candidate).toLowerCase() === String(subject).toLowerCase()) return evaluate(args[i + 1]);
    }
    return i < args.length ? evaluate(args[i]) : new FormulaError("#N/A");
  },
  AND: (args, helper) => allOrAny(args, helper, true),
  OR: (args, helper) => allOrAny(args, helper, false),
  NOT(args, { evaluate }) {
    const value = bool(evaluate(args[0] ?? { t: "missing" }));
    return isError(value) ? value : !value;
  },
  ISERROR: (args, { evaluate }) => isError(scalar(evaluate(args[0] ?? { t: "missing" }))),
  ISERR(args, { evaluate }) {
    const value = scalar(evaluate(args[0] ?? { t: "missing" }));
    return isError(value) && value.code !== "#N/A";
  },
  ISNA(args, { evaluate }) {
    const value = scalar(evaluate(args[0] ?? { t: "missing" }));
    return isError(value) && value.code === "#N/A";
  },
  ISNUMBER: (args, { evaluate }) => typeof scalar(evaluate(args[0] ?? { t: "missing" })) === "number",
  ISTEXT: (args, { evaluate }) => typeof scalar(evaluate(args[0] ?? { t: "missing" })) === "string",
  ISNONTEXT: (args, { evaluate }) => typeof scalar(evaluate(args[0] ?? { t: "missing" })) !== "string",
  ISLOGICAL: (args, { evaluate }) => typeof scalar(evaluate(args[0] ?? { t: "missing" })) === "boolean",
  ISBLANK: (args, { evaluate }) => scalar(evaluate(args[0] ?? { t: "missing" })) === null,
  ISREF: (args) => args[0]?.t === "ref" || args[0]?.t === "range",
  ISFORMULA(args, { ctx }) {
    const target = args[0];
    if (target?.t !== "ref") return new FormulaError("#VALUE!");
    return ctx.hasFormula?.(target.ref.sheet ?? ctx.sheet, target.ref.row, target.ref.col) ?? false;
  },
  SHEETS: (_args, { ctx }) => (ctx.sheetNames ? ctx.sheetNames().length : 1),
  SHEET(args, { ctx }) {
    const target = args[0];
    const name = target && target.t !== "missing" && (target.t === "ref" || target.t === "range") ? (target.ref.sheet ?? ctx.sheet) : ctx.sheet;
    const names = ctx.sheetNames?.() ?? [ctx.sheet];
    const at = names.findIndex((candidate) => candidate.toLowerCase() === name.toLowerCase());
    return at < 0 ? new FormulaError("#N/A") : at + 1;
  },
  ROW(args, { ctx }) {
    const target = args[0];
    if (!target || target.t === "missing") return ctx.row === undefined ? new FormulaError("#VALUE!") : ctx.row + 1;
    if (target.t === "ref") return target.ref.row + 1;
    if (target.t === "range") return target.ref.r1 + 1;
    return new FormulaError("#VALUE!");
  },
  COLUMN(args, { ctx }) {
    const target = args[0];
    if (!target || target.t === "missing") return ctx.col === undefined ? new FormulaError("#VALUE!") : ctx.col + 1;
    if (target.t === "ref") return target.ref.col + 1;
    if (target.t === "range") return target.ref.c1 + 1;
    return new FormulaError("#VALUE!");
  },
};

/** AND (all = true) and OR (all = false): booleans and numbers count, text and empty cells in ranges are skipped. */
function allOrAny(args: Node[], { evaluate }: LazyContext, all: boolean): Value {
  let seen = false;
  let result = all;
  const visit = (value: Value, fromRange: boolean): FormulaError | null => {
    if (isArray(value)) {
      for (const row of value) for (const item of row) {
        const failure = visit(item, true);
        if (failure) return failure;
      }
      return null;
    }
    if (isError(value)) return value;
    if (value === null || (typeof value === "string" && fromRange)) return null;
    const flag = toBoolean(value);
    if (isError(flag)) return flag;
    seen = true;
    result = all ? result && flag : result || flag;
    return null;
  };
  for (const arg of args) {
    const failure = visit(evaluate(arg), false);
    if (failure) return failure;
  }
  return seen ? result : new FormulaError("#VALUE!");
}

type Implementation = (...args: unknown[]) => unknown;

/** Our own versions, taking already-evaluated arguments. */
const OWN: Record<string, (args: (Value | undefined)[], ctx: Context) => Value> = {
  NA: () => new FormulaError("#N/A"),
  TRUE: () => true,
  FALSE: () => false,
  PI: () => Math.PI,
  DATE(args) {
    const parts = args.slice(0, 3).map((arg) => toNumber(scalar(arg ?? null)));
    for (const part of parts) if (isError(part)) return part;
    let [year, month, day] = parts as number[];
    if (parts.length < 3) return new FormulaError("#VALUE!");
    year = Math.trunc(year);
    if (year < 0 || year > 9999) return new FormulaError("#NUM!");
    if (year < 1900) year += 1900;
    const serial = (Date.UTC(year, Math.trunc(month) - 1, Math.trunc(day)) - EPOCH) / DAY;
    return serial < 0 ? new FormulaError("#NUM!") : serial;
  },
  TODAY: (_args, ctx) => Math.floor(dateToSerial((ctx.now ?? (() => new Date()))())),
  NOW: (_args, ctx) => dateToSerial((ctx.now ?? (() => new Date()))()),
  RAND: (_args, ctx) => (ctx.random ?? Math.random)(),
  RANDBETWEEN(args, ctx) {
    const low = toNumber(scalar(args[0] ?? null));
    const high = toNumber(scalar(args[1] ?? null));
    if (isError(low)) return low;
    if (isError(high)) return high;
    if (low > high) return new FormulaError("#NUM!");
    return Math.ceil(low) + Math.floor((ctx.random ?? Math.random)() * (Math.floor(high) - Math.ceil(low) + 1));
  },
};

Object.assign(OWN, EXTRA, BOND_FUNCTIONS, {
  // T.DIST accepts any x; formulajs rejects 0. The cumulative form comes from the right tail.
  "T.DIST"(args: (Value | undefined)[], ctx: Context): Value {
    const x = toNumber(scalar(args[0] ?? null));
    const df = args[1];
    const cumulative = toBoolean(scalar(args[2] ?? true));
    if (isError(x)) return x;
    if (isError(cumulative)) return cumulative;
    if (!cumulative) {
      // The density, from its definition (formulajs 4.6 returns #NUM! for every call).
      const n = toNumber(scalar(df ?? null));
      if (isError(n)) return n;
      if (n < 1) return new FormulaError("#NUM!");
      const lnA = callFunction("GAMMALN", [(n + 1) / 2], ctx, true);
      const lnB = callFunction("GAMMALN", [n / 2], ctx, true);
      if (typeof lnA !== "number" || typeof lnB !== "number") return new FormulaError("#NUM!");
      return Math.exp(lnA - lnB) / Math.sqrt(n * Math.PI) * (1 + (x * x) / n) ** (-(n + 1) / 2);
    }
    if (x === 0) return 0.5;
    const tail = callFunction("T.DIST.RT", [Math.abs(x), df], ctx);
    if (isError(tail) || typeof tail !== "number") return tail;
    return x > 0 ? 1 - tail : tail;
  },
});

/**
 * Excel's classic names, which formulajs 4 only has under their newer dotted names (STDEV is
 * STDEV.S there). FortuneSheet's formulajs 2.9 had them, so sheets in the wild use them.
 */
const LEGACY_NAMES: Record<string, string> = {
  STDEV: "STDEV.S",
  STDEVP: "STDEV.P",
  VAR: "VAR.S",
  VARP: "VAR.P",
  MODE: "MODE.SNGL",
  RANK: "RANK.EQ",
  PERCENTILE: "PERCENTILE.INC",
  QUARTILE: "QUARTILE.INC",
  PERCENTRANK: "PERCENTRANK.INC",
  WEIBULL: "WEIBULL.DIST",
  POISSON: "POISSON.DIST",
  COVAR: "COVARIANCE.P",
  CONFIDENCE: "CONFIDENCE.NORM",
  NORMDIST: "NORM.DIST",
  NORMINV: "NORM.INV",
  NORMSINV: "NORM.S.INV",
  BINOMDIST: "BINOM.DIST",
  CRITBINOM: "BINOM.INV",
  EXPONDIST: "EXPON.DIST",
  GAMMADIST: "GAMMA.DIST",
  GAMMAINV: "GAMMA.INV",
  LOGINV: "LOGNORM.INV",
  HYPGEOMDIST: "HYPGEOM.DIST",
  NEGBINOMDIST: "NEGBINOM.DIST",
  TINV: "T.INV.2T",
  CHIDIST: "CHISQ.DIST.RT",
  CHIINV: "CHISQ.INV.RT",
  FDIST: "F.DIST.RT",
  FINV: "F.INV.RT",
  BETADIST: "BETA.DIST",
  BETAINV: "BETA.INV",
  ZTEST: "Z.TEST",
  FTEST: "F.TEST",
  TTEST: "T.TEST",
  ERFPRECISE: "ERF",
  "ERF.PRECISE": "ERF",
  ERFCPRECISE: "ERFC",
  "ERFC.PRECISE": "ERFC",
};

/** Classic functions whose arguments differ from the modern one they became, so a rename is not enough. */
const LEGACY_ADAPTERS: Record<string, (args: unknown[]) => [name: string, args: unknown[]]> = {
  // NORMSDIST(z) is always the cumulative curve.
  NORMSDIST: (args) => ["NORM.S.DIST", [args[0], true]],
  // LOGNORMDIST(x, mean, sd) is always cumulative.
  LOGNORMDIST: (args) => ["LOGNORM.DIST", [args[0], args[1], args[2], true]],
  // TDIST(x, df, tails): one tail is the right tail, two is the two-tailed distribution.
  TDIST: (args) => [args[2] === 1 ? "T.DIST.RT" : "T.DIST.2T", [args[0], args[1]]],
};

/** formulajs keys dotted names as nested objects: NORM.DIST is formulajs.NORM.DIST. */
function findImplementation(requested: string): Implementation | undefined {
  const api = formulajs as unknown as Record<string, unknown>;
  // Adapters first: formulajs 4 ships its own NORMSDIST and LOGNORMDIST that return the density
  // curve where Excel returns the cumulative value.
  const adapter = LEGACY_ADAPTERS[requested];
  if (adapter) {
    return (...args: unknown[]) => {
      const [name, adapted] = adapter(args);
      return findImplementation(name)?.(...adapted) ?? new Error("#NAME?");
    };
  }
  // A classic name goes to its modern equivalent even when formulajs also exports the old name,
  // because those old exports are wrong for several distributions (CHIDIST returns the density).
  const modern = LEGACY_NAMES[requested];
  const direct = modern === undefined ? api[requested] : undefined;
  if (typeof direct === "function") return direct as Implementation;
  const name = modern ?? requested;
  let node: unknown = api;
  for (const part of name.replace(/_/g, ".").split(".")) {
    node = (node as Record<string, unknown> | undefined)?.[part];
    if (node === undefined) return undefined;
  }
  return typeof node === "function" ? (node as Implementation) : undefined;
}

const toLibrary = (value: Value | undefined): unknown =>
  value === undefined ? undefined : isError(value) ? new Error(value.code) : isArray(value) ? value.map((row) => row.map(toLibrary)) : value;

function fromLibrary(result: unknown): Value {
  if (result instanceof Error) {
    const code = ERROR_CODES.find((candidate) => candidate === result.message);
    return new FormulaError(code ?? "#VALUE!");
  }
  if (result instanceof Date) return dateToSerial(result);
  if (typeof result === "number") return Number.isFinite(result) ? result : new FormulaError("#NUM!");
  if (Array.isArray(result)) {
    // A flat array is a single row.
    const rows = result.length > 0 && Array.isArray(result[0]) ? (result as unknown[][]) : [result as unknown[]];
    return rows.map((row) => row.map(fromLibrary));
  }
  if (result === undefined) return null;
  if (typeof result === "string" || typeof result === "boolean" || result === null) return result;
  return new FormulaError("#VALUE!");
}

export function callFunction(name: string, args: (Value | undefined)[], ctx: Context, library = false): Value {
  const own = library ? undefined : OWN[name];
  if (own) return own(args, ctx);
  const implementation = findImplementation(name);
  if (!implementation) return new FormulaError("#NAME?");
  try {
    const result = fromLibrary(implementation(...args.map(toLibrary)));
    return typeof result === "string" && UPPERCASE_RESULTS.has(name) ? result.toUpperCase() : result;
  } catch {
    return new FormulaError("#VALUE!");
  }
}

/** Whether a function name is known (own, lazy or formulajs). */
export function isKnownFunction(name: string): boolean {
  const upper = name.toUpperCase();
  return upper in LAZY_FUNCTIONS || upper in OWN || findImplementation(upper) !== undefined;
}
