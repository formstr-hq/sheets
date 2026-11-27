// Excel's bond / coupon / discount-security functions (COUP*, PRICE, YIELD, DURATION, ...).
// Dates are 1900-system serial numbers. All calendar math goes through Date.UTC, never local time.
import { isArray, isError, toNumber } from "./coerce";
import { FormulaError, type Value } from "./types";

type Arg = Value | undefined;
type Fn = (args: Arg[]) => Value;

const NUM = () => new FormulaError("#NUM!");
const MAX_SERIAL = 2958465; // 9999-12-31

// ---------- calendar ----------

const DAY_MS = 86400000;
const EPOCH = Date.UTC(1899, 11, 30);

const isLeap = (y: number) => (y % 4 === 0 && y % 100 !== 0) || y % 400 === 0;
const daysInMonth = (y: number, m: number) => [31, isLeap(y) ? 29 : 28, 31, 30, 31, 30, 31, 31, 30, 31, 30, 31][m - 1];

/** Serial to [year, month 1-12, day]. Serial 60 is Excel's phantom 1900-02-29. */
function ymd(serial: number): [number, number, number] {
  if (serial === 60) return [1900, 2, 29];
  const s = serial < 60 ? serial + 1 : serial;
  const d = new Date(EPOCH + s * DAY_MS);
  return [d.getUTCFullYear(), d.getUTCMonth() + 1, d.getUTCDate()];
}

function serialOf(y: number, m: number, d: number): number {
  if (y === 1900 && m === 2 && d === 29) return 60;
  const s = Math.round((Date.UTC(y, m - 1, d) - EPOCH) / DAY_MS);
  return s < 61 ? s - 1 : s;
}

/** Day difference on the 30/360 conventions: US (NASD) when `us`, otherwise European. */
function days360(a: number, b: number, us: boolean): number {
  const [y1, m1, d1o] = ymd(a);
  const [y2, m2, d2o] = ymd(b);
  let d1 = d1o;
  let d2 = d2o;
  if (us) {
    const eom1 = m1 === 2 && d1 === daysInMonth(y1, 2);
    const eom2 = m2 === 2 && d2 === daysInMonth(y2, 2);
    if (eom1 && eom2) d2 = 30;
    if (eom1) d1 = 30;
    if (d2 === 31 && d1 >= 30) d2 = 30;
    if (d1 === 31) d1 = 30;
  } else {
    if (d1 === 31) d1 = 30;
    if (d2 === 31) d2 = 30;
  }
  return (y2 - y1) * 360 + (m2 - m1) * 30 + (d2 - d1);
}

/** Days from a to b under the basis' counting rule. */
function dayDiff(a: number, b: number, basis: number): number {
  if (basis === 0) return days360(a, b, true);
  if (basis === 4) return days360(a, b, false);
  return b - a;
}

/** Excel's YEARFRAC for start <= end. */
function yearFrac(start: number, end: number, basis: number): number {
  if (basis === 0) return days360(start, end, true) / 360;
  if (basis === 4) return days360(start, end, false) / 360;
  if (basis === 2) return (end - start) / 360;
  if (basis === 3) return (end - start) / 365;
  const [y1, m1, d1] = ymd(start);
  const [y2, m2, d2] = ymd(end);
  let denom: number;
  if (y1 === y2) {
    denom = isLeap(y1) ? 366 : 365;
  } else {
    const withinYear = y2 === y1 + 1 && (m1 > m2 || (m1 === m2 && d1 >= d2));
    if (withinYear) {
      const feb29 = (y: number) => isLeap(y) && serialOf(y, 2, 29) >= start && serialOf(y, 2, 29) <= end;
      denom = feb29(y1) || feb29(y2) ? 366 : 365;
    } else {
      let total = 0;
      for (let y = y1; y <= y2; y++) total += isLeap(y) ? 366 : 365;
      denom = total / (y2 - y1 + 1);
    }
  }
  return (end - start) / denom;
}

// ---------- coupon schedule ----------

interface Coupon {
  pcd: number;
  ncd: number;
  n: number; // coupons payable between settlement and maturity
  A: number; // days from start of coupon period to settlement
  E: number; // days in the coupon period
  DSC: number; // days from settlement to next coupon
}

function coupon(settle: number, maturity: number, freq: number, basis: number): Coupon {
  const months = 12 / freq;
  const [my, mm, md] = ymd(maturity);
  const eom = md === daysInMonth(my, mm);
  const at = (k: number): number => {
    const total = my * 12 + (mm - 1) - k * months;
    const y = Math.floor(total / 12);
    const m = (((total % 12) + 12) % 12) + 1;
    return serialOf(y, m, eom ? daysInMonth(y, m) : Math.min(md, daysInMonth(y, m)));
  };
  let k = 1;
  while (at(k) > settle) k++;
  const pcd = at(k);
  const ncd = at(k - 1);
  const A = dayDiff(pcd, settle, basis);
  let E: number;
  if (basis === 1) E = ncd - pcd;
  else if (basis === 3) E = 365 / freq;
  else E = 360 / freq;
  const DSC = basis === 0 || basis === 4 ? E - A : ncd - settle;
  return { pcd, ncd, n: k, A, E, DSC };
}

// ---------- argument handling ----------

function scalar(a: Arg): Exclude<Arg, Value[][]> {
  let v: Arg = a;
  while (v !== undefined && isArray(v)) v = v[0]?.[0] ?? null;
  return v as Exclude<Arg, Value[][]>;
}

function num(a: Arg): number | FormulaError {
  const s = scalar(a);
  if (s === undefined) return new FormulaError("#VALUE!");
  const n = toNumber(s);
  if (isError(n)) return n;
  return Number.isFinite(n) ? n : NUM();
}

function int(a: Arg): number | FormulaError {
  const n = num(a);
  return isError(n) ? n : Math.trunc(n);
}

function date(a: Arg): number | FormulaError {
  const n = int(a);
  if (isError(n)) return n;
  return n < 0 || n > MAX_SERIAL ? NUM() : n;
}

function optBasis(a: Arg): number | FormulaError {
  if (a === undefined) return 0;
  const b = int(a);
  if (isError(b)) return b;
  return b < 0 || b > 4 ? NUM() : b;
}

function freqArg(a: Arg): number | FormulaError {
  const f = int(a);
  if (isError(f)) return f;
  return f === 1 || f === 2 || f === 4 ? f : NUM();
}

/** Evaluate argument converters in order; the first error wins. */
function parse<T extends unknown[]>(...parts: { [K in keyof T]: number | FormulaError }): T | FormulaError {
  for (const p of parts) if (isError(p)) return p;
  return parts as unknown as T;
}

const minArgs = (args: Arg[], min: number, max: number): FormulaError | null =>
  args.length < min || args.length > max ? new FormulaError("#VALUE!") : null;

const finite = (n: number): number | FormulaError => (Number.isFinite(n) ? n : NUM());

// ---------- COUP* ----------

function coupFn(pick: (c: Coupon, freq: number, basis: number) => number): Fn {
  return (args) => {
    const bad = minArgs(args, 3, 4);
    if (bad) return bad;
    const p = parse<[number, number, number, number]>(date(args[0]), date(args[1]), freqArg(args[2]), optBasis(args[3]));
    if (isError(p)) return p;
    const [s, m, f, b] = p;
    if (s >= m) return NUM();
    return pick(coupon(s, m, f, b), f, b);
  };
}

const COUPNUM = coupFn((c) => c.n);
const COUPPCD = coupFn((c) => c.pcd);
const COUPNCD = coupFn((c) => c.ncd);
const COUPDAYBS = coupFn((c) => c.A);
const COUPDAYS = coupFn((c) => c.E);
const COUPDAYSNC = coupFn((c) => c.DSC);

// ---------- PRICE / YIELD / DURATION ----------

function dirtyPrice(c: Coupon, rate: number, yld: number, red: number, f: number): number {
  const cpn = (100 * rate) / f;
  if (c.n === 1) return (red + cpn) / (1 + ((c.DSC / c.E) * yld) / f) - (cpn * c.A) / c.E;
  const q = 1 + yld / f;
  const t0 = c.DSC / c.E;
  let pv = red / Math.pow(q, c.n - 1 + t0);
  for (let k = 1; k <= c.n; k++) pv += cpn / Math.pow(q, k - 1 + t0);
  return pv - (cpn * c.A) / c.E;
}

/** Solve g(x) = target for a strictly decreasing g on (lo, inf) with Newton steps guarded by a bracket. */
function solveDecreasing(g: (x: number) => number, target: number, lo: number): number | null {
  let hi = 1;
  let guard = 0;
  while (g(hi) > target) {
    hi *= 2;
    if (++guard > 60) return null;
  }
  if (g(lo) < target) return null;
  let x = Math.min(Math.max(0.1, lo + 1e-6), hi);
  for (let i = 0; i < 200; i++) {
    const fx = g(x) - target;
    if (fx === 0 || Math.abs(fx) <= 1e-13 * Math.max(1, Math.abs(target))) return x;
    if (fx > 0) lo = x;
    else hi = x;
    const h = Math.max(1e-7, Math.abs(x) * 1e-7);
    const slope = (g(x + h) - g(x - h)) / (2 * h);
    let next = slope < 0 ? x - fx / slope : NaN;
    if (!(next > lo && next < hi)) next = (lo + hi) / 2;
    if (Math.abs(next - x) < 1e-15 * Math.max(1, Math.abs(x)) && Math.abs(hi - lo) < 1e-12) return next;
    x = next;
  }
  return x;
}

const PRICE: Fn = (args) => {
  const bad = minArgs(args, 6, 7);
  if (bad) return bad;
  const p = parse<[number, number, number, number, number, number, number]>(
    date(args[0]), date(args[1]), num(args[2]), num(args[3]), num(args[4]), freqArg(args[5]), optBasis(args[6]),
  );
  if (isError(p)) return p;
  const [s, m, rate, yld, red, f, b] = p;
  if (rate < 0 || yld < 0 || red <= 0 || s >= m) return NUM();
  return finite(dirtyPrice(coupon(s, m, f, b), rate, yld, red, f));
};

const YIELD: Fn = (args) => {
  const bad = minArgs(args, 6, 7);
  if (bad) return bad;
  const p = parse<[number, number, number, number, number, number, number]>(
    date(args[0]), date(args[1]), num(args[2]), num(args[3]), num(args[4]), freqArg(args[5]), optBasis(args[6]),
  );
  if (isError(p)) return p;
  const [s, m, rate, pr, red, f, b] = p;
  if (rate < 0 || pr <= 0 || red <= 0 || s >= m) return NUM();
  const c = coupon(s, m, f, b);
  if (c.n === 1) {
    const cpn = (100 * rate) / f;
    const base = pr + (c.A / c.E) * cpn;
    return finite((((red + cpn - base) / base) * f * c.E) / c.DSC);
  }
  const y = solveDecreasing((x) => dirtyPrice(c, rate, x, red, f), pr, -f + 1e-9);
  return y === null ? NUM() : y;
};

function durationCore(args: Arg[], modified: boolean): Value {
  const bad = minArgs(args, 5, 6);
  if (bad) return bad;
  const p = parse<[number, number, number, number, number, number]>(
    date(args[0]), date(args[1]), num(args[2]), num(args[3]), freqArg(args[4]), optBasis(args[5]),
  );
  if (isError(p)) return p;
  const [s, m, rate, yld, f, b] = p;
  if (rate < 0 || yld < 0 || s >= m) return NUM();
  const c = coupon(s, m, f, b);
  const q = 1 + yld / f;
  const cpn = (100 * rate) / f;
  const t0 = c.DSC / c.E;
  let weighted = 0;
  let total = 0;
  for (let k = 1; k <= c.n; k++) {
    const t = t0 + k - 1;
    const cf = k === c.n ? cpn + 100 : cpn;
    const pv = cf / Math.pow(q, t);
    weighted += t * pv;
    total += pv;
  }
  const dur = weighted / total / f;
  return finite(modified ? dur / q : dur);
}

const DURATION: Fn = (args) => durationCore(args, false);
const MDURATION: Fn = (args) => durationCore(args, true);

// ---------- securities paying at maturity / discount securities ----------

const PRICEMAT: Fn = (args) => {
  const bad = minArgs(args, 5, 6);
  if (bad) return bad;
  const p = parse<[number, number, number, number, number, number]>(
    date(args[0]), date(args[1]), date(args[2]), num(args[3]), num(args[4]), optBasis(args[5]),
  );
  if (isError(p)) return p;
  const [s, m, issue, rate, yld, b] = p;
  if (rate < 0 || yld < 0 || s >= m || issue > s) return NUM();
  const dim = yearFrac(issue, m, b);
  const a = yearFrac(issue, s, b);
  const dsm = yearFrac(s, m, b);
  return finite((100 + dim * rate * 100) / (1 + dsm * yld) - a * rate * 100);
};

const YIELDMAT: Fn = (args) => {
  const bad = minArgs(args, 5, 6);
  if (bad) return bad;
  const p = parse<[number, number, number, number, number, number]>(
    date(args[0]), date(args[1]), date(args[2]), num(args[3]), num(args[4]), optBasis(args[5]),
  );
  if (isError(p)) return p;
  const [s, m, issue, rate, pr, b] = p;
  if (rate < 0 || pr <= 0 || s >= m || issue > s) return NUM();
  const dim = yearFrac(issue, m, b);
  const a = yearFrac(issue, s, b);
  const dsm = yearFrac(s, m, b);
  const denom = pr / 100 + a * rate;
  return finite((1 + dim * rate - denom) / denom / dsm);
};

const ACCRINTM: Fn = (args) => {
  const bad = minArgs(args, 4, 5);
  if (bad) return bad;
  const par = args[3] === undefined ? 1000 : num(args[3]);
  const p = parse<[number, number, number, number, number]>(
    date(args[0]), date(args[1]), num(args[2]), par, optBasis(args[4]),
  );
  if (isError(p)) return p;
  const [issue, settle, rate, parv, b] = p;
  if (rate <= 0 || parv <= 0 || issue >= settle) return NUM();
  return finite(parv * rate * yearFrac(issue, settle, b));
};

const INTRATE: Fn = (args) => {
  const bad = minArgs(args, 4, 5);
  if (bad) return bad;
  const p = parse<[number, number, number, number, number]>(
    date(args[0]), date(args[1]), num(args[2]), num(args[3]), optBasis(args[4]),
  );
  if (isError(p)) return p;
  const [s, m, inv, red, b] = p;
  if (inv <= 0 || red <= 0 || s >= m) return NUM();
  return finite((red - inv) / inv / yearFrac(s, m, b));
};

const RECEIVED: Fn = (args) => {
  const bad = minArgs(args, 4, 5);
  if (bad) return bad;
  const p = parse<[number, number, number, number, number]>(
    date(args[0]), date(args[1]), num(args[2]), num(args[3]), optBasis(args[4]),
  );
  if (isError(p)) return p;
  const [s, m, inv, disc, b] = p;
  if (inv <= 0 || disc <= 0 || s >= m) return NUM();
  const denom = 1 - disc * yearFrac(s, m, b);
  return denom <= 0 ? NUM() : finite(inv / denom);
};

const YIELDDISC: Fn = (args) => {
  const bad = minArgs(args, 4, 5);
  if (bad) return bad;
  const p = parse<[number, number, number, number, number]>(
    date(args[0]), date(args[1]), num(args[2]), num(args[3]), optBasis(args[4]),
  );
  if (isError(p)) return p;
  const [s, m, pr, red, b] = p;
  if (pr <= 0 || red <= 0 || s >= m) return NUM();
  return finite((red - pr) / pr / yearFrac(s, m, b));
};

// ---------- odd first / last period ----------

/** Same calendar step as `coupon`: `k` coupon periods from `anchor`, end-of-month aware. */
function shiftPeriods(anchor: number, k: number, months: number): number {
  const [y0, m0, d0] = ymd(anchor);
  const eom = d0 === daysInMonth(y0, m0);
  const total = y0 * 12 + (m0 - 1) + k * months;
  const y = Math.floor(total / 12);
  const m = (((total % 12) + 12) % 12) + 1;
  return serialOf(y, m, eom ? daysInMonth(y, m) : Math.min(d0, daysInMonth(y, m)));
}

/** Normal length of a (quasi) coupon period under the basis. */
function normalLength(lo: number, hi: number, freq: number, basis: number): number {
  if (basis === 1) return hi - lo;
  return basis === 3 ? 365 / freq : 360 / freq;
}

/** Sum over quasi coupon periods of (days of [a, b] inside the period) / (normal period length). */
function quasiFraction(periods: [number, number][], a: number, b: number, freq: number, basis: number): number {
  let sum = 0;
  for (const [lo, hi] of periods) {
    const from = Math.max(lo, a);
    const to = Math.min(hi, b);
    if (to > from) sum += dayDiff(from, to, basis) / normalLength(lo, hi, freq, basis);
  }
  return sum;
}

interface OddFirst {
  cpn: number; // coupon per period
  first: number; // fraction of a regular coupon paid on the first coupon date
  accrued: number; // fraction of a regular coupon accrued at settlement
  t1: number; // periods from settlement to the first coupon
  nc: number; // coupons from first coupon to maturity inclusive
}

function oddFirstSetup(s: number, m: number, issue: number, fc: number, rate: number, f: number, b: number): OddFirst {
  const months = 12 / f;
  // quasi coupon dates going back from the first coupon until the issue date is covered
  const dates = [fc];
  while (dates[dates.length - 1] > issue) dates.push(shiftPeriods(fc, -dates.length, months));
  // periods[j] = [dates[j+1], dates[j]], latest first
  const periods: [number, number][] = [];
  for (let j = 0; j + 1 < dates.length; j++) periods.push([dates[j + 1], dates[j]]);
  // quasi period containing settlement
  let idx = periods.findIndex(([lo, hi]) => lo <= s && s < hi);
  if (idx < 0) idx = periods.length - 1;
  const [lo, hi] = periods[idx];
  const E = normalLength(lo, hi, f, b);
  const DSC = b === 0 || b === 4 ? E - dayDiff(lo, s, b) : hi - s;
  const nc = coupon(fc, m, f, b).n + 1;
  return {
    cpn: (100 * rate) / f,
    first: quasiFraction(periods, issue, fc, f, b),
    accrued: quasiFraction(periods, issue, s, f, b),
    t1: idx + DSC / E,
    nc,
  };
}

function oddFirstPrice(o: OddFirst, yld: number, red: number, f: number): number {
  const q = 1 + yld / f;
  let pv = 0;
  for (let j = 1; j <= o.nc; j++) {
    let cf = j === 1 ? o.cpn * o.first : o.cpn;
    if (j === o.nc) cf += red;
    pv += cf / Math.pow(q, o.t1 + j - 1);
  }
  return pv - o.cpn * o.accrued;
}

interface OddFirstArgs {
  s: number;
  m: number;
  issue: number;
  fc: number;
  rate: number;
  x: number; // yield or price
  red: number;
  f: number;
  b: number;
}

function parseOddFirst(args: Arg[]): OddFirstArgs | FormulaError {
  const bad = minArgs(args, 8, 9);
  if (bad) return bad;
  const p = parse<[number, number, number, number, number, number, number, number, number]>(
    date(args[0]), date(args[1]), date(args[2]), date(args[3]), num(args[4]), num(args[5]), num(args[6]), freqArg(args[7]), optBasis(args[8]),
  );
  if (isError(p)) return p;
  const [s, m, issue, fc, rate, x, red, f, b] = p;
  if (!(m > fc && fc > s && s > issue) || rate < 0 || red <= 0) return NUM();
  return { s, m, issue, fc, rate, x, red, f, b };
}

const ODDFPRICE: Fn = (args) => {
  const a = parseOddFirst(args);
  if (isError(a)) return a;
  if (a.x < 0) return NUM();
  const o = oddFirstSetup(a.s, a.m, a.issue, a.fc, a.rate, a.f, a.b);
  return finite(oddFirstPrice(o, a.x, a.red, a.f));
};

const ODDFYIELD: Fn = (args) => {
  const a = parseOddFirst(args);
  if (isError(a)) return a;
  if (a.x <= 0) return NUM();
  const o = oddFirstSetup(a.s, a.m, a.issue, a.fc, a.rate, a.f, a.b);
  const y = solveDecreasing((x) => oddFirstPrice(o, x, a.red, a.f), a.x, -a.f + 1e-9);
  return y === null ? NUM() : y;
};

interface OddLast {
  cpn: number;
  dc: number; // odd period length in regular periods
  accrued: number;
  dsc: number; // settlement to maturity in regular periods
}

function oddLastSetup(s: number, m: number, last: number, rate: number, f: number, b: number): OddLast {
  const months = 12 / f;
  const dates = [last];
  while (dates[dates.length - 1] < m) dates.push(shiftPeriods(last, dates.length, months));
  const periods: [number, number][] = [];
  for (let j = 0; j + 1 < dates.length; j++) periods.push([dates[j], dates[j + 1]]);
  return {
    cpn: (100 * rate) / f,
    dc: quasiFraction(periods, last, m, f, b),
    accrued: quasiFraction(periods, last, s, f, b),
    dsc: quasiFraction(periods, s, m, f, b),
  };
}

function parseOddLast(args: Arg[]): { s: number; m: number; last: number; rate: number; x: number; red: number; f: number; b: number } | FormulaError {
  const bad = minArgs(args, 7, 8);
  if (bad) return bad;
  const p = parse<[number, number, number, number, number, number, number, number]>(
    date(args[0]), date(args[1]), date(args[2]), num(args[3]), num(args[4]), num(args[5]), freqArg(args[6]), optBasis(args[7]),
  );
  if (isError(p)) return p;
  const [s, m, last, rate, x, red, f, b] = p;
  if (!(m > s && s > last) || rate < 0 || red <= 0) return NUM();
  return { s, m, last, rate, x, red, f, b };
}

const ODDLPRICE: Fn = (args) => {
  const a = parseOddLast(args);
  if (isError(a)) return a;
  if (a.x < 0) return NUM();
  const o = oddLastSetup(a.s, a.m, a.last, a.rate, a.f, a.b);
  return finite((a.red + o.cpn * o.dc) / (1 + (a.x / a.f) * o.dsc) - o.cpn * o.accrued);
};

const ODDLYIELD: Fn = (args) => {
  const a = parseOddLast(args);
  if (isError(a)) return a;
  if (a.x <= 0) return NUM();
  const o = oddLastSetup(a.s, a.m, a.last, a.rate, a.f, a.b);
  const base = a.x + o.cpn * o.accrued;
  return finite((((a.red + o.cpn * o.dc - base) / base) * a.f) / o.dsc);
};

// ---------- VDB ----------

/** Double-declining-balance depreciation for one whole period (Excel DDB). */
function ddb(cost: number, salvage: number, life: number, period: number, factor: number): number {
  let rate = factor / life;
  let old: number;
  if (rate >= 1) {
    rate = 1;
    old = period === 1 ? cost : 0;
  } else {
    old = cost * Math.pow(1 - rate, period - 1);
  }
  const next = cost * Math.pow(1 - rate, period);
  const d = next < salvage ? old - salvage : old - next;
  return d < 0 ? 0 : d;
}

/** Declining balance with a switch to straight-line when that depreciates more. */
function interVdb(cost: number, salvage: number, life: number, life1: number, period: number, factor: number): number {
  const end = Math.ceil(period);
  let total = 0;
  let rest = cost - salvage;
  let straight = 0;
  let switched = false;
  for (let i = 1; i <= end; i++) {
    let term: number;
    if (!switched) {
      const declining = ddb(cost, salvage, life, i, factor);
      straight = rest / (life1 - (i - 1));
      if (straight > declining) {
        term = straight;
        switched = true;
      } else {
        term = declining;
        rest -= declining;
      }
    } else {
      term = straight;
    }
    if (i === end) term *= period + 1 - end;
    total += term;
  }
  return total;
}

const VDB: Fn = (args) => {
  const bad = minArgs(args, 5, 7);
  if (bad) return bad;
  const factorArg = args[5] === undefined || scalar(args[5]) === null ? 2 : num(args[5]);
  let noSwitch = false;
  if (args[6] !== undefined) {
    const n = num(args[6]);
    if (isError(n)) return n;
    noSwitch = n !== 0;
  }
  const p = parse<[number, number, number, number, number, number]>(
    num(args[0]), num(args[1]), num(args[2]), num(args[3]), num(args[4]), factorArg,
  );
  if (isError(p)) return p;
  const [cost, salvage, life, start0, end0, factor] = p;
  if (cost < 0 || salvage < 0 || life <= 0 || start0 < 0 || end0 < start0 || end0 > life || factor <= 0) return NUM();
  let start = start0;
  let end = end0;
  let result = 0;
  if (noSwitch) {
    const first = Math.floor(start);
    const last = Math.ceil(end);
    for (let i = first + 1; i <= last; i++) {
      let term = ddb(cost, salvage, life, i, factor);
      if (i === first + 1) term *= Math.min(end, first + 1) - start;
      else if (i === last) term *= end + 1 - last;
      result += term;
    }
    return finite(result);
  }
  let life1 = life;
  if (start !== Math.floor(start) && factor > 1 && start >= life / 2) {
    const part = start - life / 2;
    start = life / 2;
    end -= part;
    life1 += 1;
  }
  const remaining = cost - interVdb(cost, salvage, life, life1, start, factor);
  result = interVdb(remaining, salvage, life, life - start, end - start, factor);
  return finite(result);
};

export const BOND_FUNCTIONS: Record<string, (args: (Value | undefined)[]) => Value> = {
  ODDFPRICE,
  ODDFYIELD,
  ODDLPRICE,
  ODDLYIELD,
  VDB,
  COUPNUM,
  COUPPCD,
  COUPNCD,
  COUPDAYBS,
  COUPDAYS,
  COUPDAYSNC,
  PRICE,
  YIELD,
  DURATION,
  MDURATION,
  PRICEMAT,
  YIELDMAT,
  ACCRINTM,
  INTRATE,
  RECEIVED,
  YIELDDISC,
};
