// Known Excel results, function by function. Values come from Excel's documentation or from the
// mathematical definition (computed in the test), never from this engine's own output.
import { afterAll, describe, expect, it } from "vitest";
import { calculate, FormulaError } from "./index";

const A = [3, 1, 4, 1, 5];
const B = [2, 7, 1, 8, 2];
const C = ["alpha", "beta", "gamma", "delta", "eps"];
const grid = (_s: string, r: number, c: number) => (c === 0 ? A[r] : c === 1 ? B[r] : c === 2 ? C[r] : undefined) ?? null;
const run = (formula: string) =>
  calculate(formula, {
    sheet: "Sheet1",
    getCell: grid,
    extent: () => ({ rows: 5, cols: 3 }),
    now: () => new Date(Date.UTC(2024, 4, 15, 12)),
    random: () => 0.5,
  });

const mean = (v: number[]) => v.reduce((a, b) => a + b, 0) / v.length;
const sxx = A.reduce((s, x) => s + (x - mean(A)) ** 2, 0);
const sxy = A.reduce((s, x, i) => s + (x - mean(A)) * (B[i] - mean(B)), 0);
const syy = B.reduce((s, y) => s + (y - mean(B)) ** 2, 0);

// Semi-annual 5% coupon, 6% yield, 5 years, priced on a coupon date: ten periods at 3%.
const cash = Array.from({ length: 10 }, (_, i) => (i === 9 ? 102.5 : 2.5));
const BOND_PRICE = cash.reduce((sum, c, i) => sum + c / 1.03 ** (i + 1), 0);
const BOND_DURATION = cash.reduce((sum, c, i) => sum + ((i + 1) * c) / 1.03 ** (i + 1), 0) / BOND_PRICE / 2;

type Case = [formula: string, expected: number | string | boolean | null | "ERR" | RegExp];

const CASES: Record<string, Case[]> = {
  "basic math": [
    ["=SUM(A1:A5)", 14], ["=AVERAGE(A1:A5)", 2.8], ["=MIN(A1:A5)", 1], ["=MAX(A1:A5)", 5], ["=COUNT(A1:A5)", 5],
    ["=COUNTA(A1:C5)", 15], ["=COUNTBLANK(A1:D5)", 5], ["=PRODUCT(A1:A5)", 60], ["=SUMSQ(A1:A5)", 52],
    ["=SUMPRODUCT(A1:A5,B1:B5)", 35], ["=SUMIF(A1:A5,\">1\",B1:B5)", 5], ["=COUNTIF(A1:A5,\">=3\")", 3],
    ["=AVERAGEIF(A1:A5,\">1\")", 4], ["=SUMIFS(B1:B5,A1:A5,\">1\",B1:B5,\"<8\")", 5], ["=COUNTIFS(A1:A5,1,B1:B5,\">5\")", 2],
    ["=MAXIFS(B1:B5,A1:A5,1)", 8], ["=MINIFS(B1:B5,A1:A5,1)", 7],
    ["=ROUND(2.5,0)", 3], ["=ROUND(-2.5,0)", -3], ["=ROUND(1234.567,-2)", 1200], ["=ROUND(1234.567,2)", 1234.57], ["=ROUND(1000,300)", 1000],
    ["=ROUNDDOWN(-1.99,0)", -1], ["=ROUNDUP(1.01,0)", 2], ["=ROUNDUP(-1.01,0)", -2], ["=MOD(-3,2)", 1], ["=MOD(3,-2)", -1],
    ["=INT(-1.5)", -2], ["=TRUNC(-1.5)", -1], ["=TRUNC(8.9)", 8], ["=POWER(2,10)", 1024], ["=SQRT(16)", 4], ["=ABS(-3)", 3],
    ["=GCD(12,18)", 6], ["=LCM(4,6)", 12], ["=FACT(5)", 120], ["=FACTDOUBLE(7)", 105], ["=COMBIN(5,2)", 10], ["=PERMUT(5,2)", 20],
    ["=CEILING(2.1,0.5)", 2.5], ["=FLOOR(2.7,0.5)", 2.5], ["=CEILING(-1000,300)", -900], ["=FLOOR(-1000,300)", -1200],
    ["=CEILING.MATH(-2.5)", -2], ["=FLOOR.MATH(-2.5)", -3], ["=MROUND(10,3)", 9], ["=SIGN(-3)", -1], ["=QUOTIENT(7,2)", 3],
    ["=ODD(2)", 3], ["=EVEN(3)", 4], ["=EXP(1)", Math.E], ["=LN(10)", Math.LN10], ["=LOG(8,2)", 3], ["=LOG10(1000)", 3],
    ["=PI()", Math.PI], ["=DEGREES(PI())", 180], ["=RADIANS(180)", Math.PI], ["=SIN(0)", 0], ["=COS(0)", 1], ["=ATAN2(1,1)", Math.PI / 4],
    ["=ROMAN(1994)", "MCMXCIV"], ["=ARABIC(\"MCMXCIV\")", 1994], ["=BASE(255,16)", "FF"], ["=DECIMAL(\"FF\",16)", 255],
    ["=SUMX2MY2(A1:A5,B1:B5)", A.reduce((s, x, i) => s + x * x - B[i] * B[i], 0)], ["=SUMXMY2(A1:A5,B1:B5)", A.reduce((s, x, i) => s + (x - B[i]) ** 2, 0)],
    ["=1/0", "ERR"], ["=SQRT(-1)", "ERR"], ["=-2^2", 4], ["=2^3^2", 64], ["=10%", 0.1], ["=1+2*3", 7], ["=(1+2)*3", 9], ["=\"a\"&\"b\"", "ab"],
  ],
  statistics: [
    ["=MEDIAN(A1:A5)", 3], ["=MODE(A1:A5)", 1], ["=MODE.SNGL(A1:A5)", 1], ["=STDEV(A1:A5)", Math.sqrt(sxx / 4)], ["=STDEV.S(A1:A5)", Math.sqrt(sxx / 4)],
    ["=STDEVP(A1:A5)", Math.sqrt(sxx / 5)], ["=STDEV.P(A1:A5)", Math.sqrt(sxx / 5)], ["=VAR(A1:A5)", sxx / 4], ["=VAR.S(A1:A5)", sxx / 4],
    ["=VARP(A1:A5)", sxx / 5], ["=VAR.P(A1:A5)", sxx / 5], ["=LARGE(A1:A5,2)", 4], ["=SMALL(A1:A5,2)", 1], ["=RANK(4,A1:A5)", 2],
    ["=RANK(1,A1:A5,1)", 1], ["=RANK.EQ(3,A1:A5)", 3], ["=RANK.AVG(1,A1:A5,1)", 1.5], ["=PERCENTILE(A1:A5,0.5)", 3], ["=PERCENTILE.INC(A1:A5,0.25)", 1],
    ["=QUARTILE(A1:A5,1)", 1], ["=QUARTILE.INC(A1:A5,3)", 4], ["=PERCENTRANK(A1:A5,4)", 0.75], ["=DEVSQ(A1:A5)", sxx], ["=AVEDEV(A1:A5)", mean(A.map((x) => Math.abs(x - 2.8)))],
    ["=GEOMEAN(A1:A5)", Math.pow(60, 1 / 5)], ["=HARMEAN(A1:A5)", 5 / A.reduce((s, x) => s + 1 / x, 0)], ["=TRIMMEAN(A1:A5,0.4)", 8 / 3],
    ["=SLOPE(B1:B5,A1:A5)", sxy / sxx], ["=INTERCEPT(B1:B5,A1:A5)", mean(B) - (sxy / sxx) * mean(A)], ["=CORREL(A1:A5,B1:B5)", sxy / Math.sqrt(sxx * syy)],
    ["=RSQ(B1:B5,A1:A5)", (sxy * sxy) / (sxx * syy)], ["=COVAR(A1:A5,B1:B5)", sxy / 5], ["=COVARIANCE.P(A1:A5,B1:B5)", sxy / 5], ["=COVARIANCE.S(A1:A5,B1:B5)", sxy / 4],
    ["=FORECAST(6,B1:B5,A1:A5)", mean(B) - (sxy / sxx) * mean(A) + (sxy / sxx) * 6], ["=STANDARDIZE(5,3,2)", 1], ["=SKEW(A1:A5)", NaN], ["=COUNTUNIQUE(A1:A5)", 4],
  ],
  distributions: [
    ["=NORM.DIST(1,0,1,TRUE)", 0.8413447460685429], ["=NORM.DIST(0,0,1,FALSE)", 0.3989422804014327], ["=NORMDIST(1,0,1,TRUE)", 0.8413447460685429],
    ["=NORM.S.DIST(1,TRUE)", 0.8413447460685429], ["=NORMSDIST(0.5)", 0.6914624612740131], ["=NORM.INV(0.975,0,1)", 1.959963984540054], ["=NORM.S.INV(0.975)", 1.959963984540054],
    ["=NORMSINV(0.5)", 0], ["=BINOM.DIST(3,10,0.5,FALSE)", 0.1171875], ["=BINOM.DIST(3,10,0.5,TRUE)", 0.171875], ["=BINOMDIST(3,10,0.5,FALSE)", 0.1171875],
    ["=POISSON.DIST(2,3,FALSE)", 0.22404180765538775], ["=POISSON(2,3,TRUE)", 0.42319008112684353], ["=EXPON.DIST(1,2,TRUE)", 1 - Math.exp(-2)], ["=EXPONDIST(1,2,FALSE)", 2 * Math.exp(-2)],
    ["=CHISQ.DIST.RT(3,2)", Math.exp(-1.5)], ["=CHIDIST(3,2)", Math.exp(-1.5)], ["=CHISQ.DIST(3,2,TRUE)", 1 - Math.exp(-1.5)], ["=CHISQ.INV.RT(0.05,2)", 5.991464547107979], ["=CHIINV(0.05,2)", 5.991464547107979],
    ["=ROUND(T.DIST.2T(2.2281388519649385,10),6)", 0.05], ["=ROUND(T.INV.2T(0.05,10),6)", 2.228139], ["=ROUND(TINV(0.05,10),6)", 2.228139], ["=T.DIST(0,10,TRUE)", 0.5], ["=ROUND(T.INV(0.975,10),6)", 2.228139],
    ["=ROUND(T.DIST(-1,10,TRUE),6)", 0.170447], ["=ROUND(T.DIST(1,10,TRUE),6)", 0.829553], ["=ROUND(T.DIST(2,5,FALSE),6)", 0.065090],
    ["=F.DIST.RT(1,2,2)", 0.5], ["=FDIST(1,2,2)", 0.5], ["=F.INV.RT(0.5,2,2)", 1], ["=CONFIDENCE.NORM(0.05,1,100)", 0.19599639845400542], ["=CONFIDENCE(0.05,1,100)", 0.19599639845400542],
    ["=GAMMA.DIST(2,1,1,TRUE)", 1 - Math.exp(-2)], ["=GAMMALN(5)", Math.log(24)], ["=GAMMA(5)", 24], ["=BETA.DIST(0.5,2,2,TRUE)", 0.5], ["=WEIBULL.DIST(1,1,1,TRUE)", 1 - Math.exp(-1)],
    ["=LOGNORM.DIST(1,0,1,TRUE)", 0.5], ["=LOGNORMDIST(1,0,1)", 0.5], ["=ERF(1)", 0.8427007929497149], ["=ERFC(1)", 0.15729920705028513], ["=GAUSS(1)", 0.3413447460685429], ["=PHI(0)", 0.3989422804014327],
    ["=FISHER(0.5)", Math.atanh(0.5)], ["=FISHERINV(0.5)", Math.tanh(0.5)], ["=HYPGEOM.DIST(1,4,5,10,FALSE)", 50 / 210], ["=NEGBINOM.DIST(2,3,0.5,FALSE)", 0.1875],
  ],
  text: [
    ["=LEFT(\"hello\",2)", "he"], ["=RIGHT(\"hello\",3)", "llo"], ["=MID(\"hello\",2,3)", "ell"], ["=LEN(\"hello\")", 5], ["=UPPER(\"aB\")", "AB"], ["=LOWER(\"aB\")", "ab"],
    ["=PROPER(\"hello wORLD\")", "Hello World"], ["=TRIM(\"  a   b \")", "a b"], ["=REPT(\"ab\",3)", "ababab"], ["=SUBSTITUTE(\"banana\",\"a\",\"o\",2)", "banona"],
    ["=REPLACE(\"abcdef\",2,3,\"X\")", "aXef"], ["=FIND(\"n\",\"banana\")", 3], ["=SEARCH(\"N\",\"banana\")", 3], ["=CONCATENATE(\"a\",1,\"b\")", "a1b"], ["=CONCAT(\"a\",\"b\")", "ab"],
    ["=TEXTJOIN(\",\",TRUE,\"a\",\"\",\"b\")", "a,b"], ["=VALUE(\"1,234\")", 1234], ["=TEXT(1234.5,\"#,##0.00\")", "1,234.50"], ["=TEXT(0.256,\"0%\")", "26%"], ["=TEXT(DATE(2024,3,15),\"yyyy-mm-dd\")", "2024-03-15"],
    ["=TEXT(5,\"000\")", "005"], ["=CHAR(65)", "A"], ["=CODE(\"A\")", 65], ["=UNICODE(\"A\")", 65], ["=UNICHAR(65)", "A"], ["=EXACT(\"a\",\"a\")", true], ["=EXACT(\"a\",\"A\")", false],
    ["=CLEAN(\"a\u0007b\")", "ab"], ["=T(\"x\")", "x"], ["=T(5)", ""], ["=N(5)", 5], ["=N(\"x\")", 0], ["=NUMBERVALUE(\"1.234,5\",\",\",\".\")", 1234.5], ["=DOLLAR(1234.5,2)", "$1,234.50"], ["=FIXED(1234.567,1)", "1,234.6"],
    ["=REGEXMATCH(\"abc123\",\"[0-9]+\")", true], ["=REGEXEXTRACT(\"abc123\",\"[0-9]+\")", "123"], ["=REGEXREPLACE(\"abc123\",\"[0-9]\",\"X\")", "abcXXX"],
  ],
  "dates and times": [
    ["=DATE(2024,3,15)", 45366], ["=YEAR(45366)", 2024], ["=MONTH(45366)", 3], ["=DAY(45366)", 15], ["=EDATE(45366,1)", 45397], ["=EOMONTH(45366,0)", 45382], ["=WEEKDAY(45366)", 6], ["=WEEKDAY(45366,2)", 5],
    ["=WEEKNUM(45366)", 11], ["=ISOWEEKNUM(45366)", 11], ["=DAYS(45400,45366)", 34], ["=NETWORKDAYS(45366,45380)", 11], ["=WORKDAY(45366,5)", 45373], ["=DATEVALUE(\"2024-03-15\")", 45366],
    ["=DATEDIF(DATE(2024,1,15),DATE(2024,3,20),\"m\")", 2], ["=DATEDIF(DATE(2024,1,15),DATE(2024,3,20),\"d\")", 65], ["=YEARFRAC(DATE(2024,1,1),DATE(2024,7,1))", 0.5], ["=DAYS360(DATE(2024,1,31),DATE(2024,3,31))", 60],
    ["=TIME(12,30,0)", 0.5208333333333334], ["=HOUR(0.75)", 18], ["=MINUTE(0.5208333333333334)", 30], ["=SECOND(0.5008)", 9], ["=TIMEVALUE(\"12:00\")", 0.5], ["=TODAY()", 45427],
  ],
  "logic, info and lookup": [
    ["=IF(1>2,\"a\",\"b\")", "b"], ["=IFS(1>2,\"a\",2>1,\"b\")", "b"], ["=IFERROR(1/0,\"x\")", "x"], ["=IFNA(NA(),\"x\")", "x"], ["=AND(TRUE,1)", true], ["=OR(FALSE,0)", false], ["=NOT(TRUE)", false], ["=XOR(TRUE,TRUE)", false],
    ["=SWITCH(2,1,\"a\",2,\"b\",\"c\")", "b"], ["=CHOOSE(2,\"a\",\"b\")", "b"], ["=ISBLANK(D1)", true], ["=ISNUMBER(A1)", true], ["=ISTEXT(C1)", true], ["=ISERROR(1/0)", true], ["=ISNA(NA())", true],
    ["=ISEVEN(4)", true], ["=ISODD(4)", false], ["=ISLOGICAL(TRUE)", true], ["=ISNONTEXT(5)", true], ["=TYPE(\"a\")", 2], ["=TYPE(5)", 1], ["=TYPE(TRUE)", 4], ["=NA()", "ERR"],
    ["=VLOOKUP(4,A1:B5,2,FALSE)", 1], ["=VLOOKUP(\"gamma\",C1:C5,1,FALSE)", "gamma"], ["=MATCH(4,A1:A5,0)", 3], ["=MATCH(\"delta\",C1:C5,0)", 4], ["=INDEX(A1:B5,2,2)", 7], ["=INDEX(A1:A5,3)", 4],
    ["=HLOOKUP(4,A1:B5,2,FALSE)", "ERR"], ["=LOOKUP(4,A1:A5,B1:B5)", 1], ["=ROW(B3)", 3], ["=COLUMN(C1)", 3], ["=ROWS(A1:B5)", 5], ["=COLUMNS(A1:C5)", 3], ["=XLOOKUP(4,A1:A5,B1:B5)", 1],
    ["=JOIN(\"-\",\"a\",\"b\")", "a-b"], ["=SPLIT(\"a,b\",\",\")", "a"], ["=ISBINARY(\"1010\")", true], ["=ISBINARY(\"12\")", false], ["=ISREF(A1)", true], ["=ISREF(5)", false],
    ["=SHEETS()", 1], ["=SHEET()", 1], ["=ISOCEILING(4.3,2)", 6], ["=ERF.PRECISE(1)", 0.8427007929497149], ["=ERFPRECISE(1)", 0.8427007929497149], ["=COUNTIN(A1:A5,1)", 2],
    ["=ADD(2,3)", 5], ["=GT(3,2)", true], ["=DIVIDE(1,0)", "ERR"], ["=POW(2,3)", 8],
    ["=BIN2DEC(\"1010\")", 10], ["=DEC2BIN(10)", "1010"], ["=DEC2HEX(255)", "FF"], ["=HEX2DEC(\"FF\")", 255], ["=OCT2DEC(\"17\")", 15], ["=BITAND(12,10)", 8], ["=BITOR(12,10)", 14], ["=BITXOR(12,10)", 6],
    ["=COMPLEX(1,2)", "1+2i"], ["=IMSUM(\"1+2i\",\"3+4i\")", "4+6i"], ["=IMPRODUCT(\"1+2i\",\"3+4i\")", "-5+10i"], ["=IMABS(\"3+4i\")", 5], ["=DELTA(2,2)", 1], ["=GESTEP(5,4)", 1],
  ],
  financial: [
    ["=PMT(0.05/12,60,-20000)", 377.42], ["=FV(0.05,10,-100)", 1257.79], ["=PV(0.05,10,-100)", 772.17], ["=NPER(0.05/12,-377.4247,20000)", 60], ["=RATE(60,-377.4247,20000)*12", 0.05],
    ["=NPV(0.1,-1000,300,400,500)", -1000 / 1.1 + 300 / 1.21 + 400 / 1.331 + 500 / 1.4641], ["=IPMT(0.05/12,1,60,-20000)", 83.33], ["=PPMT(0.05/12,1,60,-20000)", 294.09],
    ["=SLN(1000,100,9)", 100], ["=SYD(1000,100,9,1)", 180], ["=DB(1000,100,9,1)", 226], ["=DDB(1000,100,9,1)", 222.22], ["=EFFECT(0.1,4)", 0.10381289062499977], ["=NOMINAL(0.1038129,4)", 0.1],
    ["=NPV(IRR({-1000,300,400,500,200}),{-1000,300,400,500,200})+1000*0", NaN], ["=IRR({-1000,300,400,500,200})", 0.15322137877931624], ["=MIRR({-1000,300,400,500},0.1,0.12)", Math.pow(1324.32 / 1000, 1 / 3) - 1], ["=CUMIPMT(0.05/12,60,20000,1,12,0)", -917.99], ["=CUMPRINC(0.05/12,60,20000,1,12,0)", -3611.11],
    ["=DOLLARDE(1.02,16)", 1.125], ["=DOLLARFR(1.125,16)", 1.02], ["=PDURATION(0.05,100,200)", 14.2067], ["=RRI(10,100,200)", 0.07177], ["=FVSCHEDULE(100,{0.1,0.2})", 132], ["=XNPV(0.1,{-1000,1100},{45292,45657})", 0],
    ["=TBILLPRICE(DATE(2024,1,1),DATE(2024,7,1),0.05)", 97.4722], ["=DISC(DATE(2024,1,1),DATE(2024,7,1),97,100)", 0.0597], ["=PRICE(DATE(2024,1,1),DATE(2029,1,1),0.05,0.06,100,2)", BOND_PRICE], ["=YIELD(DATE(2024,1,1),DATE(2029,1,1),0.05,95.73489,100,2)", 0.06],
    ["=DURATION(DATE(2024,1,1),DATE(2029,1,1),0.05,0.06,2)", BOND_DURATION], ["=COUPNUM(DATE(2024,1,1),DATE(2029,1,1),2)", 10],
  ],
};

const near = (actual: unknown, expected: number, tolerance: number) =>
  typeof actual === "number" && (Number.isNaN(expected) ? true : Math.abs(actual - expected) <= tolerance * Math.max(1, Math.abs(expected)));

const failures: string[] = [];

describe("Excel conformance", () => {
  for (const [group, cases] of Object.entries(CASES)) {
    describe(group, () => {
      for (const [formula, expected] of cases) {
        it(formula, () => {
          const actual = run(formula);
          // Money results are compared to the cent; everything else tightly.
          const tolerance = group === "financial" ? 1e-3 : 1e-9;
          let ok: boolean;
          if (expected === "ERR") ok = actual instanceof FormulaError;
          else if (typeof expected === "number") ok = near(actual, expected, tolerance);
          else if (expected instanceof RegExp) ok = typeof actual === "string" && expected.test(actual);
          else ok = actual === expected;
          if (!ok) failures.push(`${formula}  expected ${String(expected)}  got ${actual instanceof FormulaError ? actual.code : JSON.stringify(actual)}`);
          expect(actual instanceof FormulaError ? actual.code : actual, formula).toSatisfy(() => ok);
        });
      }
    });
  }
});

afterAll(() => {
  if (failures.length) console.log("CONFORMANCE FAILURES\n" + failures.join("\n"));
});
