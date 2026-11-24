import { describe, expect, it } from "vitest";
import {
  FormulaError,
  ParseError,
  calculate,
  isKnownFunction,
  parse,
  print,
  references,
  renameSheet,
  rewriteForEdit,
  shiftReferences,
  type Context,
  type Scalar,
} from "./index";

/** A small in-memory workbook: sheets of "row:col" -> value, 0-based. */
function book(sheets: Record<string, Record<string, Scalar>>, extra: Partial<Context> = {}): Context {
  const extent = (sheet: string) => {
    let rows = 0;
    let cols = 0;
    for (const key of Object.keys(sheets[sheet] ?? {})) {
      const [r, c] = key.split(":").map(Number);
      rows = Math.max(rows, r + 1);
      cols = Math.max(cols, c + 1);
    }
    return { rows, cols };
  };
  return {
    sheet: "Sheet1",
    getCell: (sheet, row, col) => sheets[sheet]?.[`${row}:${col}`],
    extent,
    ...extra,
  };
}

const calc = (formula: string, ctx: Context = book({ Sheet1: {} })) => calculate(formula, ctx);
const code = (value: unknown) => (value instanceof FormulaError ? value.code : value);

describe("parse and print", () => {
  const cases = [
    "=1+2*3",
    "=(1+2)*3",
    "=A1+B2",
    "=$A$1+B$2+$C3",
    "=SUM(A1:B2,C3)",
    "=SUM(A:A)",
    "=SUM(1:1)",
    '=IF(A1>=10,"big","small")',
    "=Sheet2!A1+'My Sheet'!B2:C3",
    '="He said ""hi"""',
    "=50%*A1",
    "=-A1^2",
    "=A1&B1&\"x\"",
    "={1,2;3,4}",
    "=A1-(B1-C1)",
    "=NORM.DIST(1,0,1,TRUE)",
    "=SUM(1,,2)",
    "=#N/A",
  ];
  for (const formula of cases) {
    it(`round-trips ${formula}`, () => {
      const tree = parse(formula);
      expect(parse(print(tree))).toEqual(tree);
    });
  }

  it("rejects broken formulas with a position", () => {
    expect(() => parse("=1+")).toThrow(ParseError);
    expect(() => parse("=SUM(1")).toThrow(ParseError);
    expect(() => parse('="open')).toThrow(ParseError);
  });
});

describe("operators", () => {
  it("follows Excel's precedence", () => {
    expect(calc("=1+2*3")).toBe(7);
    expect(calc("=2*3+1")).toBe(7);
    expect(calc("=-2^2")).toBe(4); // unary minus binds tighter than ^ in Excel
    expect(calc("=2^-1")).toBe(0.5);
    expect(calc("=2^3^2")).toBe(64); // ^ is left-associative
    expect(calc("=10-4-3")).toBe(3);
    expect(calc("=50%")).toBe(0.5);
    expect(calc("=200*10%")).toBe(20);
    expect(calc('="a"&1+2')).toBe("a3");
  });

  it("compares like Excel", () => {
    expect(calc('="a"="A"')).toBe(true);
    expect(calc("=1<2")).toBe(true);
    expect(calc('=1<"a"')).toBe(true); // numbers sort before text
    expect(calc('="a"<TRUE')).toBe(true); // text before booleans
    expect(calc("=A1=0", book({ Sheet1: {} }))).toBe(true); // an empty cell equals 0
    expect(calc('=A1=""', book({ Sheet1: {} }))).toBe(true);
  });

  it("coerces text to numbers and fails with #VALUE!", () => {
    expect(calc('="3"+4')).toBe(7);
    expect(calc("=TRUE+1")).toBe(2);
    expect(code(calc('="x"+1'))).toBe("#VALUE!");
  });

  it("errors on division by zero and keeps the error flowing", () => {
    expect(code(calc("=1/0"))).toBe("#DIV/0!");
    expect(code(calc("=1/0+5"))).toBe("#DIV/0!");
    expect(code(calc("=SUM(1,1/0)"))).toBe("#DIV/0!");
    expect(code(calc("=0^0"))).toBe("#NUM!");
  });
});

describe("the functions our own templates use", () => {
  const ctx = book({
    Sheet1: {
      "3:1": 100, "4:1": 250, "5:1": 50, "6:1": 0, "7:1": 0, "8:1": 0,
      "3:2": "Done", "4:2": "Todo", "5:2": "Done", "6:2": "Done", "7:2": "Todo", "8:2": null,
      "9:1": 400, "9:2": 120,
    },
  });
  it("SUM, COUNTIF, COUNTA, IF", () => {
    expect(calc("=SUM(B4:B9)", ctx)).toBe(400);
    expect(calc('=COUNTIF(C4:C9,"Done")', ctx)).toBe(3);
    expect(calc("=COUNTA(C4:C9)", ctx)).toBe(5);
    expect(calc('=COUNTA(C4:C9)-COUNTIF(C4:C9,"Done")', ctx)).toBe(2);
    expect(calc("=IF(B10=0,0,C10/B10)", ctx)).toBe(0.3);
    expect(calc('=IF(COUNTA(C4:C9)=0,0,COUNTIF(C4:C9,"Done")/COUNTA(C4:C9))', ctx)).toBe(0.6);
  });
  it("DATE returns a serial number, not a JS Date", () => {
    expect(calc("=DATE(2020,1,1)")).toBe(43831);
    expect(calc("=DATE(2025,1,15)")).toBe(45672);
    expect(calc("=DATE(2020,14,1)")).toBe(calc("=DATE(2021,2,1)")); // months roll over
    expect(calc("=DATE(2020,3,0)")).toBe(calc("=DATE(2020,2,29)"));
  });
});

describe("errors as values (FortuneSheet gets these wrong)", () => {
  it("IFERROR catches any error", () => {
    expect(calc("=IFERROR(1/0,5)")).toBe(5); // FortuneSheet returned "DIV/0"
    expect(calc("=IFERROR(NA(),5)")).toBe(5);
    expect(calc("=IFERROR(7,5)")).toBe(7);
    expect(calc("=IFNA(NA(),5)")).toBe(5);
    expect(code(calc("=IFNA(1/0,5)"))).toBe("#DIV/0!");
  });
  it("IS* functions", () => {
    expect(calc("=ISERROR(1/0)")).toBe(true);
    expect(calc("=ISNA(NA())")).toBe(true);
    expect(calc("=ISERR(NA())")).toBe(false);
    expect(calc("=ISNUMBER(3)")).toBe(true);
    expect(calc('=ISTEXT("a")')).toBe(true);
    expect(calc("=ISBLANK(A1)")).toBe(true);
  });
  it("unknown functions and names are #NAME?", () => {
    expect(code(calc("=NOSUCH(1)"))).toBe("#NAME?");
    expect(code(calc("=missing_name+1"))).toBe("#NAME?");
  });
  it("only evaluates the branch IF picks", () => {
    expect(calc("=IF(TRUE,1,1/0)")).toBe(1);
    expect(code(calc("=IF(FALSE,1,1/0)"))).toBe("#DIV/0!");
    expect(calc("=IFS(FALSE,1,TRUE,2)")).toBe(2);
    expect(calc("=CHOOSE(2,10,20,30)")).toBe(20);
    expect(calc('=SWITCH(2,1,"a",2,"b","z")')).toBe("b");
    expect(calc("=AND(TRUE,1,5)")).toBe(true);
    expect(calc("=OR(FALSE,0)")).toBe(false);
  });
});

describe("references", () => {
  const sheets = {
    Sheet1: { "0:0": 1, "1:0": 2, "2:0": 3, "0:1": 10 },
    "My Sheet": { "0:0": 100 },
  };
  it("reads other sheets, whole columns and whole rows", () => {
    const ctx = book(sheets);
    expect(calc("=SUM(A:A)", ctx)).toBe(6);
    expect(calc("=SUM(1:1)", ctx)).toBe(11);
    expect(calc("='My Sheet'!A1+1", ctx)).toBe(101);
    expect(calc("=SUM(A1:A3)*B1", ctx)).toBe(60);
  });
  it("resolves defined names", () => {
    const ctx = book(sheets, { resolveName: (name) => (name === "Rate" ? 0.5 : undefined) });
    expect(calc("=Rate*10", ctx)).toBe(5);
  });
  it("evaluates array literals and array functions", () => {
    expect(calc("=SUM({1,2;3,4})")).toBe(10);
    expect(calc("=SUMPRODUCT({1,2,3},{4,5,6})")).toBe(32);
    expect(calc("=INDEX({1,2;3,4},2,1)")).toBe(3);
    expect(calc('=VLOOKUP(2,{1,"a";2,"b"},2,FALSE)')).toBe("b");
    expect(calc("=MATCH(3,{1;2;3},0)")).toBe(3);
  });
  it("refuses to expand a range past the limit instead of freezing", () => {
    const ctx = book(sheets, { maxRangeCells: 10 });
    expect(code(calc("=SUM(A1:Z100)", ctx))).toBe("#VALUE!");
  });
  it("ROW and COLUMN", () => {
    const ctx = book(sheets, { row: 4, col: 2 });
    expect(calc("=ROW()", ctx)).toBe(5);
    expect(calc("=COLUMN()", ctx)).toBe(3);
    expect(calc("=ROW(B7)", ctx)).toBe(7);
  });
});

describe("dates do not depend on the timezone", () => {
  const original = process.env.TZ;
  for (const zone of ["UTC", "Asia/Kolkata", "America/Los_Angeles", "Pacific/Auckland"]) {
    it(`in ${zone}`, () => {
      process.env.TZ = zone;
      expect(calc("=DATE(2020,1,1)")).toBe(43831);
      expect(calc("=EDATE(43831,1)")).toBe(43862);
      expect(calc("=EOMONTH(43831,0)")).toBe(43861);
      expect(calc("=YEAR(43831)")).toBe(2020);
      expect(calc("=MONTH(43862)")).toBe(2);
      expect(calc("=DAY(43861)")).toBe(31);
      expect(calc("=WEEKDAY(43831)")).toBe(4);
      expect(calc("=TODAY()", book({ Sheet1: {} }, { now: () => new Date(2024, 4, 17, 15, 30) }))).toBe(45429);
      process.env.TZ = original;
    });
  }
});

describe("rewriting references when rows or columns change", () => {
  const rewrite = (formula: string, edit: Parameters<typeof rewriteForEdit>[1], on = "Sheet1") =>
    print(rewriteForEdit(parse(formula), edit, on));
  const insertRow = (index: number, count = 1, sheet = "Sheet1") => ({ kind: "insert" as const, axis: "row" as const, sheet, index, count });
  const deleteRow = (index: number, count = 1, sheet = "Sheet1") => ({ kind: "delete" as const, axis: "row" as const, sheet, index, count });
  const deleteCol = (index: number, count = 1) => ({ kind: "delete" as const, axis: "col" as const, sheet: "Sheet1", index, count });

  it("inserting a row moves references at or below it", () => {
    expect(rewrite("=A5", insertRow(2))).toBe("=A6");
    expect(rewrite("=A2", insertRow(2))).toBe("=A2");
    expect(rewrite("=$A$5", insertRow(2))).toBe("=$A$6"); // absolute refs move too on a structural edit
  });
  it("inserting inside a range grows it", () => {
    expect(rewrite("=SUM(A1:A10)", insertRow(4, 2))).toBe("=SUM(A1:A12)");
    expect(rewrite("=SUM(A5:A10)", insertRow(0, 3))).toBe("=SUM(A8:A13)");
    expect(rewrite("=SUM(A:A)", insertRow(3))).toBe("=SUM(A:A)");
  });
  it("deleting a row gives #REF! for its cells and shrinks ranges", () => {
    expect(rewrite("=A3", deleteRow(2))).toBe("=#REF!");
    expect(rewrite("=A5", deleteRow(2))).toBe("=A4");
    expect(rewrite("=SUM(A1:A10)", deleteRow(4, 2))).toBe("=SUM(A1:A8)");
    expect(rewrite("=SUM(A3:A4)", deleteRow(2, 3))).toBe("=SUM(#REF!)");
    expect(rewrite("=SUM(A1:A3)", deleteRow(2, 5))).toBe("=SUM(A1:A2)");
  });
  it("works for columns", () => {
    expect(rewrite("=C1+D1", deleteCol(1))).toBe("=B1+C1");
    expect(rewrite("=B1", deleteCol(1))).toBe("=#REF!");
    expect(rewrite("=SUM(A1:E1)", deleteCol(1, 2))).toBe("=SUM(A1:C1)");
  });
  it("leaves other sheets alone and follows sheet names", () => {
    expect(rewrite("=Other!A5+A5", insertRow(0, 1, "Sheet1"))).toBe("=Other!A5+A6");
    expect(rewrite("=Sheet1!A5+A5", insertRow(0, 1, "Sheet1"), "Elsewhere")).toBe("=Sheet1!A6+A5");
  });
  it("shifts relative references on copy and keeps $absolute parts", () => {
    expect(print(shiftReferences(parse("=A1+$B$2+C$3+$D4"), 2, 1))).toBe("=B3+$B$2+D$3+$D6");
    expect(print(shiftReferences(parse("=SUM(A1:B2)"), 1, 0))).toBe("=SUM(A2:B3)");
    expect(print(shiftReferences(parse("=A1"), -1, 0))).toBe("=#REF!");
  });
  it("renames sheets in references", () => {
    expect(print(renameSheet(parse("=Old!A1+'Old'!B2+A1"), "Old", "New Name"))).toBe("='New Name'!A1+'New Name'!B2+A1");
  });
  it("lists what a formula depends on", () => {
    const refs = references(parse("=SUM(A1:A3)+Other!B2*IF(C1>0,D1,0)"));
    expect(refs.map((ref) => ref.kind)).toEqual(["range", "cell", "cell", "cell"]);
    expect(refs[1].sheet).toBe("Other");
  });
});

describe("classic function names still work", () => {
  it("STDEV, RANK, MODE and friends map to the modern implementations", () => {
    const close = (value: unknown, expected: number) => expect(Number(value)).toBeCloseTo(expected, 6);
    close(calc("=STDEV({1,2,3,4})"), 1.2909944487);
    close(calc("=STDEVP({1,2,3,4})"), 1.1180339887);
    close(calc("=VAR({1,2,3,4})"), 1.6666666667);
    close(calc("=VARP({1,2,3,4})"), 1.25);
    expect(calc("=RANK(3,{5,3,1})")).toBe(2);
    expect(calc("=MODE({1,2,2,3})")).toBe(2);
    close(calc("=PERCENTILE({1,2,3,4,5},0.5)"), 3);
    close(calc("=QUARTILE({1,2,3,4,5},1)"), 2);
    close(calc("=NORMSDIST(0)"), 0.5);
    close(calc("=NORMDIST(0,0,1,TRUE)"), 0.5);
    close(calc("=LOGNORMDIST(1,0,1)"), 0.5);
    close(calc("=CHIDIST(3,2)"), 0.2231301601);
    close(calc("=BINOMDIST(2,5,0.5,TRUE)"), 0.5);
    close(calc("=TINV(0.05,10)"), 2.2281388520);
    close(calc("=NORMSINV(0.975)"), 1.9599639845);
    close(calc("=NORMINV(0.5,10,2)"), 10);
    // One-tailed TDIST is half of the two-tailed value.
    close(Number(calc("=TDIST(2,10,2)")) / 2, Number(calc("=TDIST(2,10,1)")));
  });
});

describe("function coverage", () => {
  it("knows the functions people reach for", () => {
    const names = [
      "SUM", "AVERAGE", "COUNT", "COUNTA", "COUNTIF", "COUNTIFS", "SUMIF", "SUMIFS", "SUMPRODUCT", "MAX", "MIN", "ROUND",
      "IF", "IFS", "IFERROR", "AND", "OR", "NOT", "VLOOKUP", "HLOOKUP", "INDEX", "MATCH", "LOOKUP", "CONCAT", "CONCATENATE",
      "LEFT", "RIGHT", "MID", "LEN", "UPPER", "LOWER", "TRIM", "SUBSTITUTE", "TEXT", "DATE", "TODAY", "NOW", "YEAR", "MONTH",
      "DAY", "EDATE", "EOMONTH", "WEEKDAY", "NETWORKDAYS", "PMT", "FV", "NPV", "IRR", "STDEV", "MEDIAN", "RANK", "NORM.DIST",
    ];
    expect(names.filter((name) => !isKnownFunction(name))).toEqual([]);
  });
});

describe("functionNames", () => {
  it("lists what the engine can evaluate, sorted, with dotted names", async () => {
    const { functionNames, isKnownFunction } = await import("./index");
    const names = functionNames();
    expect(names.length).toBeGreaterThan(300);
    for (const name of ["SUM", "VLOOKUP", "IF", "IFERROR", "NORM.DIST", "SUMIFS"]) expect(names).toContain(name);
    expect([...names].sort()).toEqual(names);
    expect(names.every((name) => isKnownFunction(name))).toBe(true);
  });
});

describe("ADDRESS, HYPERLINK, LENB, FINDB and the Google conversions", () => {
  const run = (formula: string) => calculate(formula, { sheet: "Sheet1", getCell: () => undefined });
  it("builds addresses in every style", () => {
    expect(run("ADDRESS(2,3)")).toBe("$C$2");
    expect(run("ADDRESS(2,3,4)")).toBe("C2");
    expect(run("ADDRESS(2,3,2)")).toBe("C$2");
    expect(run("ADDRESS(2,3,3)")).toBe("$C2");
    expect(run("ADDRESS(2,28,4)")).toBe("AB2");
    expect(run("ADDRESS(2,3,1,FALSE)")).toBe("R2C3");
    expect(run('ADDRESS(1,1,1,TRUE,"Data")')).toBe("Data!$A$1");
    expect(run('ADDRESS(1,1,4,TRUE,"My Sheet")')).toBe("'My Sheet'!A1");
    expect(run("ADDRESS(0,1)")).toBeInstanceOf(FormulaError);
  });
  it("covers the small helpers", () => {
    expect(run('HYPERLINK("https://a.test","Go")')).toBe("Go");
    expect(run('HYPERLINK("https://a.test")')).toBe("https://a.test");
    expect(run('LENB("abc")')).toBe(3);
    expect(run('FINDB("c","abcabc",4)')).toBe(6);
    expect(run('FINDB("x","abc")')).toBeInstanceOf(FormulaError);
    expect(run("TO_TEXT(12)")).toBe("12");
    expect(run("TO_PERCENT(0.5)")).toBe(0.5);
    expect(run("UNARY_PERCENT(50)")).toBe(0.5);
    expect(run("TO_DATE(45000.7)")).toBe(45000);
    expect(run("TO_PURE_NUMBER(7)")).toBe(7);
  });
});

describe("number formats shown in big sheets", () => {
  it("formats currency tokens, scientific, time and dates", async () => {
    const { formatValue } = await import("./index");
    expect(formatValue(1234.5, "[$€-407] #,##0.00")).toBe("€ 1,234.50");
    expect(formatValue(1234.5, "[$£-809]#,##0.00")).toBe("£1,234.50");
    expect(formatValue(1234.5, "[$CHF-807] #,##0.00")).toBe("CHF 1,234.50");
    expect(formatValue(1234.5, "0.00E+00")).toBe("1.23E+03");
    expect(formatValue(45000.75, "h:mm:ss AM/PM")).toBe("6:00:00 PM");
    expect(formatValue(45000.75, "yyyy-mm-dd hh:mm:ss")).toBe("2023-03-15 18:00:00");
    expect(formatValue(45000, "mmmm d, yyyy")).toBe("March 15, 2023");
    expect(formatValue(-2, "#,##0.00;[Red]-#,##0.00")).toBe("-2.00");
  });
});

describe("INDIRECT, OFFSET, CELL, MDETERM, MINVERSE", () => {
  const data = book({ Sheet1: { "0:0": 10, "1:0": 20, "2:0": 30, "0:1": 1, "1:1": 2 }, Other: { "0:0": 99 } });
  it("INDIRECT follows text to a cell, range or other sheet", () => {
    expect(calc('=INDIRECT("A2")', data)).toBe(20);
    expect(calc('=SUM(INDIRECT("A1:A3"))', data)).toBe(60);
    expect(calc('=INDIRECT("Other!A1")', data)).toBe(99);
    expect(calc('=INDIRECT("A"&2)', data)).toBe(20);
    expect(String(calc('=INDIRECT("nonsense here")', data))).toBe("#REF!");
  });
  it("OFFSET moves and resizes a reference", () => {
    expect(calc("=OFFSET(A1,1,0)", data)).toBe(20);
    expect(calc("=SUM(OFFSET(A1,0,0,3,1))", data)).toBe(60);
    expect(calc("=SUM(OFFSET(A1:A2,1,0))", data)).toBe(50);
    expect(String(calc("=OFFSET(A1,-1,0)", data))).toBe("#REF!");
  });
  it("CELL reports address, row, column, contents and type", () => {
    expect(calc('=CELL("address",B2)', data)).toBe("$B$2");
    expect(calc('=CELL("row",B2)', data)).toBe(2);
    expect(calc('=CELL("col",B2)', data)).toBe(2);
    expect(calc('=CELL("contents",A3)', data)).toBe(30);
    expect(calc('=CELL("type",A3)', data)).toBe("v");
    expect(calc('=CELL("type",D9)', data)).toBe("b");
  });
  it("MDETERM and MINVERSE work on square matrices", () => {
    expect(calc("=MDETERM({1,2;3,4})")).toBe(-2);
    expect(calc("=MDETERM({2,0;0,3})")).toBe(6);
    expect(calc("=INDEX(MINVERSE({4,7;2,6}),1,1)")).toBeCloseTo(0.6, 9);
    expect(calc("=INDEX(MINVERSE({4,7;2,6}),1,2)")).toBeCloseTo(-0.7, 9);
    expect(String(calc("=MINVERSE({1,2;2,4})"))).toBe("#NUM!");
    expect(String(calc("=MDETERM({1,2,3})"))).toBe("#VALUE!");
  });
});
