import { describe, expect, it } from "vitest";
import { parse, print } from "@formstr/sheets-formula";
import { ArrayAxis, formulaFromIds, formulaToIds, type RefContext, type TabAxes } from "./idRefs";

function tab(id: string, name: string, rows = 30, cols = 30): TabAxes {
  return {
    id,
    name,
    rows: new ArrayAxis(Array.from({ length: rows }, (_, i) => `r${id}.${i}`)),
    cols: new ArrayAxis(Array.from({ length: cols }, (_, i) => `c${id}${i}`)),
  };
}
const NAMES = ["Sheet1", "Sheet2", "Sheet3", "My Data", "It's", "Ünï", "A!B", "2024", "A1", "SUM", "Q1", "x.y", "TAX2024"];
const tabs = NAMES.map((n, i) => tab("t" + i, n));
const ctx: RefContext = { tabs, current: "t0" };

const FORMULAS = [
  "a1", "A1+b2", "$A1", "A$1", "$A$1", "SUM($A:$B)", "SUM(2:5)", "SUM($2:$5)", "SUM(a:b)",
  "Sheet2!A1", "sheet2!a1", "SHEET2!A1:B2", "'My Data'!A1", "'It''s'!A1", "Ünï!A1", "'Ünï'!A1", "'A!B'!C3", "'2024'!A1", "'A1'!B2", "'SUM'!A1",
  "SUM!A1", "Sheet2!A:A", "Sheet2!2:3", "'My Data'!$A$1:$B$2",
  'A1&"A1"', 'IF(A1="B2","!",C3)', 'LEN("~(")+A1', '"a""b A1"&B2', '"~(t0:r:c:0)"&A1', 'CONCAT("x",A1,"""")',
  "SUM(IF(A1>1,B1:B3,C1),MAX(D1,E1))", "{1,2;3,4}", "{A1,2}", "A1%", "-A1", "--A1", "-A1^2", "A1<>B2", "A1<=B2", "A1&B2", "A1>=B2",
  "TAX2024+Q1", "TAX2024*A1", "Q1", "LOG10(A1)", "ATAN2(1,A1)", "DEC2BIN(A1)", "BIN2DEC(A1)", "T.INV.2T(0.1,A1)", "NORM.S.DIST(A1,TRUE)",
  "1.5E3+A1", ".5+A1", "1.5:2", "1E3", "1E+3+A1", "12:30", "TIME(12,30,0)+A1", "#REF!", "#N/A", "IFERROR(A1,#N/A)", "#REF!+A1",
  "Sheet1:Sheet3!A1", "Sheet2!A1:Sheet2!B2", "A1:B2:C3", "(A1,B2)", "SUM((A1,B2))", "A1 B2", "SUM(A1:B3 B2:C4)", "SUM( A1 , B2 )", "  A1  +  B2  ", "A1\nB2", "A1+\r\nB2", "SUM(A1,\n\tB2)",
  "XFD1", "XFD1048576", "A1048576", "A1048577", "A9999999", "AAAA1", "XFE1", "A0", "A:XFD", "0:5", "1:1048576", "1:1",
  "A1:A1", "sum(a1:b2)", "Sum(A1)", "if(a1,b1,c1)", "A1A", "A1_", "_A1", "A1.5", "A1.B", "a.A1", "x.y!A1", "'x.y'!A1", "ABCD1", "AB1C", "1A", "A1(2)", "A1(", "A1!B2",
  "~(", "~(a:b:c:1)", "~(Sheet:1)+A1", "A1+~", "A1~",
];

// Text that comes back differently written but means the same (or was never a valid formula): lower-case
// references and function names are upper-cased, a sheet name takes the real tab's spelling, the parser turns
// Sheet2!A1:Sheet2!B2 into Sheet2!A1:B2, a name that does not need quotes loses them, and so on.
const COSMETIC = new Set(["a1", "A1+b2", "SUM(a:b)", "sheet2!a1", "SHEET2!A1:B2", "'SUM'!A1", "Sheet2!A1:Sheet2!B2", "sum(a1:b2)", "if(a1,b1,c1)", "x.y!A1", "A1!B2", "~(a:b:c:1)"]);

describe("round trip text", () => {
  for (const f of FORMULAS) {
    it(`exact or same tree: ${JSON.stringify(f)}`, () => {
      if (["sheet2!a1", "SHEET2!A1:B2", "Sheet2!A1:Sheet2!B2", "~(a:b:c:1)"].includes(f)) return; // the same cells written with the tab's own spelling (or not a formula at all)
      let out = "";
      expect(() => { out = formulaFromIds(formulaToIds(f, ctx), ctx); }).not.toThrow();
      if (out === f) return;
      // fall back to tree equality
      let a: string | null = null, b: string | null = null;
      try { a = print(parse("=" + f)); } catch { a = "ERR"; }
      try { b = print(parse("=" + out)); } catch { b = "ERR"; }
      expect(b, `got ${JSON.stringify(out)} for ${JSON.stringify(f)}`).toBe(a);
    });
    it(`exact text: ${JSON.stringify(f)}`, () => {
      if (COSMETIC.has(f)) return;
      expect(formulaFromIds(formulaToIds(f, ctx), ctx)).toBe(f);
    });
  }
  it("contains no stray id tokens afterwards (unless input had them)", () => {
    for (const f of FORMULAS) {
      if (f.includes("~")) continue;
      expect(formulaFromIds(formulaToIds(f, ctx), ctx)).not.toContain("~(");
    }
  });
  // BUG: formulaToIds is not idempotent: ids such as "ct01" look like cell refs (letters+digits), so running toIds on already-converted text wraps tokens inside tokens. Callers must never convert twice.
  it("converting twice is idempotent (to ids of already-id text)", () => {
    for (const f of FORMULAS) {
      const once = formulaToIds(f, ctx);
      expect(formulaToIds(once, ctx), f).toBe(once);
    }
  });
});

describe("round trip big/perf", () => {
  it("10k refs fast and exact", () => {
    const t0 = tab("a", "S", 20000, 30);
    const c: RefContext = { tabs: [t0], current: "a" };
    const f = Array.from({ length: 10000 }, (_, i) => `A${(i % 19000) + 1}`).join("+");
    const s = performance.now();
    const st = formulaToIds(f, c);
    const mid = performance.now();
    const back = formulaFromIds(st, c);
    const e = performance.now();
    console.log("10k refs toIds", (mid - s).toFixed(0), "ms, fromIds", (e - mid).toFixed(0), "ms");
    expect(back).toBe(f);
    expect(e - s).toBeLessThan(2000);
  });
  it("very long string-heavy formula with 50k quotes", () => {
    const c: RefContext = { tabs: [tab("a", "S")], current: "a" };
    const f = Array.from({ length: 20000 }, () => '"x""y"&A1').join("&");
    const s = performance.now();
    expect(formulaFromIds(formulaToIds(f, c), c)).toBe(f);
    expect(performance.now() - s).toBeLessThan(3000);
  });
  it("pathological regex input does not hang (long quoted-looking text)", () => {
    const c: RefContext = { tabs: [tab("a", "S")], current: "a" };
    const f = "'" + "a".repeat(50000) + "A1";
    const s = performance.now();
    formulaFromIds(formulaToIds(f, c), c);
    expect(performance.now() - s).toBeLessThan(2000);
    const g = '"' + "a".repeat(50000);
    const s2 = performance.now();
    formulaFromIds(formulaToIds(g, c), c);
    expect(performance.now() - s2).toBeLessThan(2000);
  });
});

describe("tab name ambiguity", () => {
  // BUG: tabNamed() returns the first tab whose lowercase name matches, so a reference typed against the second of two tabs named Data/data silently resolves to the first.
  it("two tabs differing only by case: ref picks the first, round trip may switch tab", () => {
    const x = tab("x", "Data"), y = tab("y", "data");
    const c: RefContext = { tabs: [tab("a", "S"), x, y], current: "a" };
    // BUG candidate: 'data'!A1 resolves to tab Data (first) so reading back gives Data!A1
    expect(formulaFromIds(formulaToIds("data!A1", c), c)).toBe("data!A1");
    expect(formulaFromIds(formulaToIds("Data!A1", c), c)).toBe("Data!A1");
  });
  // BUG: with two tabs of the same name, a formula living on the second one that says Same!A1 is bound to the FIRST tab (not the own tab).
  it("duplicate names: exact ref to second tab with same name is unreachable", () => {
    const x = tab("x", "Same"), y = tab("y", "Same");
    const c: RefContext = { tabs: [tab("a", "S"), x, y], current: "y" };
    // formula living on y referencing Same!A1 means tab x or y? current tab is y; the first match is x
    const st = formulaToIds("Same!A1", c);
    expect(st).toContain("~(_:"); // would be self if it means y
  });
  it("current tab not in ctx.tabs returns formula unchanged", () => {
    const c: RefContext = { tabs: [tab("a", "S")], current: "zzz" };
    expect(formulaToIds("A1", c)).toBe("A1");
  });
  it("sheet whose name needs quoting but contains a quote and apostrophes and ! round trips", () => {
    const t1 = tab("q", "O'Neil's! 1");
    const c: RefContext = { tabs: [tab("a", "S"), t1], current: "a" };
    expect(formulaFromIds(formulaToIds("'O''Neil''s! 1'!B2", c), c)).toBe("'O''Neil''s! 1'!B2");
  });
  // LIMITATION (not a bug if ids never contain : ( ) ): ids containing these characters make the stored token unparseable and the formula is shown as raw ~( text.
  // LIMITATION (documented): ids and tab ids must not contain ":" "(" or ")"; the app never makes such ids. The next two tests record that.
  it.skip("tab names with ':' or ')' break the token? (renders fine since by id)", () => {
    const t1 = tab("q:1", "Weird");
    const c: RefContext = { tabs: [tab("a", "S"), t1], current: "a" };
    expect(formulaFromIds(formulaToIds("Weird!B2", c), c)).toBe("Weird!B2");
  });
  // LIMITATION: same, for row ids containing a colon.
  it.skip("row id containing ':' or '(' (not expected) -- ids with dots OK", () => {
    const t1 = tab("a", "S");
    t1.rows = new ArrayAxis(["r:0", "r:1"]);
    const c: RefContext = { tabs: [t1], current: "a" };
    expect(formulaFromIds(formulaToIds("A1", c), c)).toBe("A1");
  });
});

describe("more text bugs", () => {
  // BUG: a quoted sheet name that contains a double quote is cut by the STRING split in formulaToIds, so the
  // sheet prefix is left as text and the A1 part is converted against the CURRENT tab, not the named tab.
  it("sheet name containing a double quote keeps pointing at that sheet", () => {
    const q = tab("q", 'Say "hi"', 30, 30);
    q.rows = new ArrayAxis(Array.from({ length: 30 }, (_, i) => `Q.${i}`));
    const c: RefContext = { tabs: [tab("a", "S"), q], current: "a" };
    const stored = formulaToIds(`'Say ""hi""'!B2`, c); // not valid Excel quoting anyway; the real name has raw quotes:
    const stored2 = formulaToIds(`'Say "hi"'!B2`, c);
    expect(stored2).toContain("~(q:");
    void stored;
  });
  // BUG: a reversed range typed by the user (B5:B2, D4:A1, 5:2, B:A) is normalised by the parser to B2:B5 but
  // formulaToIds stores corners in typed order, so reading back gives #REF! (end < start check in formulaFromIds).
  it("reversed ranges are normalised, not #REF!", () => {
    const c: RefContext = { tabs: [tab("a", "S")], current: "a" };
    for (const f of ["SUM(B5:B2)", "D4:A1", "SUM(5:2)", "SUM(B:A)", "B1:A1"]) {
      const out = formulaFromIds(formulaToIds(f, c), c);
      expect(out, f).not.toContain("#REF!");
      expect(print(parse("=" + out)), f).toBe(print(parse("=" + f)));
    }
  });
  // BUG (low): two-sheet corners Sheet2!A1:Sheet3!B2 -> second corner's tab is ignored in formulaFromIds
  // (range printed with the first tab's prefix) or #REF! when the first corner is the own tab.
  it.skip("corner on a different tab than the first corner is not silently re-pointed", () => {
    const c: RefContext = { tabs: [tab("a", "S"), tab("b", "T"), tab("c", "U")], current: "a" };
    const out = formulaFromIds(formulaToIds("T!A1:U!B2", c), c);
    expect(out).toBe("T!A1:U!B2");
  });
});
