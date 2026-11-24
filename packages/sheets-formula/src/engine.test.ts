import { describe, expect, it } from "vitest";
import { Engine, FormulaError } from "./index";

function setup() {
  const constants = new Map<string, number | string>();
  const key = (s: string, r: number, c: number) => `${s.toLowerCase()}|${r}|${c}`;
  const engine = new Engine({
    getConstant: (s, r, c) => constants.get(key(s, r, c)) ?? null,
    extent: () => ({ rows: 100, cols: 10 }),
  });
  const put = (s: string, r: number, c: number, v: number | string) => {
    constants.set(key(s, r, c), v);
    engine.notifyChanged(s, r, c);
  };
  return { engine, put };
}

describe("Engine", () => {
  it("computes a formula and updates it when an input changes", () => {
    const { engine, put } = setup();
    put("S", 0, 0, 2);
    engine.setFormula("S", 0, 1, "=A1*10");
    expect(engine.value("S", 0, 1)).toBe(20);
    put("S", 0, 0, 5);
    expect(engine.recalculate()).toEqual([{ sheet: "S", row: 0, col: 1, value: 50 }]);
  });

  it("recomputes chains in order and only reports real changes", () => {
    const { engine, put } = setup();
    put("S", 0, 0, 1);
    engine.setFormula("S", 1, 0, "=A1+1");
    engine.setFormula("S", 2, 0, "=A2+1");
    engine.setFormula("S", 3, 0, "=MIN(A1,1)"); // stays 1 when A1 grows
    engine.recalculate();
    put("S", 0, 0, 10);
    const changes = engine.recalculate().map((c) => [c.row, c.value]);
    expect(changes.sort()).toEqual([[1, 11], [2, 12]]);
  });

  it("handles ranges, including cells added to a range later", () => {
    const { engine, put } = setup();
    put("S", 0, 0, 1);
    put("S", 1, 0, 2);
    engine.setFormula("S", 0, 2, "=SUM(A1:A5)");
    expect(engine.value("S", 0, 2)).toBe(3);
    put("S", 4, 0, 10);
    expect(engine.value("S", 0, 2)).toBe(13);
    engine.setFormula("S", 2, 0, "=A1+A2"); // a formula inside the range
    expect(engine.value("S", 0, 2)).toBe(16);
  });

  it("whole-column references see later edits", () => {
    const { engine, put } = setup();
    engine.setFormula("S", 0, 1, "=SUM(A:A)");
    put("S", 7, 0, 4);
    expect(engine.value("S", 0, 1)).toBe(4);
  });

  it("works across sheets and follows a sheet rename", () => {
    const { engine, put } = setup();
    put("Data", 0, 0, 3);
    engine.setFormula("Calc", 0, 0, "=Data!A1*2");
    expect(engine.value("Calc", 0, 0)).toBe(6);
    engine.renameSheet("Data", "Numbers");
    expect(engine.formulaText("Calc", 0, 0)).toBe("=Numbers!A1*2");
  });

  it("marks cycles #CIRC! and still computes cells hanging off them", () => {
    const { engine, put } = setup();
    engine.setFormula("S", 0, 0, "=B1+1");
    engine.setFormula("S", 0, 1, "=A1+1");
    engine.setFormula("S", 0, 2, "=1+1");
    engine.setFormula("S", 0, 3, "=ISERROR(A1)");
    put("S", 5, 5, 1);
    expect(engine.value("S", 0, 0)).toEqual(new FormulaError("#CIRC!"));
    expect(engine.value("S", 0, 1)).toEqual(new FormulaError("#CIRC!"));
    expect(engine.value("S", 0, 2)).toBe(2);
    expect(engine.value("S", 0, 3)).toBe(true);
  });

  it("a self-reference is circular and breaking the loop recovers", () => {
    const { engine } = setup();
    engine.setFormula("S", 0, 0, "=A1");
    expect(engine.value("S", 0, 0)).toEqual(new FormulaError("#CIRC!"));
    engine.setFormula("S", 0, 0, "=5");
    expect(engine.value("S", 0, 0)).toBe(5);
  });

  it("survives a very long chain without overflowing the stack", () => {
    const { engine, put } = setup();
    put("S", 0, 0, 1);
    const n = 20000;
    for (let r = 1; r <= n; r++) engine.setFormula("S", r, 0, `=A${r}+1`);
    expect(engine.value("S", n, 0)).toBe(n + 1);
    put("S", 0, 0, 2);
    expect(engine.value("S", n, 0)).toBe(n + 2);
  });

  it("inserting a row rewrites formulas and keeps results", () => {
    const { engine, put } = setup();
    put("S", 0, 0, 1);
    put("S", 1, 0, 2);
    engine.setFormula("S", 2, 0, "=SUM(A1:A2)");
    engine.applyEdit({ kind: "insert", axis: "row", sheet: "S", index: 1, count: 1 });
    expect(engine.formulaText("S", 3, 0)).toBe("=SUM(A1:A3)");
    expect(engine.hasFormula("S", 2, 0)).toBe(false);
  });

  it("deleting a referenced row gives #REF! and deleting the formula's row removes it", () => {
    const { engine, put } = setup();
    put("S", 1, 0, 2);
    engine.setFormula("S", 0, 1, "=A2*2");
    engine.setFormula("S", 1, 1, "=1");
    engine.applyEdit({ kind: "delete", axis: "row", sheet: "S", index: 1, count: 1 });
    expect(engine.value("S", 0, 1)).toEqual(new FormulaError("#REF!"));
    expect(engine.formulaCount).toBe(1);
  });

  it("clearing a formula lets typed values show through to dependents", () => {
    const { engine, put } = setup();
    engine.setFormula("S", 0, 0, "=7");
    engine.setFormula("S", 0, 1, "=A1+1");
    expect(engine.value("S", 0, 1)).toBe(8);
    engine.clearFormula("S", 0, 0);
    put("S", 0, 0, 1);
    expect(engine.value("S", 0, 1)).toBe(2);
  });

  it("volatile formulas recompute on demand", () => {
    let t = new Date(Date.UTC(2024, 0, 1));
    const engine = new Engine({ getConstant: () => null, now: () => t });
    engine.setFormula("S", 0, 0, "=TODAY()");
    const first = engine.value("S", 0, 0);
    t = new Date(Date.UTC(2024, 0, 3));
    engine.recalculateVolatile();
    expect(engine.value("S", 0, 0)).toBe((first as number) + 2);
  });

  it("keeps unparsable text as #NAME?", () => {
    const { engine } = setup();
    engine.setFormula("S", 0, 0, "=1+");
    expect(engine.value("S", 0, 0)).toEqual(new FormulaError("#NAME?"));
  });
});
