import { describe, expect, it } from "vitest";
import { ArrayAxis, RunAxis, formulaFromIds, formulaToIds, type RefContext } from "./idRefs";
import type { RowRun } from "../workbook/ids";

const cols = () => new ArrayAxis(Array.from({ length: 30 }, (_, i) => `c${i}`));

function bench(label: string, runs: RowRun[], rowCount: number, n = 100_000) {
  const ctx: RefContext = { tabs: [{ id: "a", name: "S", rows: new RunAxis(runs), cols: cols() }], current: "a" };
  const formulas = Array.from({ length: n }, (_, i) => `SUM(B${(i * 7) % rowCount + 1}:B${((i * 7) % rowCount) + 3})+C${(i * 13) % rowCount + 1}*2`);
  const t0 = performance.now();
  const stored = formulas.map((f) => formulaToIds(f, ctx));
  const t1 = performance.now();
  const out = stored.map((s) => formulaFromIds(s, ctx));
  const t2 = performance.now();
  console.log(`${label}: ${n} formulas toIds ${(t1 - t0).toFixed(0)} ms, fromIds ${(t2 - t1).toFixed(0)} ms`);
  expect(out[12345]).toBe(formulas[12345]);
  return { to: t1 - t0, from: t2 - t1 };
}

describe("performance", () => {
  it("100k formulas, one million-row run", () => {
    const r = bench("1 run", [["k3f9", 0, 1_000_000]], 1_000_000);
    expect(r.to).toBeLessThan(5000);
    expect(r.from).toBeLessThan(5000);
  });
  // BUG (perf): RunAxis.idAt is a linear scan over the runs, so each reference costs O(runs).
  it("100k formulas, 5000 runs (spread over the whole sheet)", () => {
    const runs: RowRun[] = Array.from({ length: 5000 }, (_, i) => [`t${i % 50}`, i * 1000, 200] as RowRun);
    const r = bench("5000 runs", runs, 5000 * 200);
    expect(r.to).toBeLessThan(5000);
    expect(r.from).toBeLessThan(5000);
  });
  it("100k formulas, ArrayAxis of 100k rows", () => {
    const ctx: RefContext = {
      tabs: [{ id: "a", name: "S", rows: new ArrayAxis(Array.from({ length: 100_000 }, (_, i) => `r${i}`)), cols: cols() }],
      current: "a",
    };
    const formulas = Array.from({ length: 100_000 }, (_, i) => `SUM(B${(i % 99990) + 1}:B${(i % 99990) + 5})`);
    const t0 = performance.now();
    const s = formulas.map((f) => formulaToIds(f, ctx));
    const o = s.map((x) => formulaFromIds(x, ctx));
    console.log(`array axis: ${(performance.now() - t0).toFixed(0)} ms`);
    expect(o[77]).toBe(formulas[77]);
    expect(performance.now() - t0).toBeLessThan(5000);
  });
  it("RunAxis.length is O(runs) per access (idAt/length in a loop)", () => {
    const runs: RowRun[] = Array.from({ length: 5000 }, (_, i) => [`t${i}`, 0, 10] as RowRun);
    const a = new RunAxis(runs);
    const t0 = performance.now();
    let sum = 0;
    for (let i = 0; i < 20000; i++) sum += a.length;
    console.log(`20000 x length with 5000 runs: ${(performance.now() - t0).toFixed(0)} ms`, sum);
    expect(performance.now() - t0).toBeLessThan(1000);
  });
});
