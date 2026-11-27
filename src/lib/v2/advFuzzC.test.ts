// Differential fuzzing, family C: structure-heavy (inserts/deletes of rows and columns, blocks covering whole chunks,
// 100 000 row inserts, deleting rows that formulas read or sit in), with saves and reopens in between.
import { describe, expect, it } from "vitest";
import { DEFAULT_MIX, fuzz, report, type Mix } from "./advFuzzHarness";

const MIX: Mix = { ...DEFAULT_MIX, struct: 70, type: 14, sort: 1, drag: 2, fill: 2, paste: 3, move: 2, replace: 0, layout: 3, tabs: 1, undo: 0, save: 10, recalc: 2, nasty: 0.35 };

describe("adversarial differential fuzz C (structure heavy, nasty sizes)", () => {
  for (const seed of [301, 302, 303, 304, 305, 306, 307, 308, 309, 310, 311, 312]) {
    it(`seed ${seed}`, async () => {
      const r = await fuzz(seed, 45, { tabs: 2, rows: 2500, mix: MIX }, (s) => console.log(s));
      if (r) console.log(report(seed, r));
      expect(r === null ? "ok" : report(seed, r)).toBe("ok");
    }, 500_000);
  }
});
