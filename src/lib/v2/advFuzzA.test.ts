// Differential fuzzing, family A: every kind of operation mixed (structure, typing, sort, fill, paste, move, find/replace,
// layout, tabs, undo/redo, save/reopen, recalculation). Reference = positional formulas, subject = formulas by id.
// 3 tabs of 1500 rows = about 100 chunks, the cache holds 6.
import { describe, expect, it } from "vitest";
import { fuzz, report } from "./advFuzzHarness";

describe("adversarial differential fuzz A (all operation kinds)", () => {
  for (const seed of [101, 102, 103, 104, 105, 106, 107, 108, 109, 110, 111, 112]) {
    it(`seed ${seed}`, async () => {
      const r = await fuzz(seed, 40, { tabs: 3, rows: 1500 }, (s) => console.log(s));
      if (r) console.log(report(seed, r));
      expect(r === null ? "ok" : report(seed, r)).toBe("ok");
    }, 500_000);
  }
});
