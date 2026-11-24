// Differential fuzzing, family D: tab operations (add, rename with formula rewrite, duplicate after a save, reorder, delete)
// mixed with undo/redo, structure edits and cross-tab formulas; 4 tabs.
import { describe, expect, it } from "vitest";
import { DEFAULT_MIX, fuzz, report, type Mix } from "./advFuzzHarness";

const MIX: Mix = { ...DEFAULT_MIX, struct: 25, type: 20, tabs: 25, undo: 0, save: 12, sort: 1, drag: 1, fill: 1, paste: 3, move: 1, replace: 1, layout: 2, recalc: 3, nasty: 0.1 };

describe("adversarial differential fuzz D (tabs, undo, cross-tab references)", () => {
  for (const seed of [401, 402, 403, 404, 405, 406, 407, 408, 409, 410]) {
    it(`seed ${seed}`, async () => {
      const r = await fuzz(seed, 40, { tabs: 4, rows: 1000, mix: MIX }, (s) => console.log(s));
      if (r) console.log(report(seed, r));
      expect(r === null ? "ok" : report(seed, r)).toBe("ok");
    }, 500_000);
  }
});
