// Differential fuzzing, family G: undo and redo mixed with edits, structure edits, saves and tab changes.
// BUG: this family currently finds differences that depend on which chunks each store happens to keep in memory
// (BigSheetStore.apply drops an undo/redo change for a chunk that is not in memory; restoreSnapshot drops chunks read since).
// Known minimal-ish sequence (4 tabs x 1000 rows): see the first test below.
import { describe, expect, it } from "vitest";
import { DEFAULT_MIX, fuzz, replay, report, world, type Mix, type Op } from "./advFuzzHarness";

const MIX: Mix = { ...DEFAULT_MIX, struct: 20, type: 25, tabs: 6, undo: 25, save: 6, sort: 2, drag: 3, fill: 2, paste: 6, move: 3, replace: 1, layout: 2, recalc: 2, nasty: 0.05 };

describe("adversarial differential fuzz G (undo and redo)", () => {
  // BUG: after these 9 operations the reference has undone the typed formula (cell shows the original text) and the id store still has it.
  // Both stores run the same undo calls; they differ in which chunks are cached (the id store pins every chunk that holds formulas).
  it("known failing sequence: structural undo followed by cell undo after an automatic save", async () => {
    const w = await world(4, 1000);
    const ops: Op[] = [
      { k: "layout", tab: 4, patch: { columnWidths: { "5": 200 } } },
      { k: "paste", tab: 1, r1: 89, c1: 0, r2: 100, c2: 3, toTab: 2, row: 437, col: 9 },
      { k: "type", tab: 1, row: 593, col: 1, text: "=SUM('Tab 1'!H947:K974)" },
      { k: "struct", kind: "insert", axis: "col", tab: 0, index: 11, count: 2 },
      { k: "drag", tab: 0, r1: 588, c1: 10, r2: 588, c2: 12, grow: 36, dir: "down" },
      { k: "rename", tab: 4, name: "Tab 3" },
      { k: "recalc" },
      { k: "undo" },
      { k: "undo" },
    ];
    const f = await replay(w, ops);
    expect(f ? f.why : "ok").toBe("ok");
  }, 300_000);

  for (const seed of [501, 502, 503, 504]) {
    it(`seed ${seed}`, async () => {
      const r = await fuzz(seed, 40, { tabs: 3, rows: 1000, mix: MIX }, (s) => console.log(s));
      if (r) console.log(report(seed, r));
      expect(r === null ? "ok" : report(seed, r)).toBe("ok");
    }, 500_000);
  }
});
