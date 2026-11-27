// Deliberately nasty, scripted sequences (run through the same two-store comparison as the fuzzers), plus two stand-alone
// checks that found bugs which exist in both modes (so they cannot show up as a difference between the stores).
import { describe, expect, it } from "vitest";
import { FormulaIndex } from "./bigFormulas";
import { editStructure } from "./bigStructure";
import { apply, describeOp, dump, openSide, replay, saveSide, typeCell, world, type Op } from "./advFuzzHarness";

const save = (reopen = true): Op => ({ k: "save", reopen });
const rows = (kind: "insert" | "delete", tab: number, index: number, count: number): Op => ({ k: "struct", kind, axis: "row", tab, index, count });

async function expectSame(tabs: number, nrows: number, ops: Op[]): Promise<void> {
  const w = await world(tabs, nrows);
  const f = await replay(w, ops);
  expect(f ? `${f.why}\n at op ${f.at}: ${f.at < ops.length ? describeOp(ops[f.at]) : "end"}` : "ok").toBe("ok");
}

describe("adversarial scripted cases", () => {
  it("insert 100 000 rows in the middle of a tab, save, delete across the new rows, recalc", async () => {
    await expectSame(2, 1500, [rows("insert", 0, 700, 100_000), save(), rows("delete", 0, 650, 70), save(false), rows("delete", 0, 99_000, 1500), { k: "recalc" }, save()]);
  }, 400_000);

  it("delete a block that covers entire chunks (with formulas inside and formulas elsewhere that read them)", async () => {
    await expectSame(2, 1500, [rows("delete", 0, 100, 800), { k: "recalc" }, save(), rows("insert", 0, 50, 7), rows("delete", 1, 300, 600), save(false), { k: "undo" }, save()]);
  }, 400_000);

  it("delete the rows that hold formulas other formulas read (anchors), then undo", async () => {
    await expectSame(3, 1500, [rows("delete", 0, 495, 6), rows("delete", 1, 766, 8), save(false), rows("delete", 2, 1394, 6), { k: "recalc" }, { k: "undo" }, save(), rows("insert", 0, 490, 3), save()]);
  }, 400_000);

  it("formulas that point past the last row (A9999, ranges to row 5000), then inserts and deletes before and at the end", async () => {
    await expectSame(2, 1500, [
      { k: "type", tab: 0, row: 20, col: 12, text: "=A9999+1" },
      { k: "type", tab: 0, row: 21, col: 12, text: "=SUM(A1:A5000)" },
      { k: "type", tab: 1, row: 22, col: 12, text: "=SUM('Tab 1'!B10:B3000)" },
      save(),
      rows("insert", 0, 10, 5),
      rows("insert", 0, 1505, 4),
      rows("delete", 0, 1400, 100),
      save(),
      rows("delete", 0, 2, 30),
      save(),
    ]);
  }, 400_000);

  it("duplicate a tab after structure edits, then edit the copy and the original differently", async () => {
    await expectSame(2, 1200, [
      rows("insert", 0, 300, 10),
      save(),
      { k: "dup", tab: 0, name: "Copy A" },
      rows("delete", 1, 100, 20),
      rows("insert", 0, 50, 3),
      save(false),
      { k: "type", tab: 1, row: 105, col: 12, text: "=SUM(A1:A400)+'Tab 2'!A4" },
      rows("delete", 0, 10, 15),
      { k: "recalc" },
      save(),
    ]);
  }, 400_000);

  it("rename a tab that other tabs read, then insert and delete rows in it and reorder tabs", async () => {
    await expectSame(3, 1200, [
      { k: "rename", tab: 1, name: "Summary" },
      rows("insert", 1, 40, 12),
      { k: "remap", order: [2, 0, 1] },
      rows("delete", 2, 200, 40),
      save(false),
      { k: "rename", tab: 0, name: "Data 2" },
      { k: "recalc" },
      save(),
    ]);
  }, 400_000);

  it("fill series, drag fill, paste and move of formula blocks across chunk boundaries, then row inserts above them", async () => {
    await expectSame(2, 1500, [
      { k: "drag", tab: 0, r1: 497, c1: 10, r2: 497, c2: 11, grow: 80, dir: "down" },
      { k: "fill", tab: 1, r1: 768, c1: 10, r2: 830, c2: 14, dir: "down" },
      { k: "paste", tab: 0, r1: 497, c1: 10, r2: 501, c2: 15, toTab: 1, row: 1000, col: 10 },
      { k: "move", tab: 0, r1: 497, c1: 10, r2: 505, c2: 12, dr: 250, dc: 0 },
      rows("insert", 0, 40, 9),
      rows("delete", 1, 900, 120),
      save(false),
      { k: "recalc" },
      save(),
    ]);
  }, 400_000);

  // BUG: (pre-existing, both modes; found because the two stores keep different chunks in memory): BigSheetStore.apply() skips a change
  // whose chunk is not in memory ("edited chunks are pinned, so this does not happen"). After a save nothing is pinned any more, but the
  // cell steps stay on the undo stack (markSaved only drops structural steps). Type, save, look at other parts of the sheet, press undo:
  // the step is consumed and nothing changes.
  for (const byId of [false, true]) {
    it(`undo of a cell edit still works after a save and after its chunk left memory (${byId ? "formulas by id" : "positional formulas"})`, async () => {
      const w = await world(1, 1500);
      const side = await openSide(w, structuredClone(w.root), byId);
      await typeCell(side, 0, 300, 2, "424242");
      await saveSide(w, side, false);
      for (let c = 20; c < 34; c++) await side.store.load(0, c);
      expect(side.store.hasChunk(0, side.store.chunkFor(0, 300))).toBe(false);
      side.store.undo();
      await side.store.load(0, side.store.chunkFor(0, 300));
      expect(side.store.cell(0, 300, 2)?.value).not.toBe(424242);
    }, 300_000);
  }

  // BUG: (new with the switch-over): the first structural edit of a workbook stored by position switches it to ids (enableRefsById),
  // which leaves every chunk that holds formulas edited and therefore pinned in the 6-chunk cache. A delete whose first and last
  // rows lie in two different partly-deleted chunks loads them one after the other, the second load evicts the first, and
  // deleteRows throws "Cannot read properties of undefined (reading 'cells')" half way (sheet size already changed). The positional
  // store handles the same edit, because nothing is pinned yet.
  it("the first insert or delete after the switch-over works across a chunk boundary", async () => {
    await expectSame(3, 1500, [rows("delete", 0, 434, 6), save()]);
  }, 400_000);

  // BUG: (exists in the positional mode too, so it is not a difference between the stores): after deleting a column, a chunk that
  // is not in memory is read back with the column ids of the new list. Columns that were never given an id are called "c<position>",
  // and a stored cell whose id is no longer in the list falls back to "c<n> means position n" (colIndexOf in workbook/ids.ts).
  // The deleted column's cells therefore come back at the position of the column that moved into its place, and collide with it.
  for (const byId of [false, true]) {
    it(`deleting a column does not bring its cells back in neighbouring columns after save and reopen (${byId ? "formulas by id" : "positional formulas"})`, async () => {
      const w = await world(1, 600);
      let side = await openSide(w, structuredClone(w.root), byId);
      // column E is empty in some rows: there the cell of column D that was deleted has nothing to be overwritten by
      for (const r of [10, 11, 12, 200, 201, 400]) await typeCell(side, 0, r, 4, "");
      side = await saveSide(w, side, true);
      const before = await dump(side.store);
      await editStructure(side.store, side.index, 0, { kind: "delete", axis: "col", index: 3, count: 1 });
      const saved = await saveSide(w, side, true);
      const after = await dump(saved.store);
      // expected: every cell of old column c > 3 sits at c - 1, old column 3 is gone
      const expected: Record<string, string> = {};
      for (const [key, value] of Object.entries(before.tabs[0].cells)) {
        const [r, c] = key.split(":").map(Number);
        if (c === 3) continue;
        expected[`${r}:${c > 3 ? c - 1 : c}`] = value;
      }
      const wrong = Object.keys({ ...expected, ...after.tabs[0].cells }).filter((k) => expected[k] !== after.tabs[0].cells[k]);
      expect(wrong.length, `cells that differ from the expected result, e.g. ${wrong.slice(0, 5).join(", ")}`).toBe(0);
    }, 300_000);
  }

  // BUG: (pre-existing in the cache, reachable in both modes): BigSheetStore.applyEdit loads the tab's first chunk, then insertRows /
  // deleteRows load other chunks. Edited chunks are pinned in the 6-chunk cache, so with 6 or more of them every other load evicts
  // the clean first chunk, and `head[key]` at the end of applyEdit throws "Cannot convert undefined or null to object" after the rows
  // were already inserted (sheet size and chunk list changed, layout and ids not): the sheet is left half edited.
  it("inserting rows works while six or more other chunks hold unsaved edits", async () => {
    const w = await world(2, 1500);
    const side = await openSide(w, structuredClone(w.root), true);
    for (let k = 0; k < 8; k++) await typeCell(side, 1, 100 + k * 150, 2, String(k));
    expect(side.store.dirtyCount).toBeGreaterThanOrEqual(6);
    const rowsBefore = side.store.tabs[0].rows;
    let error: unknown = null;
    try {
      await editStructure(side.store, new FormulaIndex(), 0, { kind: "insert", axis: "row", index: 700, count: 3 });
    } catch (e) {
      error = e;
    }
    expect(error, `threw ${String(error)}; rows ${rowsBefore} -> ${side.store.tabs[0].rows}`).toBeNull();
    void apply;
  }, 300_000);
});
