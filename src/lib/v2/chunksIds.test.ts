import { describe, expect, it } from "vitest";
import type { WorkbookCell, WorkbookSheet } from "../workbook";
import { insertRuns, legacyRows, mintRows, rowCountOf } from "../workbook/ids";
import { assembleChunks, colsForSheet, planTab, readChunk, writeChunk, type ChunkMemory } from "./chunks";
import { frame, unframe } from "./codec";
import { resolveChunks, type ChunkEntry } from "./root";

const MAX = 64 * 1024;

function sheet(rows: number, cols: number): WorkbookSheet {
  const cells: Record<string, WorkbookCell> = {};
  for (let r = 0; r < rows; r++) {
    for (let c = 0; c < cols; c++) cells[`${r}:${c}`] = { value: `row ${r} col ${c} ${"x".repeat((r * 7 + c) % 23)}`, type: "text" };
  }
  return { id: "s1", name: "S", order: 0, rowCount: rows, colCount: cols, cells };
}

type Plan = Awaited<ReturnType<typeof planTab>>;
const entries = (plan: Plan) => plan.map((chunk) => ({ rowStart: chunk.rowStart, rowCount: chunk.rowCount }));
const bodies = (plan: Plan) =>
  Promise.all(plan.map((chunk) => unframe(chunk.frame as Uint8Array))) as Promise<Parameters<typeof assembleChunks>[1]>;
const memoryOf = (plan: Plan) =>
  plan.map((chunk) => ({
    id: chunk.id,
    ref: { h: chunk.hash, k: "k", p: chunk.size },
    rowStart: chunk.rowStart,
    rowCount: chunk.rowCount,
    servers: [],
  })) as unknown as ChunkMemory[];

describe("id chunks", () => {
  it("store row ids and column ids, and read back to the same sheet", async () => {
    const source = sheet(3000, 8);
    const plan = await planTab(source, { max: MAX });
    expect(plan.length).toBeGreaterThan(1);
    const stored = await bodies(plan);
    expect(stored[0].rows).toBeDefined();
    expect(Object.keys(stored[0].cells ?? {})[0]).toBe("0:c0");
    const cols = colsForSheet(source);
    expect(cols).toEqual(["c0", "c1", "c2", "c3", "c4", "c5", "c6", "c7"]);
    const back = assembleChunks(entries(plan), stored, cols);
    expect(back.cells).toEqual(source.cells);
    expect(back.rowIds).toEqual(legacyRows(0, 3000));
    expect(back.colIds).toEqual(cols);
  });

  it("an inserted row changes only the chunk it lands in", async () => {
    const source = sheet(3000, 8);
    const first = await planTab(source, { max: MAX });
    const assembled = assembleChunks(entries(first), await bodies(first), colsForSheet(source));

    // Insert a row at position 5: every later row moves down one position but keeps its id.
    const rowIds = insertRuns(assembled.rowIds ?? [], 5, mintRows("dev1", 0, 1));
    const cells: Record<string, WorkbookCell> = {};
    for (const [key, cell] of Object.entries(assembled.cells)) {
      const [r, c] = key.split(":").map(Number);
      cells[`${r >= 5 ? r + 1 : r}:${c}`] = cell;
    }
    // The new row has a cell too, so it is stored.
    cells["5:0"] = { value: "new row", type: "text" };
    expect(rowCountOf(rowIds)).toBe(3001);
    const edited: WorkbookSheet = { ...source, rowCount: 3001, cells, rowIds, colIds: assembled.colIds };

    // Same cut points as before (shifted by the insert), so chunks that did not gain the row keep their bytes.
    const previous = memoryOf(first).map((chunk, i) => ({ ...chunk, rowStart: i === 0 ? 0 : chunk.rowStart + 1 }));
    const second = await planTab(edited, { max: MAX, previous });
    const changed = second.filter((chunk) => !chunk.reuse).length;
    expect(second.length).toBeGreaterThan(2);
    expect(changed).toBe(1);
  });

  it("legacy chunks (no row ids, numeric columns) still read", async () => {
    const legacy0 = await frame({ rowCount: 10, colCount: 3, cells: { "0:0": { value: 1 }, "1:2": { value: 2 } } });
    const legacy1 = await frame({ cells: { "0:1": { value: 3 } } });
    const read = assembleChunks(
      [
        { rowStart: 0, rowCount: 5 },
        { rowStart: 5, rowCount: 5 },
      ],
      [(await unframe(legacy0)) as never, (await unframe(legacy1)) as never],
    );
    expect(read.cells).toEqual({ "0:0": { value: 1 }, "1:2": { value: 2 }, "5:1": { value: 3 } });
    expect(read.rowIds).toBeUndefined();
    expect(read.colIds).toBeUndefined();
    expect(readChunk({ cells: {} }, 5, 5).rows).toEqual([["0000", 5, 5]]);
  });

  it("readChunk drops cells of a deleted column and writeChunk is its inverse", () => {
    const stored = {
      rows: [["a", 0, 2]] as [string, number, number][],
      cells: { "0:k1": { value: 1 }, "1:gone": { value: 2 }, "1:k2": { value: 3 } },
    };
    const read = readChunk(stored, 0, 2, ["k1", "k2"]);
    expect(read.cells).toEqual({ "0:0": { value: 1 }, "1:1": { value: 3 } });
    expect(writeChunk(read, ["k1", "k2"])).toEqual({
      rows: stored.rows,
      cells: { "0:k1": { value: 1 }, "1:k2": { value: 3 } },
    });
  });

  it("the manifest does not need rowStart: it is the sum of the counts before", async () => {
    const list = [{ id: "a", rowCount: 4 }, { id: "b", rowCount: 6 }, { id: "c", rowCount: 1 }] as unknown as ChunkEntry[];
    const out = await resolveChunks({ id: "t", name: "t", order: 0, chunks: list }, [], async () => ({}));
    expect(out?.map((entry) => entry.rowStart)).toEqual([0, 4, 10]);
  });
});
