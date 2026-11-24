import { describe, expect, it } from "vitest";
import type { WorkbookCell, WorkbookSheet } from "../workbook";
import { assembleChunks, fillBytes, MIN_CHUNK_BYTES, planTab, type ChunkMemory } from "./chunks";
import { ChunkStreamWriter, type RowInput, type StreamedChunk } from "./chunkStream";
import { unframe } from "./codec";

const MAX = 64 * 1024;

function lcg(seed: number) {
  let s = seed >>> 0;
  return () => (s = (Math.imul(s, 1664525) + 1013904223) >>> 0) / 2 ** 32;
}

interface Built {
  sheet: WorkbookSheet;
  rows: RowInput[];
}

function build(rowCount: number, cols: number, seed: number, mode: "random" | "flat" | "sparse" = "random"): Built {
  const rnd = lcg(seed);
  const cells: Record<string, WorkbookCell> = {};
  const rows: RowInput[] = [];
  for (let r = 0; r < rowCount; r++) {
    if (mode === "sparse" && r % 11 !== 0) continue;
    const row: RowInput = { row: r, cells: [] };
    for (let c = 0; c < cols; c++) {
      if (mode === "sparse" && c % 4 !== r % 4) continue;
      const cell: WorkbookCell =
        mode === "flat"
          ? { value: c, type: "number" }
          : (r + c) % 2
            ? { value: Math.round(rnd() * 1e6) / 100, display: String(r), type: "number" }
            : { value: `w-${Math.floor(rnd() * 1e12).toString(36)}`, type: "text" };
      row.cells.push([c, cell]);
      cells[`${r}:${c}`] = cell;
    }
    if (row.cells.length > 0) rows.push(row);
  }
  const sheet: WorkbookSheet = {
    id: "s",
    name: "S",
    order: 0,
    rowCount: rowCount + 5,
    colCount: cols + 1,
    cells,
    columnWidths: { 0: 100 },
    merges: [{ r: 0, c: 0, rs: 2, cs: 2 }],
  };
  return { sheet, rows };
}

const layoutOf = (sheet: WorkbookSheet): Record<string, unknown> => {
  const { id: _id, name: _name, order: _order, hidden: _hidden, color: _color, showGridLines: _grid, cells: _cells, ...layout } = sheet;
  void [_id, _name, _order, _hidden, _color, _grid, _cells];
  return layout;
};

async function run(built: Built, options: { max?: number; headRawBytes?: number; rawCap?: number } = {}) {
  const emitted: StreamedChunk[] = [];
  const writer = new ChunkStreamWriter({
    max: options.max ?? MAX,
    headRawBytes: options.headRawBytes,
    rawCap: options.rawCap,
    onChunk: async (chunk) => void emitted.push(chunk),
  });
  for (const row of built.rows) await writer.push(row);
  const infos = await writer.finish(layoutOf(built.sheet));
  return { emitted, infos, writer };
}

async function reassemble(emitted: StreamedChunk[], infos: { rowStart: number; hash: string }[]) {
  const byHash = new Map(emitted.map((chunk) => [chunk.hash, chunk]));
  const bodies = [];
  for (const info of infos) bodies.push(await unframe(byHash.get(info.hash)!.frame));
  return assembleChunks(infos, bodies as never[]);
}

describe("ChunkStreamWriter", () => {
  it("writes a small sheet as one chunk with the same bytes as planTab", async () => {
    const built = build(40, 6, 1);
    const { emitted, infos } = await run(built, { max: 8 * 1024 * 1024 });
    const planned = await planTab(built.sheet, { max: 8 * 1024 * 1024 });
    expect(emitted).toHaveLength(1);
    expect(infos).toHaveLength(1);
    expect(emitted[0].isHead).toBe(true);
    expect(emitted[0].hash).toBe(planned[0].hash);
    expect(emitted[0].frame).toEqual(planned[0].frame);
    expect([infos[0].rowStart, infos[0].rowCount]).toEqual([planned[0].rowStart, planned[0].rowCount]);
    expect(infos[0].cells).toBe(40 * 6);
  });

  it("emits the head chunk last, with the layout, and chunks identical to planTab for the same cuts", async () => {
    const built = build(6000, 12, 2);
    const { emitted, infos } = await run(built, { headRawBytes: 150_000 });
    expect(infos.length).toBeGreaterThan(5);
    expect(emitted.at(-1)!.isHead).toBe(true);
    expect(emitted.filter((chunk) => chunk.isHead)).toHaveLength(1);
    expect(emitted.map((chunk) => chunk.seq)).toEqual(emitted.map((_, i) => i));
    expect(infos[0]).toMatchObject({ rowStart: 0, isHead: true });
    // Layout is in the head body only.
    const bodies = await Promise.all(emitted.map((chunk) => unframe(chunk.frame) as Promise<Record<string, unknown>>));
    expect(bodies.filter((body) => "merges" in body)).toHaveLength(1);
    expect("merges" in bodies.at(-1)!).toBe(true);

    // Ranges tile the sheet.
    let expectedStart = 0;
    for (const info of infos) {
      expect(info.rowStart).toBe(expectedStart);
      expectedStart += info.rowCount;
    }
    expect(expectedStart).toBe(built.sheet.rowCount);
    expect(infos.reduce((sum, info) => sum + info.cells, 0)).toBe(6000 * 12);

    // planTab, told to cut at the same rows, produces byte-identical chunks.
    const previous = infos.map((info) => ({ id: info.id, ref: { h: info.hash }, rowStart: info.rowStart, rowCount: info.rowCount, servers: [] })) as unknown as ChunkMemory[];
    const planned = await planTab(built.sheet, { max: MAX, previous });
    expect(planned.map((chunk) => chunk.hash)).toEqual(infos.map((info) => info.hash));
    expect(planned.every((chunk) => chunk.reuse)).toBe(true);
    const fresh = await planTab(built.sheet, { max: MAX, previous: previous.map((p) => ({ ...p, ref: { h: "x" } })) as unknown as ChunkMemory[] });
    expect(fresh.map((chunk) => chunk.hash)).toEqual(infos.map((info) => info.hash));
    expect(fresh.map((chunk) => [chunk.rowStart, chunk.rowCount])).toEqual(infos.map((info) => [info.rowStart, info.rowCount]));
  });

  it("reassembles to the source sheet", async () => {
    const built = build(5000, 10, 3);
    const { emitted, infos } = await run(built, { headRawBytes: 100_000 });
    const sheet = await reassemble(emitted, infos);
    expect(sheet.cells).toEqual(built.sheet.cells);
    expect(sheet).toMatchObject({ rowCount: built.sheet.rowCount, colCount: built.sheet.colCount, merges: built.sheet.merges, columnWidths: { 0: 100 } });
  });

  it("handles sparse rows and gaps", async () => {
    const built = build(400_000, 8, 4, "sparse");
    const { emitted, infos } = await run(built, { headRawBytes: 50_000 });
    expect(infos.length).toBeGreaterThan(1);
    const sheet = await reassemble(emitted, infos);
    expect(sheet.cells).toEqual(built.sheet.cells);
    const last = infos.at(-1)!;
    expect(last.rowStart + last.rowCount).toBe(built.sheet.rowCount);
  });

  it("writes an empty tab as one layout-only chunk equal to planTab", async () => {
    const built = build(0, 0, 5);
    const { emitted, infos } = await run(built);
    const planned = await planTab(built.sheet, { max: MAX });
    expect(infos).toHaveLength(1);
    expect(infos[0]).toMatchObject({ rowStart: 0, rowCount: 0, cells: 0, isHead: true });
    expect(emitted[0].hash).toBe(planned[0].hash);
  });

  it("ignores rows without cells and rejects rows out of order", async () => {
    const writer = new ChunkStreamWriter({ max: MAX, onChunk: async () => {} });
    await writer.push({ row: 3, cells: [] });
    await writer.push({ row: 5, cells: [[0, { value: 1 }]] });
    await expect(writer.push({ row: 5, cells: [[0, { value: 2 }]] })).rejects.toThrow(/not after/);
    const infos = await writer.finish({});
    expect(infos).toHaveLength(1);
    expect(infos[0]).toMatchObject({ rowStart: 0, rowCount: 6, cells: 1 });
    await expect(writer.finish({})).rejects.toThrow();
  });

  it("keeps a single row larger than the max in one chunk", async () => {
    const noise = lcg(9);
    const big: RowInput = { row: 2, cells: [[0, { value: Array.from({ length: 40_000 }, () => noise().toString(36)).join("") }]] };
    const chunks: StreamedChunk[] = [];
    const sink = new ChunkStreamWriter({ max: MIN_CHUNK_BYTES, headRawBytes: 100, onChunk: async (chunk) => void chunks.push(chunk) });
    await sink.push({ row: 0, cells: [[0, { value: "a" }]] });
    await sink.push({ row: 1, cells: [[0, { value: "b" }]] });
    await sink.push(big);
    await sink.push({ row: 3, cells: [[0, { value: "c" }]] });
    const infos = await sink.finish({});
    const huge = infos.find((info) => info.rowStart <= 2 && info.rowStart + info.rowCount > 2 && info.cells === 1 && info.size > MIN_CHUNK_BYTES);
    expect(huge).toBeDefined();
    expect(infos.reduce((sum, info) => sum + info.cells, 0)).toBe(4);
  });

  it("applies back-pressure: while the sink is busy, memory held stays bounded", async () => {
    const rawCap = 400_000;
    let release: () => void = () => {};
    let blocked = 0;
    const gate = new Promise<void>((resolve) => (release = resolve));
    const writer = new ChunkStreamWriter({
      max: MAX,
      headRawBytes: 50_000,
      rawCap,
      onChunk: async () => {
        blocked++;
        await gate;
      },
    });
    const built = build(20_000, 10, 6);
    let pushed = 0;
    let peak = 0;
    const producer = (async () => {
      for (const row of built.rows) {
        await writer.push(row);
        pushed++;
        peak = Math.max(peak, writer.heldBytes);
      }
    })();
    await new Promise((resolve) => setTimeout(resolve, 300));
    expect(blocked).toBe(1); // the producer is stuck inside push(), waiting for the sink
    const stalledAt = pushed;
    await new Promise((resolve) => setTimeout(resolve, 100));
    expect(pushed).toBe(stalledAt);
    expect(pushed).toBeLessThan(built.rows.length);
    expect(peak).toBeLessThanOrEqual(rawCap + 50_000 + 2_000);
    release();
    await producer;
    await writer.finish({});
    expect(peak).toBeLessThanOrEqual(rawCap + 50_000 + 2_000);
  });

  for (const mode of ["random", "flat"] as const) {
    it(`cuts chunks between half the fill size and the max (${mode} data)`, async () => {
      const built = build(mode === "flat" ? 60_000 : 12_000, 14, 7, mode);
      const { infos } = await run(built, { headRawBytes: 100_000, rawCap: 8 * 1024 * 1024 });
      expect(infos.length).toBeGreaterThan(4);
      const inner = infos.slice(1, -1);
      expect(inner.length).toBeGreaterThan(2);
      for (const info of inner) {
        expect(info.size).toBeLessThanOrEqual(MAX);
        expect(info.size).toBeGreaterThanOrEqual(fillBytes(MAX) * 0.5);
      }
    }, 30_000); // 60k rows: several seconds on a busy machine
  }

  it("reads the size limit before every cut", async () => {
    let limit = 256 * 1024;
    const chunks: StreamedChunk[] = [];
    const writer = new ChunkStreamWriter({
      max: () => limit,
      headRawBytes: 50_000,
      onChunk: async (chunk) => {
        chunks.push(chunk);
        limit = MIN_CHUNK_BYTES; // a server told us it only takes 64 KB
      },
    });
    for (const row of build(40_000, 10, 8).rows) await writer.push(row);
    await writer.finish({});
    const tail = chunks.slice(1, -1);
    expect(tail.length).toBeGreaterThan(3);
    for (const chunk of tail.slice(1)) expect(chunk.frame.length).toBeLessThanOrEqual(MIN_CHUNK_BYTES);
  });
});
