import { describe, expect, it } from "vitest";
import type { WorkbookCell, WorkbookSheet } from "../workbook";
import { canonicalJson, unframe } from "./codec";
import { sha256Hex } from "./hash";
import { planTab, type ChunkMemory, type PlannedChunk } from "./chunks";

// Golden values for planTab, recorded before chunks.ts was refactored to share its encoder with the
// streaming writer. The refactor must not change what gets stored. Each chunk is summarised by its
// row range and the hash of its canonical JSON body (not the deflated bytes, so a different zlib
// build cannot break the test; cut points still depend on compressed sizes).

function lcg(seed: number) {
  let s = seed >>> 0;
  return () => (s = (Math.imul(s, 1664525) + 1013904223) >>> 0) / 2 ** 32;
}

function makeSheet(rows: number, cols: number, seed: number, opts: { gaps?: boolean; styled?: boolean } = {}): WorkbookSheet {
  const rnd = lcg(seed);
  const cells: Record<string, WorkbookCell> = {};
  for (let r = 0; r < rows; r++) {
    if (opts.gaps && r % 7 === 3) continue;
    for (let c = 0; c < cols; c++) {
      if (opts.gaps && rnd() < 0.3) continue;
      const kind = (r + c) % 3;
      const cell: WorkbookCell =
        kind === 0
          ? { value: Math.round(rnd() * 1e6) / 100, display: String(r), type: "number" }
          : kind === 1
            ? { value: `w-${Math.floor(rnd() * 1e9).toString(36)}`, type: "text" }
            : { value: r * (c + 1), type: "number" };
      if (opts.styled && c === 0) cell.style = { bold: true, background: "#ffee00" };
      cells[`${r}:${c}`] = cell;
    }
  }
  return {
    id: "s1",
    name: "Golden",
    order: 0,
    rowCount: rows + 10,
    colCount: cols + 2,
    cells,
    columnWidths: { 0: 120, 3: 80 },
    rowHeights: { 1: 40 },
    merges: [{ r: 0, c: 0, rs: 1, cs: 2 }],
    defaultRowHeight: 30,
    defaultColWidth: 116,
  };
}

async function summarise(chunks: PlannedChunk[]) {
  const out: string[] = [];
  for (const chunk of chunks) {
    const body = await unframe(chunk.frame as Uint8Array);
    const digest = await sha256Hex(new TextEncoder().encode(canonicalJson(body)));
    out.push(`${chunk.rowStart}+${chunk.rowCount}:${digest.slice(0, 16)}`);
  }
  return out;
}

const MAX = 64 * 1024;

describe("planTab golden output", () => {
  it("small sheet stays one chunk", async () => {
    expect(await summarise(await planTab(makeSheet(50, 6, 1), { max: 8 * 1024 * 1024 }))).toMatchInlineSnapshot(`
      [
        "0+60:84630da0fe2ee996",
      ]
    `);
  });

  it("empty sheet is one layout-only chunk", async () => {
    expect(await summarise(await planTab(makeSheet(0, 0, 1), { max: 8 * 1024 * 1024 }))).toMatchInlineSnapshot(`
      [
        "0+10:096deee32e478ef3",
      ]
    `);
  });

  it("fresh cut of a multi-chunk sheet", async () => {
    expect(await summarise(await planTab(makeSheet(6000, 12, 2), { max: MAX }))).toMatchInlineSnapshot(`
      [
        "0+355:c07be5f501819679",
        "355+498:2d4184edcbabb9d1",
        "853+488:8f592fbf67803a89",
        "1341+492:4dae26990aec1e5d",
        "1833+496:2681488b9cee110e",
        "2329+496:21a27b741208f565",
        "2825+497:aa47303aa0aad0a5",
        "3322+497:34ef846bcec2fbba",
        "3819+497:b39e52518b8598bd",
        "4316+497:0cb123572c2df7d4",
        "4813+497:450a03e4dfd67d65",
        "5310+497:32e99f7f7efb8a5a",
        "5807+203:28b7433938e5c1ec",
      ]
    `);
  });

  it("fresh cut with gaps, sparse rows and styles", async () => {
    expect(await summarise(await planTab(makeSheet(9000, 8, 3, { gaps: true, styled: true }), { max: MAX }))).toMatchInlineSnapshot(`
      [
        "0+790:8f62e4ddd467ea50",
        "790+1137:d135c2084701dc1c",
        "1927+1139:41e6263be15c30bc",
        "3066+1176:bf3f41047b7a0ab3",
        "4242+1141:d7b60129fe44c26e",
        "5383+1168:789774f9dbb7ef42",
        "6551+1158:b57f40e26e424df3",
        "7709+1147:815ee8b12eb2580c",
        "8856+154:9c342d197427175e",
      ]
    `);
  });

  it("re-save with previous cut points and an edit", async () => {
    const sheet = makeSheet(6000, 12, 2);
    const first = await planTab(sheet, { max: MAX });
    const memory: ChunkMemory[] = first.map((chunk) => ({
      id: chunk.id,
      ref: { h: chunk.hash, k: "k", p: chunk.size },
      rowStart: chunk.rowStart,
      rowCount: chunk.rowCount,
      servers: [],
    })) as unknown as ChunkMemory[];
    sheet.cells["2500:1"] = { value: "edited", type: "text" };
    for (let c = 0; c < 12; c++) delete sheet.cells[`4000:${c}`];
    const second = await planTab(sheet, { max: MAX, previous: memory });
    expect(await summarise(second.map((chunk) => ({ ...chunk, frame: chunk.frame ?? new Uint8Array() })).filter((c) => c.frame.length > 0))).toMatchInlineSnapshot(`
      [
        "2329+496:d6a3259e1cd4ac27",
        "3819+497:5b7c3d268685880f",
      ]
    `);
    expect(second.map((c) => `${c.rowStart}+${c.rowCount}${c.reuse ? "R" : ""}`)).toMatchInlineSnapshot(`
      [
        "0+355R",
        "355+498R",
        "853+488R",
        "1341+492R",
        "1833+496R",
        "2329+496",
        "2825+497R",
        "3322+497R",
        "3819+497",
        "4316+497R",
        "4813+497R",
        "5310+497R",
        "5807+203R",
      ]
    `);
  });
});
