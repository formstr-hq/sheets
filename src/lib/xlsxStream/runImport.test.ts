import { describe, expect, it } from "vitest";
import { fortuneSheetsToWorkbook, type WorkbookCell } from "../workbook";
import { assembleChunks } from "../v2/chunks";
import { unframe } from "../v2/codec";
import { parseSpreadsheetFile } from "../xlsx";
import { buildFixture, buildLargeXlsx, largeValue } from "./testFixtures";
import { zipBlob } from "./testZip";
import { runImport, type EmittedChunk, type ImportOptions, type ImportSink, type ReadProgress, type TabSummary } from "./runImport";

function collector(extra: Partial<ImportSink> = {}) {
  const chunks: EmittedChunk[] = [];
  const tabs: TabSummary[] = [];
  const progress: ReadProgress[] = [];
  const sink: ImportSink = {
    chunk: async (chunk) => void chunks.push(chunk),
    tabDone: (summary) => void tabs.push(summary),
    progress: (p) => void progress.push(p),
    ...extra,
  };
  return { chunks, tabs, progress, sink };
}

const OPTIONS: ImportOptions = { maxChunkBytes: 64 * 1024, headRawBytes: 100_000, rawCap: 1_000_000, progressIntervalMs: 0 };

async function assemble(tab: TabSummary, chunks: EmittedChunk[]) {
  const mine = chunks.filter((chunk) => chunk.tab === tab.tab);
  const byHash = new Map(mine.map((chunk) => [chunk.hash, chunk]));
  const bodies = [];
  for (const entry of tab.entries) bodies.push(await unframe(byHash.get(entry.hash)!.frame));
  return assembleChunks(tab.entries, bodies as never[]);
}

describe("runImport", () => {
  it("imports a fixture to the same cells and layout as the in-memory importer", async () => {
    const bytes = await buildFixture();
    const { chunks, tabs, sink } = collector();
    const summary = await runImport(new File([bytes as BlobPart], "Report 2024.xlsx"), sink, OPTIONS);
    expect(summary.title).toBe("Report 2024");
    expect(tabs.map((t) => [t.name, t.hidden])).toEqual([
      ["Main & Data", false],
      ["Second", false],
      ["Hidden", true],
      ["Empty", false],
      ["Wide", false],
    ]);
    expect(summary.tabs).toEqual(tabs);
    expect(summary.cells).toBe(tabs.reduce((sum, t) => sum + t.cells, 0));
    expect(summary.dropped.join("\n")).toMatch(/1 data validation/);

    const imported = await parseSpreadsheetFile(new File([bytes as BlobPart], "x.xlsx"), "t");
    const old = fortuneSheetsToWorkbook({ title: "t", sheets: imported.sheets }).sheets;
    for (const tab of tabs) {
      const sheet = await assemble(tab, chunks);
      const expected: Record<string, WorkbookCell> = {};
      for (const [key, cell] of Object.entries(old[tab.tab].cells)) {
        const { validation: _validation, ...rest } = cell;
        void _validation;
        expected[key] = rest;
      }
      expect(sheet.cells).toEqual(expected);
      expect(sheet).toMatchObject({ rowCount: old[tab.tab].rowCount, colCount: old[tab.tab].colCount });
      expect(tab.cells).toBe(Object.keys(expected).length);
    }
  });

  it("streams a multi-tab file with shared strings after the sheets into many chunks", async () => {
    const options = { tabs: 3, rows: 4000, cols: 10, sharedStrings: true };
    const bytes = await buildLargeXlsx(options);
    const { chunks, tabs, progress, sink } = collector();
    const summary = await runImport(new Blob([bytes as BlobPart]), sink, OPTIONS);

    expect(summary.cells).toBe(3 * 4001 * 10);
    for (const tab of tabs) {
      expect(tab.filledRows).toBe(4001);
      expect(tab.cells).toBe(4001 * 10);
      expect(tab.chunks).toBeGreaterThan(3);
      expect(chunks.filter((c) => c.tab === tab.tab).at(-1)!.isHead).toBe(true);
      const sheet = await assemble(tab, chunks);
      expect(Object.keys(sheet.cells)).toHaveLength(4001 * 10);
      for (const [row, col] of [[0, 0], [0, 9], [1, 1], [1999, 4], [4000, 9], [4000, 0]]) {
        const cell = sheet.cells[`${row}:${col}`];
        expect(cell.value).toBe(largeValue(tab.tab, row, col));
      }
      expect(sheet.cells["0:0"].style?.bold).toBe(true);
      expect(sheet.rowCount).toBe(4001);
      // Entries tile the rows.
      let at = 0;
      for (const entry of tab.entries) {
        expect(entry.rowStart).toBe(at);
        at += entry.rowCount;
      }
      expect(at).toBe(4001);
    }
    // Progress is monotonic and ends complete.
    expect(progress.length).toBeGreaterThan(5);
    for (let i = 1; i < progress.length; i++) {
      expect(progress[i].bytesRead).toBeGreaterThanOrEqual(progress[i - 1].bytesRead);
      expect(progress[i].cells).toBeGreaterThanOrEqual(progress[i - 1].cells);
    }
    expect(progress.at(-1)!.bytesRead).toBe(progress.at(-1)!.bytesTotal);
    expect(progress.at(-1)!.cells).toBe(summary.cells);
  });

  it("reads inline strings when the file has no shared strings table", async () => {
    const bytes = await buildLargeXlsx({ tabs: 1, rows: 100, cols: 4, sharedStrings: false });
    const { chunks, tabs, sink } = collector();
    await runImport(new Blob([bytes as BlobPart]), sink, OPTIONS);
    const sheet = await assemble(tabs[0], chunks);
    expect(sheet.cells["5:1"].value).toBe(largeValue(0, 5, 1));
  });

  it("waits for the sink before reading further (back-pressure)", async () => {
    const bytes = await buildLargeXlsx({ tabs: 1, rows: 6000, cols: 10, sharedStrings: true });
    let release: () => void = () => {};
    const gate = new Promise<void>((resolve) => (release = resolve));
    let seen = 0;
    const { progress, sink } = collector({
      chunk: async () => {
        seen++;
        await gate;
      },
    });
    const running = runImport(new Blob([bytes as BlobPart]), sink, OPTIONS);
    await new Promise((resolve) => setTimeout(resolve, 400));
    expect(seen).toBe(1);
    const stalled = progress.at(-1)!.cells;
    await new Promise((resolve) => setTimeout(resolve, 200));
    expect(progress.at(-1)!.cells).toBe(stalled);
    expect(stalled).toBeLessThan(60_000);
    release();
    const summary = await running;
    expect(summary.cells).toBe(6001 * 10);
  });

  it("stops with CANCELLED when the signal is aborted", async () => {
    const bytes = await buildLargeXlsx({ tabs: 2, rows: 6000, cols: 10, sharedStrings: true });
    const controller = new AbortController();
    const { sink } = collector({
      chunk: async () => {
        controller.abort();
      },
    });
    await expect(runImport(new Blob([bytes as BlobPart]), sink, { ...OPTIONS, signal: controller.signal })).rejects.toMatchObject({ code: "CANCELLED" });
  });

  it("aborts cleanly when the shared strings table is over budget", async () => {
    const bytes = await buildLargeXlsx({ tabs: 1, rows: 3000, cols: 6, sharedStrings: true });
    const { sink } = collector();
    await expect(runImport(new Blob([bytes as BlobPart]), sink, { ...OPTIONS, stringBudgetBytes: 20_000 })).rejects.toMatchObject({
      code: "STRING_TABLE_TOO_LARGE",
    });
  });

  it("reports files that are not xlsx", async () => {
    const { sink } = collector();
    await expect(runImport(new Blob(["a,b,c\n1,2,3\n"]), sink, OPTIONS)).rejects.toMatchObject({ code: "NOT_XLSX" });
    await expect(runImport(zipBlob([{ name: "readme.txt", data: "hi" }]), sink, OPTIONS)).rejects.toMatchObject({ code: "NOT_XLSX" });
    const encrypted = zipBlob([
      { name: "xl/workbook.xml", data: "x", encrypted: true },
    ]);
    await expect(runImport(encrypted, sink, OPTIONS)).rejects.toMatchObject({ code: "UNSUPPORTED_ZIP" });
    // An OLE2 container (password-protected xlsx, or an old .xls) starts with this signature.
    const ole = new Blob([new Uint8Array([0xd0, 0xcf, 0x11, 0xe0, 0xa1, 0xb1, 0x1a, 0xe1]), new Uint8Array(600)]);
    await expect(runImport(ole, sink, OPTIONS)).rejects.toMatchObject({ code: "NOT_XLSX" });
  });
});
