/// <reference types="node" />
import { statSync, openAsBlob, writeFileSync } from "node:fs";
import { basename } from "node:path";
import { describe, expect, it } from "vitest";
import { runImport, type TabSummary } from "../src/lib/xlsxStream/runImport";

// Manual measurement, skipped unless BIG_XLSX points at a file. Runs the real streaming import
// (zip -> parse -> chunk -> deflate -> hash) in Node with a sink that discards the frames, so the
// numbers are the worker's cost without any upload. Use a small heap to prove memory is bounded:
//
//   BIG_XLSX=C:/Users/me/Downloads/big-demo-99.9mb.xlsx \
//   NODE_OPTIONS=--max-old-space-size=384 pnpm vitest run tools/stream-import.test.ts
//
// Optional: BIG_XLSX_EXPECT='{"cells":11032560,"sheets":[{"rows":91938,"cells":2758140},...]}' (the
// output of `node tools/count-xlsx.mjs <file>`) to check against an independent count.
const file = process.env.BIG_XLSX;
const mb = (n: number) => Math.round(n / 1024 / 1024);

describe.skipIf(!file)("big xlsx streaming import", () => {
  it("imports without holding the sheet in memory", { timeout: 3_600_000 }, async () => {
    const path = file as string;
    const size = statSync(path).size;
    const source = await openAsBlob(path);
    let peakRss = 0;
    let peakHeap = 0;
    const sample = () => {
      const m = process.memoryUsage();
      peakRss = Math.max(peakRss, m.rss);
      peakHeap = Math.max(peakHeap, m.heapUsed);
    };
    const timer = setInterval(sample, 100);
    let chunks = 0;
    let frameBytes = 0;
    let maxFrame = 0;
    let lastLog = 0;
    const started = performance.now();
    const tabs: TabSummary[] = [];
    try {
      const summary = await runImport(
        source,
        {
          chunk: async (chunk) => {
            chunks++;
            frameBytes += chunk.frame.length;
            maxFrame = Math.max(maxFrame, chunk.frame.length);
          },
          tabDone: (tab) => void tabs.push(tab),
          progress: (p) => {
            if (Date.now() - lastLog > 10_000) {
              lastLog = Date.now();
              console.log(`  ... tab ${p.tab + 1}/${p.tabCount} ${mb(p.bytesRead)}/${mb(p.bytesTotal)} MB, ${p.cells} cells, rss ${mb(process.memoryUsage().rss)} MB`);
            }
          },
        },
        { maxChunkBytes: 8 * 1024 * 1024 },
      );
      sample();
      const seconds = (performance.now() - started) / 1000;
      const report = [
        `[streamImport ${basename(path)}] ${mb(size)} MB file`,
        `  time ${seconds.toFixed(1)} s, ${Math.round(summary.cells / seconds)} cells/s`,
        `  cells ${summary.cells}, tabs ${tabs.length}, rows per tab ${tabs.map((t) => t.filledRows).join("/")}`,
        `  chunks ${chunks}, ${mb(frameBytes)} MB compressed, largest ${(maxFrame / 1024 / 1024).toFixed(2)} MB`,
        `  peak rss ${mb(peakRss)} MB, peak heap ${mb(peakHeap)} MB`,
        `  dropped: ${summary.dropped.join(" | ") || "nothing"}`,
      ].join("\n");
      console.log(report);
      // vitest may swallow console output; BIG_XLSX_REPORT=<path> also writes the report to a file.
      if (process.env.BIG_XLSX_REPORT) writeFileSync(process.env.BIG_XLSX_REPORT, report + "\n");

      expect(summary.cells).toBe(tabs.reduce((sum, t) => sum + t.cells, 0));
      for (const tab of tabs) {
        expect(tab.entries.reduce((sum, e) => sum + e.cells, 0)).toBe(tab.cells);
        let at = 0;
        for (const e of tab.entries) {
          expect(e.rowStart).toBe(at);
          at += e.rowCount;
        }
      }
      if (process.env.BIG_XLSX_EXPECT) {
        const expected = JSON.parse(process.env.BIG_XLSX_EXPECT) as { cells: number; sheets: { rows: number; cells: number }[] };
        expect(summary.cells).toBe(expected.cells);
        expect(tabs.map((t) => ({ rows: t.filledRows, cells: t.cells }))).toEqual(expected.sheets.map((s) => ({ rows: s.rows, cells: s.cells })));
      }
    } finally {
      clearInterval(timer);
    }
  });
});
