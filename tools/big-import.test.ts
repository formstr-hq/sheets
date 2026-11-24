import { readFileSync } from "node:fs";
import { basename } from "node:path";
import { describe, expect, it } from "vitest";
import { buildSheetEvent, decodeSheetEvent, generateDocKeys, type TabMemory } from "../src/lib/crypto";
import { parseSpreadsheetFile } from "../src/lib/xlsx";
import { createWorkbookFromFortune, workbookToFortuneSheets } from "../src/lib/workbook";
import { forgetServerLimits } from "../src/lib/v2/blossom";
import { fakeBlossom, localSigner } from "../src/lib/v2/testUtils";

// Manual measurement, skipped unless BIG_XLSX points at a file:
//   BIG_XLSX=C:/Users/me/Downloads/big-demo-40mb.xlsx pnpm vitest run tools/big-import.test.ts
// Runs the real import pipeline (xlsx parse -> workbook model -> chunked save -> load -> editor
// model) in Node with its default heap limit, which is a fair stand-in for a browser tab.
const file = process.env.BIG_XLSX;
const mb = (n: number) => Math.round(n / 1024 / 1024);

describe.skipIf(!file)("big xlsx import", () => {
  it("imports, saves and loads", { timeout: 3_600_000 }, async () => {
    forgetServerLimits();
    const path = file as string;
    const bytes = readFileSync(path);
    const log: string[] = [];
    let peak = 0;
    const timer = setInterval(() => (peak = Math.max(peak, process.memoryUsage().rss)), 100);
    const step = async <T,>(name: string, work: () => Promise<T> | T): Promise<T> => {
      const t = performance.now();
      const result = await work();
      log.push(`${name}: ${Math.round(performance.now() - t)} ms, heap ${mb(process.memoryUsage().heapUsed)} MB, rss ${mb(process.memoryUsage().rss)} MB`);
      return result;
    };
    try {
      const upload = new File([bytes], basename(path));
      const imported = await step("xlsx parse", () => parseSpreadsheetFile(upload, "Big"));
      const cells = imported.sheets.reduce((sum, sheet) => sum + (sheet.celldata?.length ?? 0), 0);
      log.push(`cells ${cells}, rows per tab ${imported.sheets.map((s) => s.row).join("/")}`);
      const doc = await step("to workbook", () => createWorkbookFromFortune("Big", imported.sheets));
      imported.sheets.length = 0;
      const net = fakeBlossom();
      const keys = generateDocKeys();
      const reuse = new Map<string, TabMemory>();
      const event = await step("chunked save", () =>
        buildSheetEvent({
          sheetId: "b",
          document: doc,
          viewKey: keys.viewKey,
          editKey: keys.editKey,
          blossom: { servers: ["https://a.example"], signer: localSigner(), reuse, env: { fetch: net.fetch } },
        }),
      );
      const stored = [...(net.stores.get("https://a.example")?.values() ?? [])];
      log.push(`stored ${stored.length} blobs, ${mb(stored.reduce((s, b) => s + b.byteLength, 0))} MB`);
      const loaded = await step("load", () => decodeSheetEvent(event, keys.viewKey, "x", { servers: ["https://a.example"], env: { fetch: net.fetch } }));
      await step("to editor model", () => workbookToFortuneSheets(loaded));
      expect(loaded.sheets.length).toBe(doc.sheets.length);
    } finally {
      clearInterval(timer);
      console.log(`[bigImport ${basename(file as string)}]\n  ${log.join("\n  ")}\n  peak rss ${mb(peak)} MB`);
    }
  });
});
