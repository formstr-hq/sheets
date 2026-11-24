import { describe, expect, it } from "vitest";
import { buildSheetEvent, decodeSheetEvent, generateDocKeys, type TabMemory } from "../crypto";
import { createEmptyWorkbook, type WorkbookCell, type WorkbookDocument } from "../workbook";
import { forgetServerLimits, learnServerLimit } from "./blossom";
import { fakeBlossom, localSigner } from "./testUtils";

// A big synthetic workbook, no editor: proves save and load through the fake Blossom keep working
// with many chunks, and prints how long it takes and how much memory it needs.

const TABS = 4;
const ROWS = 10_000;
const COLS = 30;
const WORDS = ["alpha", "bravo", "delta", "omega", "sigma", "kappa", "zeta", "theta"];

function bigWorkbook(): WorkbookDocument {
  let seed = 42;
  const rnd = () => (seed = (seed * 1664525 + 1013904223) >>> 0) / 2 ** 32;
  const doc = createEmptyWorkbook("Large");
  doc.sheets = Array.from({ length: TABS }, (_, t) => {
    const cells: Record<string, WorkbookCell> = {};
    for (let r = 0; r < ROWS; r++) {
      for (let c = 0; c < COLS; c++) {
        cells[`${r}:${c}`] =
          c % 3 === 0
            ? { value: Math.round(rnd() * 1e6) / 100, type: "number" }
            : c % 3 === 1
              ? { value: `${WORDS[Math.floor(rnd() * 8)]}-${Math.floor(rnd() * 1e9).toString(36)}`, type: "text" }
              : { value: r * (c + 1), type: "number" };
      }
    }
    return { id: `t${t}`, name: `Tab ${t + 1}`, order: t, rowCount: ROWS, colCount: COLS, cells };
  });
  doc.activeSheetId = "t0";
  return doc;
}

const proc = (globalThis as unknown as { process: { memoryUsage(): { rss: number } } }).process;
const mb = (bytes: number) => Math.round(bytes / 1024 / 1024);

describe("large workbook", () => {
  it(`saves and loads ${(TABS * ROWS * COLS) / 1e6}M cells across ${TABS} worksheets`, { timeout: 600_000 }, async () => {
    forgetServerLimits();
    // Pretend the server refused a 2 MB upload, so its limit is 1.5 MB: the 2 MB worksheets need several chunks each.
    learnServerLimit("https://a.example", 2 * 1024 * 1024);
    const doc = bigWorkbook();
    const net = fakeBlossom();
    const keys = generateDocKeys();
    const reuse = new Map<string, TabMemory>();

    let peak = proc.memoryUsage().rss;
    const sampler = setInterval(() => {
      peak = Math.max(peak, proc.memoryUsage().rss);
    }, 50);
    const baseline = mb(proc.memoryUsage().rss);

    const t0 = performance.now();
    const event = await buildSheetEvent({
      sheetId: "big",
      document: doc,
      viewKey: keys.viewKey,
      editKey: keys.editKey,
      blossom: { servers: ["https://a.example"], signer: localSigner(), reuse, env: { fetch: net.fetch } },
    });
    const saveMs = performance.now() - t0;
    const stored = [...(net.stores.get("https://a.example")?.values() ?? [])];
    const storedBytes = stored.reduce((sum, bytes) => sum + bytes.byteLength, 0);

    const t1 = performance.now();
    const loaded = await decodeSheetEvent(event, keys.viewKey, "x", { servers: ["https://a.example"], env: { fetch: net.fetch } });
    const loadMs = performance.now() - t1;

    // Unchanged second save: nothing re-uploaded.
    const uploadsBefore = net.uploads.length;
    const t2 = performance.now();
    await buildSheetEvent({
      sheetId: "big",
      document: doc,
      viewKey: keys.viewKey,
      editKey: keys.editKey,
      blossom: { servers: ["https://a.example"], signer: localSigner(), reuse, env: { fetch: net.fetch } },
    });
    const resaveMs = performance.now() - t2;
    clearInterval(sampler);

    console.log(
      `[large] ${TABS * ROWS * COLS} cells: ${stored.length} chunks, ${mb(storedBytes)} MB stored, ` +
        `largest ${mb(Math.max(...stored.map((b) => b.byteLength)))} MB | save ${Math.round(saveMs)} ms, load ${Math.round(loadMs)} ms, ` +
        `unchanged re-save ${Math.round(resaveMs)} ms | rss ${baseline} MB -> peak ${mb(peak)} MB`,
    );

    expect(stored.length).toBeGreaterThan(TABS);
    expect(Math.max(...stored.map((b) => b.byteLength))).toBeLessThanOrEqual(1.5 * 1024 * 1024 + 1024);
    expect(loaded.sheets).toHaveLength(TABS);
    for (let t = 0; t < TABS; t++) {
      expect(Object.keys(loaded.sheets[t].cells)).toHaveLength(ROWS * COLS);
      expect(loaded.sheets[t].cells[`${ROWS - 1}:${COLS - 1}`]).toEqual(doc.sheets[t].cells[`${ROWS - 1}:${COLS - 1}`]);
      expect(loaded.sheets[t].cells["5000:1"]).toEqual(doc.sheets[t].cells["5000:1"]);
    }
    expect(net.uploads.length).toBe(uploadsBefore);
  });
});
