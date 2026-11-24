import { describe, expect, it } from "vitest";
import { base64ToBytes, bytesToBase64, frame, unframe } from "./codec";
import {
  ROOT_MAX_BYTES,
  SheetTooLargeError,
  buildInlineRoot,
  isSheetRootV2,
  parseRootV2,
  serializeRoot,
} from "./root";
import { createEmptyWorkbook, normalizeWorkbook, type WorkbookCell, type WorkbookDocument } from "../workbook";

function workbookWithRows(rows: number, cols = 20): WorkbookDocument {
  const doc = createEmptyWorkbook("Big");
  const cells: Record<string, WorkbookCell> = {};
  for (let r = 0; r < rows; r++) {
    for (let c = 0; c < cols; c++) {
      cells[`${r}:${c}`] = { value: `row ${r} col ${c} some repeated text`, type: "text" };
    }
  }
  doc.sheets[0] = { ...doc.sheets[0], rowCount: rows, colCount: cols, cells };
  return doc;
}

/** Pseudo-random text that deflate cannot shrink much. */
function noisyWorkbook(cellCount: number): WorkbookDocument {
  const doc = createEmptyWorkbook("Noisy");
  let seed = 7;
  const next = () => (seed = (seed * 1664525 + 1013904223) >>> 0);
  const cells: Record<string, WorkbookCell> = {};
  for (let i = 0; i < cellCount; i++) {
    cells[`${i}:0`] = { value: next().toString(36) + next().toString(36) + next().toString(36), type: "text" };
  }
  doc.sheets[0] = { ...doc.sheets[0], rowCount: cellCount, colCount: 1, cells };
  return doc;
}

describe("FSB2 codec", () => {
  it("round-trips JSON through compression", async () => {
    const value = { hello: "world", list: [1, 2, 3], nested: { a: "x".repeat(1000) } };
    const bytes = await frame(value);
    expect(Array.from(bytes.subarray(0, 4))).toEqual([0x46, 0x53, 0x42, 0x32]);
    expect(await unframe(bytes)).toEqual(value);
  });

  it("compresses repetitive data", async () => {
    const value = { text: "abc".repeat(10_000) };
    const bytes = await frame(value);
    expect(bytes.length).toBeLessThan(JSON.stringify(value).length / 10);
  });

  it("produces identical bytes for equal content built in a different key order", async () => {
    const a = { id: "t", cells: { "0:0": { value: 1, type: "number" }, "0:1": { type: "text", value: "x" } } };
    const b = { cells: { "0:1": { value: "x", type: "text" }, "0:0": { type: "number", value: 1 } }, id: "t" };
    expect(Array.from(await frame(a))).toEqual(Array.from(await frame(b)));
  });

  it("rejects data that is not FSB2", async () => {
    await expect(unframe(new Uint8Array([1, 2, 3, 4, 5, 6, 7]))).rejects.toThrow("Not an FSB2 payload");
  });

  it("round-trips base64 for large buffers", () => {
    const bytes = new Uint8Array(100_000).map((_, i) => (i * 31) % 256);
    expect(Array.from(base64ToBytes(bytesToBase64(bytes)))).toEqual(Array.from(bytes));
  });
});

describe("v2 root", () => {
  it("round-trips a workbook with several tabs and layout", async () => {
    const doc = createEmptyWorkbook("Roundtrip");
    doc.sheets[0].cells["0:0"] = {
      value: 42,
      formula: "=A2*2",
      display: "42",
      type: "number",
      style: { bold: true, background: "#FF0000" },
    };
    doc.sheets[0].columnWidths = { 0: 200 };
    doc.sheets[0].merges = [{ r: 1, c: 1, rs: 2, cs: 2 }];
    doc.sheets[0].frozen = { rows: 1, columns: 0 };
    doc.sheets.push({ ...doc.sheets[0], id: "tab2", name: "Second", order: 1, color: "#10B981", cells: { "3:3": { value: "hi" } } });

    const root = await buildInlineRoot(doc);
    expect(isSheetRootV2(root)).toBe(true);
    expect(root.tabs).toHaveLength(2);
    const restored = normalizeWorkbook(await parseRootV2(JSON.parse(serializeRoot(root))));

    expect(restored.title).toBe("Roundtrip");
    expect(restored.sheets.map((s) => s.name)).toEqual(doc.sheets.map((s) => s.name));
    expect(restored.sheets[0].cells["0:0"]).toEqual(doc.sheets[0].cells["0:0"]);
    expect(restored.sheets[0].columnWidths).toEqual({ 0: 200 });
    expect(restored.sheets[0].merges).toEqual(doc.sheets[0].merges);
    expect(restored.sheets[0].frozen).toEqual({ rows: 1, columns: 0 });
    expect(restored.sheets[1].color).toBe("#10B981");
    expect(restored.sheets[1].cells["3:3"]).toEqual({ value: "hi" });
  });

  it("stores a sheet whose plain JSON is several times over the inline limit", async () => {
    const doc = workbookWithRows(400, 10);
    expect(JSON.stringify(doc).length).toBeGreaterThan(ROOT_MAX_BYTES * 4);

    const json = serializeRoot(await buildInlineRoot(doc));
    expect(new TextEncoder().encode(json).length).toBeLessThan(ROOT_MAX_BYTES);
    const restored = await parseRootV2(JSON.parse(json));
    expect(restored.sheets[0].cells["399:9"]).toEqual(doc.sheets[0].cells["399:9"]);
    expect(Object.keys(restored.sheets[0].cells)).toHaveLength(4000);
  });

  it("refuses a sheet that cannot fit even when compressed", async () => {
    const root = await buildInlineRoot(noisyWorkbook(20_000));
    expect(() => serializeRoot(root)).toThrow(SheetTooLargeError);
  });
});
