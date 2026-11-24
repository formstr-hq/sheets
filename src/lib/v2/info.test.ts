import { describe, expect, it } from "vitest";
import { measureWorkbook, encryptedSize } from "./info";
import { uploadFrames } from "./blossom";
import { MAX_CHUNK_BYTES, planTab } from "./chunks";
import { fakeBlossom, localSigner } from "./testUtils";
import { createEmptyWorkbook, type WorkbookCell, type WorkbookDocument } from "../workbook";
import { formatBytes } from "../utils";

function workbook(tabs: number, rows: number, cols = 10): WorkbookDocument {
  const doc = createEmptyWorkbook("Info test");
  const base = doc.sheets[0];
  doc.sheets = Array.from({ length: tabs }, (_, tab) => {
    const cells: Record<string, WorkbookCell> = {};
    for (let r = 0; r < rows; r++) {
      for (let c = 0; c < cols; c++) cells[`${r}:${c}`] = { value: `tab ${tab} row ${r} col ${c} text`, type: "text" };
    }
    return { ...base, id: `t${tab}`, name: `Tab ${tab}`, order: tab, cells };
  });
  doc.activeSheetId = "t0";
  return doc;
}

describe("measureWorkbook", () => {
  it("reports a small sheet as stored inside the event", async () => {
    const info = await measureWorkbook(workbook(1, 5));
    expect(info.storage).toBe("inline");
    expect(info.blobCount).toBe(0);
    expect(info.blobBytes).toBe(0);
    expect(info.cells).toBe(50);
    expect(info.tabs).toHaveLength(1);
    expect(info.eventBytes).toBeGreaterThan(0);
  });

  it("reports a large sheet as Blossom files with a small event", async () => {
    const info = await measureWorkbook(workbook(2, 800));
    expect(info.storage).toBe("blossom");
    expect(info.blobCount).toBe(2);
    expect(info.tabs.map((tab) => tab.cells)).toEqual([8000, 8000]);
    expect(info.blobBytes).toBe(info.tabs[0].storedBytes + info.tabs[1].storedBytes);
    expect(info.eventBytes).toBeLessThan(5_000);
    // Compression is the point: the stored files are much smaller than the plain JSON.
    expect(info.blobBytes).toBeLessThan(info.rawBytes / 3);
  });

  it("predicts the exact size of the encrypted blob that gets uploaded", async () => {
    const doc = workbook(1, 800);
    const info = await measureWorkbook(doc);
    const [plan] = await planTab(doc.sheets[0], { max: MAX_CHUNK_BYTES });
    const frameBytes = plan.frame as Uint8Array;
    const [uploaded] = await uploadFrames([frameBytes], {
      servers: ["https://blossom.example"],
      signer: localSigner(),
      fetch: fakeBlossom().fetch,
    });
    expect(info.tabs[0].storedBytes).toBe(uploaded.ref.n);
  });
});

describe("encryptedSize", () => {
  it("adds one 16-byte tag per segment, and at least one", () => {
    expect(encryptedSize(0, 100)).toBe(16);
    expect(encryptedSize(100, 100)).toBe(116);
    expect(encryptedSize(101, 100)).toBe(133);
  });
});

describe("formatBytes", () => {
  it("formats sizes", () => {
    expect(formatBytes(0)).toBe("0 B");
    expect(formatBytes(1023)).toBe("1023 B");
    expect(formatBytes(1536)).toBe("1.5 KB");
    expect(formatBytes(5 * 1024 * 1024)).toBe("5.0 MB");
    expect(formatBytes(250 * 1024 * 1024)).toBe("250 MB");
    expect(formatBytes(-1)).toBe("—");
  });
});
