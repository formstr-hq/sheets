import { describe, expect, it } from "vitest";
import { finalizeEvent } from "nostr-tools";
import { hexToBytes } from "nostr-tools/utils";
import {
  buildSheetEvent,
  decodeSheetEvent,
  decryptWithSecret,
  encryptWithSecret,
  generateDocKeys,
} from "./crypto";
import { createEmptyWorkbook, type WorkbookCell, type WorkbookDocument } from "./workbook";
import { SheetTooLargeError } from "./v2/root";
import { fakeBlossom, localSigner } from "./v2/testUtils";

const SERVER = "https://blossom.example";

/** A workbook whose tabs are each too big to share one inline event. */
function bigWorkbook(tabs: number, rows = 800, cols = 10): WorkbookDocument {
  const doc = createEmptyWorkbook("Big workbook");
  const base = doc.sheets[0];
  doc.sheets = Array.from({ length: tabs }, (_, tab) => {
    const cells: Record<string, WorkbookCell> = {};
    for (let r = 0; r < rows; r++) {
      for (let c = 0; c < cols; c++) cells[`${r}:${c}`] = { value: `tab ${tab} row ${r} col ${c} text`, type: "text" };
    }
    // Counts at the editor's minimums, so loading (which normalizes) does not change the tab.
    return {
      ...base,
      id: `tab${tab}`,
      name: `Tab ${tab}`,
      order: tab,
      rowCount: Math.max(rows, 200),
      colCount: Math.max(cols, 26),
      cells,
    };
  });
  doc.activeSheetId = "tab0";
  return doc;
}

describe("sheet crypto", () => {
  it("encrypts and decrypts with a self-derived conversation key", () => {
    const keys = generateDocKeys();
    const ciphertext = encryptWithSecret("private cell contents", keys.viewKey);
    expect(ciphertext).not.toContain("private");
    expect(decryptWithSecret(ciphertext, keys.viewKey)).toBe("private cell contents");
  });

  it("signs sheet events with the edit key and encrypts with the view key", async () => {
    const keys = generateDocKeys();
    const doc = createEmptyWorkbook("Encrypted test");
    const event = await buildSheetEvent({
      sheetId: "test",
      document: doc,
      viewKey: keys.viewKey,
      editKey: keys.editKey,
      createdAt: 1,
    });
    expect(event.pubkey).toBe(keys.editPubkey);
    expect(event.kind).toBe(33460);
    expect(event.tags).toContainEqual(["d", "test"]);
    expect(event.tags).toContainEqual(["encrypted", "nip44"]);
    expect(event.tags).toContainEqual(["v", "2"]);
    expect(event.tags).toContainEqual(["storage", "inline"]);
    const decrypted = JSON.parse(decryptWithSecret(event.content, keys.viewKey)) as {
      title: string;
      v: number;
    };
    expect(decrypted.title).toBe("Encrypted test");
    expect(decrypted.v).toBe(2);
    expect((await decodeSheetEvent(event, keys.viewKey)).title).toBe("Encrypted test");
  });

  it("still reads v1 events written before the v2 format", async () => {
    const keys = generateDocKeys();
    const doc = createEmptyWorkbook("Old sheet");
    doc.sheets[0].cells["0:0"] = { value: "legacy", type: "text" };
    // Exactly what the previous buildSheetEvent produced: the workbook JSON as the content.
    const v1Event = finalizeEvent(
      {
        kind: 33460,
        created_at: 1,
        tags: [
          ["d", "old"],
          ["t", "formstr-sheet"],
          ["client", "formstr-sheets"],
          ["encrypted", "nip44"],
        ],
        content: encryptWithSecret(JSON.stringify(doc), keys.viewKey),
      },
      hexToBytes(keys.editKey),
    );
    const decoded = await decodeSheetEvent(v1Event, keys.viewKey);
    expect(decoded.title).toBe("Old sheet");
    expect(decoded.sheets[0].cells["0:0"]).toEqual({ value: "legacy", type: "text" });
  });

  it("still reads v1 events in the older fortune-sheet format", async () => {
    const keys = generateDocKeys();
    const fortuneStyle = {
      title: "Fortune era",
      sheets: [{ id: "s1", name: "Sheet1", status: 1, order: 0, celldata: [{ r: 0, c: 0, v: { v: "hi", m: "hi", ct: { fa: "General", t: "g" } } }] }],
    };
    const event = finalizeEvent(
      {
        kind: 33460,
        created_at: 1,
        tags: [["d", "fortune"]],
        content: encryptWithSecret(JSON.stringify(fortuneStyle), keys.viewKey),
      },
      hexToBytes(keys.editKey),
    );
    const decoded = await decodeSheetEvent(event, keys.viewKey);
    expect(decoded.title).toBe("Fortune era");
    expect(decoded.sheets[0].cells["0:0"]?.value).toBe("hi");
  });

  it("fits a sheet into the relay content limit that the v1 format exceeded", async () => {
    // relay.formstr.app advertises max_content_length 102400 in its NIP-11 document.
    const RELAY_CONTENT_LIMIT = 102_400;
    const keys = generateDocKeys();
    const doc = createEmptyWorkbook("Large");
    const cells: Record<string, { value: string; type: "text" }> = {};
    for (let r = 0; r < 400; r++) {
      for (let c = 0; c < 10; c++) cells[`${r}:${c}`] = { value: `row ${r} col ${c} some repeated text`, type: "text" };
    }
    doc.sheets[0] = { ...doc.sheets[0], rowCount: 400, colCount: 10, cells };

    // What the v1 format published: the workbook JSON encrypted directly.
    const v1Content = encryptWithSecret(JSON.stringify(doc), keys.viewKey);
    expect(v1Content.length).toBeGreaterThan(RELAY_CONTENT_LIMIT);

    const event = await buildSheetEvent({ sheetId: "large", document: doc, viewKey: keys.viewKey, editKey: keys.editKey });
    expect(event.content.length).toBeLessThan(RELAY_CONTENT_LIMIT);
    const decoded = await decodeSheetEvent(event, keys.viewKey);
    expect(Object.keys(decoded.sheets[0].cells)).toHaveLength(4000);
  });
});

describe("sheet storage on Blossom", () => {
  const keys = generateDocKeys();

  it("keeps a small sheet inline and never contacts Blossom", async () => {
    const net = fakeBlossom();
    const event = await buildSheetEvent({
      sheetId: "small",
      document: createEmptyWorkbook("Small"),
      viewKey: keys.viewKey,
      editKey: keys.editKey,
      blossom: { servers: [SERVER], signer: localSigner(), env: { fetch: net.fetch } },
    });
    expect(event.tags).toContainEqual(["storage", "inline"]);
    expect(net.uploads).toHaveLength(0);
  });

  it("stores a large sheet on Blossom and leaves only a small pointer in the event", async () => {
    const net = fakeBlossom();
    const doc = bigWorkbook(1);
    const event = await buildSheetEvent({
      sheetId: "big",
      document: doc,
      viewKey: keys.viewKey,
      editKey: keys.editKey,
      blossom: { servers: [SERVER], signer: localSigner(), env: { fetch: net.fetch } },
    });

    expect(event.tags).toContainEqual(["v", "2"]);
    expect(event.tags).toContainEqual(["storage", "blossom"]);
    expect(event.content.length).toBeLessThan(5_000);
    expect(net.uploads).toHaveLength(1);
    // The cell text is nowhere in what the relay or the Blossom server sees.
    expect(event.content).not.toContain("tab 0 row");
    const blob = net.stores.get(SERVER)!.values().next().value!;
    expect(new TextDecoder("latin1").decode(blob)).not.toContain("tab 0 row");
  });

  it("loads that sheet on a fresh device that only has the event and the view key", async () => {
    const net = fakeBlossom();
    const doc = bigWorkbook(2);
    const event = await buildSheetEvent({
      sheetId: "big",
      document: doc,
      viewKey: keys.viewKey,
      editKey: keys.editKey,
      blossom: { servers: [SERVER], signer: localSigner(), env: { fetch: net.fetch } },
    });

    // No server list of its own: the root names where its blobs live.
    const loaded = await decodeSheetEvent(event, keys.viewKey, undefined, { servers: [], env: { fetch: net.fetch } });
    expect(loaded.title).toBe("Big workbook");
    expect(loaded.sheets.map((sheet) => sheet.name)).toEqual(["Tab 0", "Tab 1"]);
    expect(Object.keys(loaded.sheets[1].cells)).toHaveLength(8000);
    expect(loaded.sheets[1].cells["799:9"]).toEqual(doc.sheets[1].cells["799:9"]);
  });

  it("uploads only the tabs that changed since the last save", async () => {
    const net = fakeBlossom();
    const reuse = new Map();
    const blossom = { servers: [SERVER], signer: localSigner(), reuse, env: { fetch: net.fetch } };
    const doc = bigWorkbook(3);
    const params = { sheetId: "big", viewKey: keys.viewKey, editKey: keys.editKey, blossom };

    await buildSheetEvent({ ...params, document: doc });
    expect(net.uploads).toHaveLength(3);

    const edited: WorkbookDocument = {
      ...doc,
      sheets: doc.sheets.map((sheet, i) =>
        i === 1 ? { ...sheet, cells: { ...sheet.cells, "0:0": { value: "edited", type: "text" as const } } } : sheet,
      ),
    };
    const event = await buildSheetEvent({ ...params, document: edited });
    expect(net.uploads).toHaveLength(4);

    const loaded = await decodeSheetEvent(event, keys.viewKey, undefined, { servers: [], env: { fetch: net.fetch } });
    expect(loaded.sheets[1].cells["0:0"]?.value).toBe("edited");
    expect(loaded.sheets[0].cells["0:0"]?.value).toBe("tab 0 row 0 col 0 text");
  });

  it("knows what is already uploaded after loading, so the first save does not resend everything", async () => {
    const net = fakeBlossom();
    const doc = bigWorkbook(2);
    const event = await buildSheetEvent({
      sheetId: "big",
      document: doc,
      viewKey: keys.viewKey,
      editKey: keys.editKey,
      blossom: { servers: [SERVER], signer: localSigner(), env: { fetch: net.fetch } },
    });
    expect(net.uploads).toHaveLength(2);

    const reuse = new Map();
    const loaded = await decodeSheetEvent(event, keys.viewKey, undefined, { servers: [], reuse, env: { fetch: net.fetch } });
    await buildSheetEvent({
      sheetId: "big",
      document: loaded,
      viewKey: keys.viewKey,
      editKey: keys.editKey,
      blossom: { servers: [SERVER], signer: localSigner(), reuse, env: { fetch: net.fetch } },
    });
    expect(net.uploads).toHaveLength(2);
  });

  it("falls back to the reader's own servers when the root's servers are gone", async () => {
    const net = fakeBlossom();
    const event = await buildSheetEvent({
      sheetId: "big",
      document: bigWorkbook(1),
      viewKey: keys.viewKey,
      editKey: keys.editKey,
      blossom: { servers: [SERVER], signer: localSigner(), env: { fetch: net.fetch } },
    });
    const mirror = "https://mirror.example";
    net.stores.set(mirror, new Map(net.stores.get(SERVER)!));
    net.down.add(SERVER);
    const loaded = await decodeSheetEvent(event, keys.viewKey, undefined, { servers: [mirror], env: { fetch: net.fetch } });
    expect(loaded.sheets[0].name).toBe("Tab 0");
  });

  it("refuses to save a sheet that needs Blossom when the user has no server", async () => {
    const net = fakeBlossom();
    const params = { sheetId: "big", document: bigWorkbook(1), viewKey: keys.viewKey, editKey: keys.editKey };
    await expect(
      buildSheetEvent({ ...params, blossom: { servers: [], signer: localSigner(), env: { fetch: net.fetch } } }),
    ).rejects.toThrow(SheetTooLargeError);
    await expect(buildSheetEvent(params)).rejects.toThrow(SheetTooLargeError);
    expect(net.uploads).toHaveLength(0);
  });

  it("turns a sheet that v1 could not publish into an event well under the relay limit", async () => {
    const net = fakeBlossom();
    const doc = bigWorkbook(1);
    // v1 stored this JSON directly in the event; it is far over any relay content limit.
    expect(encryptWithSecret(JSON.stringify(doc), keys.viewKey).length).toBeGreaterThan(102_400);
    const event = await buildSheetEvent({
      sheetId: "upgrade",
      document: doc,
      viewKey: keys.viewKey,
      editKey: keys.editKey,
      blossom: { servers: [SERVER], signer: localSigner(), env: { fetch: net.fetch } },
    });
    expect(event.content.length).toBeLessThan(102_400);
  });
});

