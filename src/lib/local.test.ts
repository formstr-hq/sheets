import { afterEach, beforeEach, describe, expect, it } from "vitest";
import { loadDocCache, saveDocCache } from "./local";
import { createEmptyWorkbook } from "./workbook";
import type { LibraryEntry, SheetDocument } from "./types";

/** A localStorage with a small quota, like a browser's ~5 MB one. */
function limitedStorage(quotaChars: number): Storage {
  const data = new Map<string, string>();
  return {
    get length() {
      return data.size;
    },
    clear: () => data.clear(),
    key: (index) => Array.from(data.keys())[index] ?? null,
    getItem: (key) => data.get(key) ?? null,
    removeItem: (key) => void data.delete(key),
    setItem: (key, value) => {
      if (value.length > quotaChars) throw new DOMException("quota exceeded", "QuotaExceededError");
      data.set(key, value);
    },
  };
}

const entry = { coordinate: "33460:abc:sheet" } as LibraryEntry;

function docWithText(text: string): SheetDocument {
  const doc = createEmptyWorkbook("Cache test");
  doc.sheets[0].cells["0:0"] = { value: text, type: "text" };
  return doc;
}

describe("device copy of a sheet", () => {
  beforeEach(() => {
    globalThis.localStorage = limitedStorage(20_000);
  });
  afterEach(() => {
    delete (globalThis as { localStorage?: Storage }).localStorage;
  });

  it("stores and reads back a sheet that fits", () => {
    expect(saveDocCache(entry, docWithText("hello"), { pending: true })).toBe(true);
    const cached = loadDocCache(entry);
    expect(cached?.pending).toBe(true);
    expect(cached?.doc.sheets[0].cells["0:0"]?.value).toBe("hello");
  });

  it("drops an older unsent copy when a newer, larger one cannot be stored", () => {
    saveDocCache(entry, docWithText("old version"), { pending: true });
    // The sheet grew past the quota: the newer copy cannot be stored.
    expect(saveDocCache(entry, docWithText("x".repeat(50_000)), { pending: true })).toBe(false);
    // The stale "old version" must not survive to be loaded instead of the newer data.
    expect(loadDocCache(entry)).toBeNull();
  });
});
