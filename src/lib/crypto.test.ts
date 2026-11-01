import { describe, expect, it } from "vitest";
import { buildSheetEvent, decryptWithSecret, encryptWithSecret, generateDocKeys } from "./crypto";
import { createEmptyWorkbook } from "./workbook";

describe("sheet crypto", () => {
  it("encrypts and decrypts with a self-derived conversation key", () => {
    const keys = generateDocKeys();
    const ciphertext = encryptWithSecret("private cell contents", keys.viewKey);
    expect(ciphertext).not.toContain("private");
    expect(decryptWithSecret(ciphertext, keys.viewKey)).toBe("private cell contents");
  });

  it("signs sheet events with the edit key and encrypts with the view key", () => {
    const keys = generateDocKeys();
    const doc = createEmptyWorkbook("Encrypted test");
    const event = buildSheetEvent({
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
    const decrypted = JSON.parse(decryptWithSecret(event.content, keys.viewKey)) as {
      title: string;
    };
    expect(decrypted.title).toBe("Encrypted test");
  });
});
