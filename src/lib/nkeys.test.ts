import { describe, expect, it } from "vitest";
import { decodeNKeys, encodeNKeys } from "./nkeys";

describe("nkeys", () => {
  it("round-trips view and edit keys", () => {
    const values = {
      viewKey: "a".repeat(64),
      editKey: "b".repeat(64),
    };
    const encoded = encodeNKeys(values);
    expect(encoded.startsWith("nkeys1")).toBe(true);
    expect(decodeNKeys(encoded)).toEqual(values);
  });

  it("round-trips a single view key", () => {
    const encoded = encodeNKeys({ viewKey: "f".repeat(64) });
    expect(decodeNKeys(encoded)).toEqual({ viewKey: "f".repeat(64) });
  });
});
