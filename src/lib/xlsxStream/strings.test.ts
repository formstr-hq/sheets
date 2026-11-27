import { describe, expect, it } from "vitest";
import { loadSharedStrings, PackedStrings } from "./strings";
import { openZip } from "./zip";
import { zipBlob } from "./testZip";

const sst = (inner: string) => `<?xml version="1.0"?><sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">${inner}</sst>`;
const load = async (xml: string, options?: Parameters<typeof loadSharedStrings>[1]) =>
  loadSharedStrings(await openZip(zipBlob([{ name: "xl/sharedStrings.xml", data: xml }])), options);

describe("PackedStrings", () => {
  it("round-trips strings across tiny pages, including unicode and empty", () => {
    const packed = new PackedStrings({ pageBytes: 64 });
    const values = ["", "a", "hello", "héllo wörld", "中文字符", "\u{1F600}\u{1F680}", "x".repeat(40), "y".repeat(300), "", "last"];
    values.forEach((value, i) => expect(packed.add(value)).toBe(i));
    values.forEach((value, i) => expect(packed.get(i)).toBe(value));
    expect(packed.length).toBe(values.length);
    expect(packed.get(99)).toBe("");
    expect(packed.get(-1)).toBe("");
  });

  it("round-trips a million strings", () => {
    const packed = new PackedStrings();
    const make = (i: number) => `value-${i}-${(i * 7919).toString(36)}`;
    for (let i = 0; i < 1_000_000; i++) packed.add(make(i));
    expect(packed.length).toBe(1_000_000);
    for (const i of [0, 1, 999, 65_535, 65_536, 500_000, 999_999]) expect(packed.get(i)).toBe(make(i));
    expect(packed.bytes).toBeLessThan(40 * 1024 * 1024);
  });

  it("aborts cleanly when the budget is exceeded", () => {
    const packed = new PackedStrings({ pageBytes: 1024, budgetBytes: 200 * 1024 });
    const fill = () => {
      for (let i = 0; i < 100_000; i++) packed.add("some text that takes room ".repeat(4));
    };
    expect(fill).toThrowError(/too many distinct text values/);
    try {
      fill();
    } catch (error) {
      expect(error).toMatchObject({ code: "STRING_TABLE_TOO_LARGE" });
    }
  });
});

describe("loadSharedStrings", () => {
  it("decodes entities, flattens rich text, skips phonetic runs", async () => {
    const strings = await load(
      sst(
        "<si><t>Fish &amp; Chips &lt;1&gt;</t></si>" +
          '<si><r><rPr><b/></rPr><t>Bold </t></r><r><t xml:space="preserve">and plain </t></r><r><t>end</t></r></si>' +
          "<si><t>漢字</t><rPh sb=\"0\" eb=\"2\"><t>kanji</t></rPh><phoneticPr fontId=\"1\"/></si>" +
          "<si/>" +
          "<si><t>\u{1F600} café</t></si>" +
          '<si><t xml:space="preserve">  padded  </t></si>',
      ),
    );
    expect(strings).not.toBeNull();
    const all = Array.from({ length: strings!.length }, (_, i) => strings!.get(i));
    expect(all).toEqual(["Fish & Chips <1>", "Bold and plain end", "漢字", "", "\u{1F600} café", "  padded  "]);
  });

  it("returns null when the file has no shared strings", async () => {
    expect(await loadSharedStrings(await openZip(zipBlob([{ name: "x", data: "1" }])))).toBeNull();
  });

  it("works when the table is split across many read slices", async () => {
    const items = Array.from({ length: 5000 }, (_, i) => `<si><t>item € ${i}</t></si>`).join("");
    const zip = await openZip(zipBlob([{ name: "xl/sharedStrings.xml", data: sst(items) }]), { readSize: 7 });
    const strings = await loadSharedStrings(zip);
    expect(strings!.length).toBe(5000);
    expect(strings!.get(4999)).toBe("item € 4999");
  });

  it("aborts with STRING_TABLE_TOO_LARGE past the budget", async () => {
    const items = Array.from({ length: 5000 }, (_, i) => `<si><t>${"z".repeat(100)}${i}</t></si>`).join("");
    await expect(load(sst(items), { budgetBytes: 100 * 1024, pageBytes: 4096 })).rejects.toMatchObject({ code: "STRING_TABLE_TOO_LARGE" });
  });
});
