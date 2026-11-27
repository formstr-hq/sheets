import { zipSync, strToU8 } from "fflate";
import { describe, expect, it } from "vitest";
import { openZip, ZipError } from "./zip";
import { buildZip, zipBlob } from "./testZip";

async function textOf(gen: AsyncGenerator<string>): Promise<string> {
  let all = "";
  for await (const piece of gen) all += piece;
  return all;
}

const sample = "héllo wörld € 中文 \u{1F600} ".repeat(50);

describe("openZip", () => {
  it("reads stored and deflated entries made by fflate", async () => {
    const bytes = zipSync({
      "a.txt": [strToU8("stored text"), { level: 0 }],
      "dir/b.txt": strToU8(sample),
    });
    const zip = await openZip(new Blob([bytes as BlobPart]));
    expect([...zip.entries.keys()].sort()).toEqual(["a.txt", "dir/b.txt"]);
    expect(await textOf(zip.text("a.txt"))).toBe("stored text");
    expect(await textOf(zip.text("dir/b.txt"))).toBe(sample);
    expect(zip.has("/a.txt")).toBe(true);
    expect(await textOf(zip.text("/dir/b.txt"))).toBe(sample);
  });

  it("keeps multi-byte characters intact at every read boundary", async () => {
    for (const method of [0, 8]) {
      const blob = zipBlob([{ name: "t.xml", data: sample, method }]);
      const zip = await openZip(blob);
      for (let readSize = 1; readSize <= 7; readSize++) {
        expect(await textOf(zip.text("t.xml", { readSize }))).toBe(sample);
      }
    }
  });

  it("handles data-descriptor entries (sizes only in the central directory)", async () => {
    const blob = zipBlob([
      { name: "one.xml", data: sample, descriptor: true },
      { name: "two.xml", data: "second", descriptor: true, method: 0 },
    ]);
    const zip = await openZip(blob);
    expect(await textOf(zip.text("one.xml"))).toBe(sample);
    expect(await textOf(zip.text("two.xml"))).toBe("second");
  });

  it("reads ZIP64 directories", async () => {
    const blob = zipBlob([{ name: "x.xml", data: sample }, { name: "y.xml", data: "why" }], { zip64: true });
    const zip = await openZip(blob);
    expect(zip.entries.get("x.xml")?.usize).toBe(new TextEncoder().encode(sample).length);
    expect(await textOf(zip.text("x.xml"))).toBe(sample);
    expect(await textOf(zip.text("y.xml"))).toBe("why");
  });

  it("finds the directory behind a zip comment", async () => {
    const blob = zipBlob([{ name: "x", data: "ok" }], { comment: "c".repeat(300) });
    expect(await textOf((await openZip(blob)).text("x"))).toBe("ok");
  });

  it("splits long text into several pieces without breaking surrogate pairs", async () => {
    const long = "\u{1F600}".repeat(600_000); // 1.2M UTF-16 units
    const zip = await openZip(zipBlob([{ name: "l", data: long }]));
    const pieces: string[] = [];
    for await (const piece of zip.text("l")) pieces.push(piece);
    expect(pieces.length).toBeGreaterThan(1);
    expect(pieces.join("")).toBe(long);
    for (const piece of pieces) expect(piece.length).toBeLessThanOrEqual(512 * 1024);
  });

  it("reports a missing entry", async () => {
    const zip = await openZip(zipBlob([{ name: "x", data: "1" }]));
    await expect(textOf(zip.text("nope"))).rejects.toMatchObject({ code: "MISSING_ENTRY" });
  });

  it("rejects encrypted and unsupported-method entries when they are read", async () => {
    const zip = await openZip(
      zipBlob([
        { name: "enc", data: "1", encrypted: true },
        { name: "bz", data: "1", method: 12 },
      ]),
    );
    await expect(textOf(zip.text("enc"))).rejects.toMatchObject({ code: "UNSUPPORTED_ZIP" });
    await expect(textOf(zip.text("bz"))).rejects.toMatchObject({ code: "UNSUPPORTED_ZIP" });
  });

  it("rejects files that are not zips or are truncated", async () => {
    await expect(openZip(new Blob(["not a zip file at all, just text"]))).rejects.toMatchObject({ code: "NOT_ZIP" });
    await expect(openZip(new Blob([]))).rejects.toBeInstanceOf(ZipError);
    const good = buildZip([{ name: "x", data: sample }]);
    await expect(openZip(new Blob([good.slice(0, good.length - 40) as BlobPart]))).rejects.toBeInstanceOf(ZipError);
  });

  it("fails on a damaged deflate stream", async () => {
    const good = buildZip([{ name: "x", data: sample }]);
    const bad = good.slice();
    for (let i = 40; i < 60; i++) bad[i] = 0xff;
    const zip = await openZip(new Blob([bad as BlobPart]));
    await expect(textOf(zip.text("x"))).rejects.toMatchObject({ code: "CORRUPT_ZIP" });
  });

  it("stops inflating an entry that outgrows its declared size", async () => {
    const real = "a".repeat(5_000_000);
    const bytes = buildZip([{ name: "bomb", data: real }]);
    // Patch the declared uncompressed size in the central directory down to 10 bytes.
    const view = new DataView(bytes.buffer);
    let at = -1;
    for (let i = bytes.length - 22; i >= 0; i--) {
      if (view.getUint32(i, true) === 0x02014b50) {
        at = i;
        break;
      }
    }
    view.setUint32(at + 24, 10, true);
    const zip = await openZip(new Blob([bytes as BlobPart]));
    await expect(textOf(zip.text("bomb"))).rejects.toMatchObject({ code: "CORRUPT_ZIP" });
  });

  it("stops when the signal is aborted", async () => {
    const controller = new AbortController();
    const zip = await openZip(zipBlob([{ name: "x", data: "z".repeat(2000), method: 0 }]));
    const it = zip.text("x", { readSize: 100, signal: controller.signal });
    await it.next();
    controller.abort();
    await expect(it.next()).rejects.toMatchObject({ code: "CANCELLED" });
  });

  it("reads an empty entry", async () => {
    const zip = await openZip(zipBlob([{ name: "e", data: "" }]));
    expect(await textOf(zip.text("e"))).toBe("");
  });
});
