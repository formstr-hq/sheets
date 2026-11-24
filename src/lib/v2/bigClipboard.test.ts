import { describe, expect, it } from "vitest";
import { fromTsv, toTsv } from "./bigClipboard";

describe("tab-separated text", () => {
  it("writes rows and quotes cells that hold tabs, line breaks or quotes", () => {
    expect(toTsv([["a", "b"], ["1", ""]])).toBe("a\tb\n1\t");
    expect(toTsv([["say \"hi\"", "two\nlines", "tab\there"]])).toBe('"say ""hi"""\t"two\nlines"\t"tab\there"');
  });

  it("reads it back, including quoted cells", () => {
    const rows = [["a", ""], ["say \"hi\"", "two\nlines"], ["x", "y"]];
    expect(fromTsv(toTsv(rows))).toEqual(rows);
  });

  it("accepts Windows line endings, a trailing newline and empty cells", () => {
    expect(fromTsv("1\t2\r\n3\t\r\n")).toEqual([["1", "2"], ["3", ""]]);
    expect(fromTsv("")).toEqual([]);
    expect(fromTsv("only")).toEqual([["only"]]);
  });
});
