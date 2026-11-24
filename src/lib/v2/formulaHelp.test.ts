import { describe, expect, it } from "vitest";
import { helpFor, tokenizeFormula } from "./formulaHelp";

describe("formula argument help", () => {
  it("finds the function and the argument the cursor is in", () => {
    expect(helpFor("=SUM(")).toMatchObject({ name: "SUM", current: 0 });
    expect(helpFor("=SUMIF(A1:A5,")).toMatchObject({ name: "SUMIF", current: 1, args: ["range", "criteria", "[sum_range]"] });
    expect(helpFor("=IF(A1>1, SUM(B1:B3), ")).toMatchObject({ name: "IF", current: 2 });
    expect(helpFor("=IF(A1>1, SUM(B1:B3,")).toMatchObject({ name: "SUM", current: 1 });
  });

  it("ignores brackets and commas inside text, and closed calls", () => {
    expect(helpFor('=IF(A1="a,(b", ')).toMatchObject({ name: "IF", current: 1 });
    expect(helpFor("=SUM(A1:A3)")).toBeNull();
    expect(helpFor("=SUM(A1:A3)+ROUND(")).toMatchObject({ name: "ROUND", current: 0 });
    expect(helpFor("=(1+2)*")).toBeNull();
    expect(helpFor("plain text")).toBeNull();
  });

  it("stays on the repeating argument of a list function and shows … for unknown functions", () => {
    expect(helpFor("=SUM(1,2,3,4,5,")).toMatchObject({ name: "SUM", current: 1 });
    expect(helpFor("=SOMETHINGNEW(1,")).toMatchObject({ name: "SOMETHINGNEW", args: ["…"], current: 0 });
    expect(helpFor("=TODAY(")).toMatchObject({ name: "TODAY", args: [] });
  });
});

describe("formula colouring", () => {
  it("splits into references, functions, text, numbers and operators", () => {
    const tokens = tokenizeFormula('SUM(A1:B2)+"hi"*2.5');
    expect(tokens.map((t) => [t.text, t.kind])).toEqual([
      ["SUM", "function"],
      ["(", "operator"],
      ["A1:B2", "ref"],
      [")", "operator"],
      ["+", "operator"],
      ['"hi"', "string"],
      ["*", "operator"],
      ["2.5", "number"],
    ]);
  });

  it("gives each different reference its own colour and the same reference the same colour", () => {
    const refs = tokenizeFormula("A1+B1+$A$1+Sheet2!C3").filter((t) => t.kind === "ref");
    expect(refs.map((t) => t.text)).toEqual(["A1", "B1", "$A$1", "Sheet2!C3"]);
    expect(refs[0].color).toBe(refs[2].color);
    expect(new Set([refs[0].color, refs[1].color, refs[3].color]).size).toBe(3);
  });

  it("does not take function names or words for references", () => {
    const kinds = tokenizeFormula("LOG10(A1)+ABC").map((t) => t.kind);
    expect(kinds[0]).toBe("function");
    expect(tokenizeFormula("LOG10(A1)").find((t) => t.text === "LOG10")?.kind).toBe("function");
    expect(tokenizeFormula("IF(TRUE,1,2)").map((t) => t.text).join("")).toBe("IF(TRUE,1,2)");
  });
});
