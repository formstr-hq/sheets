import { describe, expect, it } from "vitest";
import type { Sheet } from "@fortune-sheet/core";
import {
  fortuneSheetsToWorkbook,
  normalizeWorkbook,
  workbookToFortuneSheets,
} from "./adapter";
import { WORKBOOK_SCHEMA } from "./types";

describe("canonical workbook", () => {
  it("normalizes legacy formulas without a leading equals sign", () => {
    const doc = normalizeWorkbook({
      schema: WORKBOOK_SCHEMA,
      version: 1,
      title: "Test",
      activeSheetId: "sheet",
      createdAt: 1,
      updatedAt: 1,
      sheets: [
        {
          id: "sheet",
          name: "Sheet 1",
          order: 0,
          rowCount: 200,
          colCount: 26,
          cells: {
            "0:2": { value: 5, formula: "A1+B1", type: "number" },
          },
        },
      ],
    });
    expect(doc.sheets[0].cells["0:2"].formula).toBe("=A1+B1");
  });

  it("adapts FortuneSheet data into canonical cells and back", () => {
    const sheet: Sheet = {
      id: "sheet",
      name: "Sheet 1",
      status: 1,
      order: 0,
      row: 200,
      column: 26,
      celldata: [
        {
          r: 0,
          c: 0,
          v: {
            v: "Done",
            m: "Done",
            ct: { fa: "General", t: "s", s: { fs: "tag" } },
            bl: 1,
            bg: "#ffeeaa",
          },
        },
        {
          r: 1,
          c: 1,
          v: { v: 3, f: "=SUM(A1:A2)", m: "3", ct: { fa: "0.00", t: "n" } },
        },
        {
          r: 2,
          c: 2,
          v: {
            v: "Ada",
            m: "Ada",
            ct: {
              fa: "General",
              t: "s",
              s: { fs: "user", pubkey: "a".repeat(64), npub: "npub1test" },
            },
          },
        },
      ],
      config: { columnlen: { 0: 160 }, rowlen: { 0: 40 }, merge: {} },
    };
    const doc = fortuneSheetsToWorkbook({ title: "Test", sheets: [sheet] });
    expect(doc.schema).toBe(WORKBOOK_SCHEMA);
    expect(doc.sheets[0].cells["0:0"].tag).toBe("Done");
    expect(doc.sheets[0].cells["0:0"].style?.bold).toBe(true);
    expect(doc.sheets[0].cells["1:1"].formula).toBe("=SUM(A1:A2)");
    expect(doc.sheets[0].cells["2:2"].person?.pubkey).toBe("a".repeat(64));

    const restored = workbookToFortuneSheets(doc)[0];
    const tag = restored.celldata?.find((item) => item.r === 0 && item.c === 0)?.v;
    const formula = restored.celldata?.find((item) => item.r === 1 && item.c === 1)?.v;
    expect((tag as { ct?: { s?: { fs?: string } } }).ct?.s?.fs).toBe("tag");
    expect(formula?.f).toBe("=SUM(A1:A2)");
    expect(restored.config?.columnlen?.[0]).toBe(160);
  });

  it("detects a legacy FortuneSheet document and upgrades it", () => {
    const doc = normalizeWorkbook({
      version: 1,
      title: "Legacy",
      createdAt: 1,
      updatedAt: 2,
      sheets: [
        {
          id: "legacy",
          name: "Sheet 1",
          row: 200,
          column: 26,
          status: 1,
          celldata: [{ r: 0, c: 0, v: { v: "hello", m: "hello" } }],
        },
      ],
    });
    expect(doc.schema).toBe(WORKBOOK_SCHEMA);
    expect(doc.activeSheetId).toBe("legacy");
    expect(doc.sheets[0].cells["0:0"].value).toBe("hello");
  });
});

describe("per-cell sheet data keys (FortuneSheet uses row_col)", () => {
  const base = (cells: Record<string, unknown>) =>
    normalizeWorkbook({
      schema: WORKBOOK_SCHEMA,
      version: 1,
      title: "Keys",
      activeSheetId: "s",
      createdAt: 1,
      updatedAt: 1,
      sheets: [{ id: "s", name: "S", order: 0, rowCount: 200, colCount: 26, cells }],
    });

  it("writes validations and links under row_col so FortuneSheet reads them", () => {
    const doc = base({
      "2:3": { value: "Done", validation: { type: "dropdown", options: ["Todo", "Done"] } },
      "4:1": { value: "site", link: { url: "https://example.com" } },
    });
    const [sheet] = workbookToFortuneSheets(doc);
    expect(Object.keys(sheet.dataVerification ?? {})).toEqual(["2_3"]);
    expect(Object.keys((sheet as unknown as { hyperlink: object }).hyperlink)).toEqual(["4_1"]);
  });

  it("round-trips validations and links through FortuneSheet's own key format", () => {
    const doc = base({
      "2:3": { value: "Done", validation: { type: "dropdown", options: ["Todo", "Done"] } },
      "4:1": { value: "site", link: { url: "https://example.com" } },
    });
    const back = fortuneSheetsToWorkbook({ title: "Keys", sheets: workbookToFortuneSheets(doc) });
    expect(back.sheets[0].cells["2:3"].validation?.options).toEqual(["Todo", "Done"]);
    expect(back.sheets[0].cells["4:1"].link?.url).toBe("https://example.com");
  });

  it("keeps a link on an empty cell and still reads documents saved with row:col keys", () => {
    const sheets = [
      {
        id: "s",
        name: "S",
        order: 0,
        status: 1,
        row: 200,
        column: 26,
        celldata: [{ r: 1, c: 1, v: { v: "x", m: "x" } }],
        hyperlink: { "5_5": { linkType: "webpage", linkAddress: "https://a.test" } },
        dataVerification: { "1:1": { type: "dropdown", value1: "a,b" } },
      },
    ] as unknown as Sheet[];
    const doc = fortuneSheetsToWorkbook({ title: "Legacy", sheets });
    expect(doc.sheets[0].cells["5:5"]?.link?.url).toBe("https://a.test");
    expect(doc.sheets[0].cells["1:1"]?.validation?.options).toEqual(["a", "b"]);
  });
});
