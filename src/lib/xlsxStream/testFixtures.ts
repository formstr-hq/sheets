/// <reference types="node" />
// Test helper: xlsx files built with ExcelJS.
import { PassThrough } from "node:stream";
import ExcelJS from "exceljs";

export async function buildFixture(): Promise<Uint8Array> {
  const workbook = new ExcelJS.Workbook();
  const main = workbook.addWorksheet("Main & Data");
  main.columns = [{ width: 12 }, { width: 30 }, { width: 8 }, {}, { width: 20 }];
  main.getRow(1).values = ["Name", "Value", "Flag", "When", "Note"];
  main.getRow(1).font = { bold: true, size: 14, name: "Arial", color: { argb: "FF336699" } };
  main.getRow(1).fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FFFFEE00" } };
  main.getRow(1).height = 30;
  main.addRow(["plain", 1234.5678, true, new Date(Date.UTC(2024, 1, 29)), "caf\u00e9 \u4e2d\u6587 \u{1F600}"]);
  main.addRow(["multi\nline", -0.5, false, new Date(Date.UTC(1999, 11, 31, 13, 30)), "x".repeat(500)]);
  main.getCell("B2").numFmt = "#,##0.00";
  main.getCell("B3").numFmt = "0.0%";
  main.getCell("D2").numFmt = "dd/mm/yyyy";
  main.getCell("A3").alignment = { horizontal: "center", vertical: "top", wrapText: true };
  main.getCell("C3").font = { italic: true, underline: true, strike: true };
  main.getCell("A4").value = { richText: [{ text: "Rich ", font: { bold: true } }, { text: "text" }] };
  main.getCell("B4").value = { formula: "B2*2", result: 2469.1356 };
  main.getCell("C4").value = { formula: "A2&A3", result: "plainmulti\nline" };
  main.getCell("D4").value = { formula: "D2+1", result: new Date(Date.UTC(2024, 2, 1)) };
  main.getCell("E4").value = { error: "#N/A" } as ExcelJS.CellErrorValue;
  main.getCell("A5").value = { text: "link", hyperlink: "https://example.com" };
  main.getCell("B5").value = "5";
  main.getCell("C5").value = "";
  main.getCell("D5").fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FF00FF00" } }; // styled, empty
  main.getCell("A8").value = "after gap";
  main.getCell("H8").value = 42;
  main.getRow(8).height = 45.5;
  main.getRow(6).height = 22; // empty row: the in-memory importer ignores its height
  main.mergeCells("A10:C10");
  main.getCell("A10").value = "merged title";
  main.mergeCells("E12:F14");
  main.getCell("E12").value = 7;
  const validations = (main as unknown as { dataValidations: { add(address: string, rule: unknown): void } }).dataValidations;
  validations.add("B20", { type: "list", allowBlank: true, formulae: ['"a,b,c"'] });
  main.getCell("B20").value = "a";

  const second = workbook.addWorksheet("Second");
  for (let r = 1; r <= 300; r++) second.addRow([r, `row ${r}`, r % 7 === 0 ? null : r * 1.5, r % 2 === 0 ? "even" : "odd"]);

  const hidden = workbook.addWorksheet("Hidden", { state: "hidden" });
  hidden.getCell("B2").value = "secret";

  workbook.addWorksheet("Empty");

  const wide = workbook.addWorksheet("Wide");
  wide.getCell(1, 100).value = "col 100";
  wide.getCell(3, 40).value = 3.14;

  return new Uint8Array(await workbook.xlsx.writeBuffer());
}

export interface LargeOptions {
  tabs: number;
  rows: number;
  cols: number;
  /** Write shared strings (after the sheets, like Excel does) instead of inline strings. */
  sharedStrings: boolean;
}

/** Deterministic cell value for a large fixture: numbers in col 0 mod 3, text in 1, formulas-free ints in 2. */
export function largeValue(tab: number, row: number, col: number): number | string {
  if (row === 0) return `Column ${col + 1}`;
  const n = (tab * 7919 + row * 104729 + col * 1299709) >>> 0;
  if (col % 3 === 0) return (n % 1_000_000) / 100;
  if (col % 3 === 1) return `item-${n.toString(36)}-${row % 50}`;
  return row * (col + 1);
}

/** A multi-tab xlsx written with ExcelJS's streaming writer (strings table after the sheets). */
export async function buildLargeXlsx(options: LargeOptions): Promise<Uint8Array> {
  const stream = new PassThrough();
  const parts: Buffer[] = [];
  stream.on("data", (part: Buffer) => parts.push(part));
  const done = new Promise<void>((resolve, reject) => {
    stream.on("end", resolve);
    stream.on("error", reject);
  });
  const workbook = new ExcelJS.stream.xlsx.WorkbookWriter({ stream, useSharedStrings: options.sharedStrings, useStyles: true });
  for (let tab = 0; tab < options.tabs; tab++) {
    const sheet = workbook.addWorksheet(`Tab ${tab + 1}`);
    for (let row = 0; row <= options.rows; row++) {
      const values = Array.from({ length: options.cols }, (_, col) => largeValue(tab, row, col));
      const added = sheet.addRow(values);
      if (row === 0) added.font = { bold: true };
      added.commit();
    }
    sheet.commit();
  }
  await workbook.commit();
  await done;
  return new Uint8Array(Buffer.concat(parts));
}
