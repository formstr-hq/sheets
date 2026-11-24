/// <reference types="node" />
import { existsSync, readFileSync } from "node:fs";
import { describe, expect, it } from "vitest";
import { fortuneSheetsToWorkbook, type WorkbookCell, type WorkbookSheet } from "../workbook";
import { parseSpreadsheetFile } from "../xlsx";
import { buildFixture } from "./testFixtures";
import { loadSharedStrings } from "./strings";
import { readSheet } from "./sheetStream";
import { readWorkbookMeta } from "./workbookMeta";
import { openZip } from "./zip";

// The streaming reader against the in-memory importer: the same file must produce the same cells,
// widths, heights, merges and sizes. Documented differences are removed before comparing:
//  - data validations (the stream path does not read them),
//  - hidden state (only the stream path records it).

interface Imported {
  name: string;
  hidden: boolean;
  cells: Record<string, WorkbookCell>;
  layout: Record<string, unknown>;
}

async function streamImport(blob: Blob): Promise<Imported[]> {
  const zip = await openZip(blob);
  const meta = await readWorkbookMeta(zip);
  const strings = await loadSharedStrings(zip);
  const out: Imported[] = [];
  for (const ref of meta.sheets) {
    const stream = readSheet(zip, ref, { strings, styles: meta.styles, date1904: meta.date1904 });
    const cells: Record<string, WorkbookCell> = {};
    for await (const row of stream.rows()) for (const [col, cell] of row.cells) cells[`${row.row}:${col}`] = cell;
    out.push({ name: ref.name, hidden: ref.hidden, cells, layout: stream.scan().layout });
  }
  return out;
}

async function oldImport(bytes: Uint8Array): Promise<WorkbookSheet[]> {
  const imported = await parseSpreadsheetFile(new File([bytes as BlobPart], "fixture.xlsx"), "t");
  return fortuneSheetsToWorkbook({ title: "t", sheets: imported.sheets }).sheets;
}

function withoutValidation(cells: Record<string, WorkbookCell>): Record<string, WorkbookCell> {
  const out: Record<string, WorkbookCell> = {};
  for (const [key, cell] of Object.entries(cells)) {
    const { validation: _validation, ...rest } = cell;
    void _validation;
    out[key] = rest;
  }
  return out;
}

const LAYOUT_KEYS = ["rowCount", "colCount", "defaultRowHeight", "defaultColWidth", "columnWidths", "rowHeights", "merges"] as const;

async function expectParity(bytes: Uint8Array, options: { ignoreFormulaText?: boolean } = {}): Promise<{ cells: number; diffs: string[] }> {
  const [streamed, old] = await Promise.all([streamImport(new Blob([bytes as BlobPart])), oldImport(bytes)]);
  expect(streamed.map((s) => s.name)).toEqual(old.map((s) => s.name));
  let cells = 0;
  const diffs: string[] = [];
  streamed.forEach((sheet, i) => {
    const expected = withoutValidation(old[i].cells);
    const keys = new Set([...Object.keys(expected), ...Object.keys(sheet.cells)]);
    for (const key of keys) {
      cells++;
      const a = sheet.cells[key];
      const b = expected[key];
      if (JSON.stringify(a) !== JSON.stringify(b)) {
        if (options.ignoreFormulaText && a && b && a.formula === undefined && b.formula !== undefined) continue;
        diffs.push(`${sheet.name}!${key}\n  stream ${JSON.stringify(a)}\n  old    ${JSON.stringify(b)}`);
      }
    }
    for (const field of LAYOUT_KEYS) {
      const a = (sheet.layout as Record<string, unknown>)[field];
      const b = (old[i] as unknown as Record<string, unknown>)[field];
      if (JSON.stringify(a) !== JSON.stringify(b)) diffs.push(`${sheet.name} layout.${field}\n  stream ${JSON.stringify(a)}\n  old    ${JSON.stringify(b)}`);
    }
  });
  return { cells, diffs };
}

describe("streaming import parity with the in-memory importer", () => {
  it("matches on a fixture with strings, dates, styles, formulas, merges, widths and heights", async () => {
    const bytes = await buildFixture();
    const { cells, diffs } = await expectParity(bytes);
    expect(cells).toBeGreaterThan(1000);
    expect(diffs).toEqual([]);
  });

  it("records hidden sheets (the in-memory importer does not)", async () => {
    const streamed = await streamImport(new Blob([(await buildFixture()) as BlobPart]));
    expect(streamed.map((s) => [s.name, s.hidden])).toEqual([
      ["Main & Data", false],
      ["Second", false],
      ["Hidden", true],
      ["Empty", false],
      ["Wide", false],
    ]);
  });

  it("reads merges, widths and heights into the layout", async () => {
    const streamed = await streamImport(new Blob([(await buildFixture()) as BlobPart]));
    const layout = streamed[0].layout as Record<string, unknown>;
    expect(layout.merges).toEqual([
      { r: 9, c: 0, rs: 1, cs: 3 },
      { r: 11, c: 4, rs: 3, cs: 2 },
    ]);
    expect(layout.columnWidths).toMatchObject({ 0: 90, 1: 225 });
    expect(layout.rowHeights).toMatchObject({ 0: 40, 7: 61 });
    expect(layout.rowHeights).not.toHaveProperty("5"); // empty row, like the old importer
  });
});

const DOWNLOADS = "C:/Users/Ido69/Downloads";
const REAL = [
  "Haryana Real Estate Regulatory Authority, Gurugram.xlsx",
  "SRM STEP MAJOR ASSESSMENT 1 - STUDENTS SCHEDULE.xlsx",
  "dcbd7181-f489-46b3-92a1-0af995a383a4.xlsx",
];

// Read-only use of files from the developer's Downloads folder; skipped where they do not exist.
describe("streaming import parity on real files", () => {
  for (const name of REAL) {
    it.skipIf(!existsSync(`${DOWNLOADS}/${name}`))(name, async () => {
      const bytes = new Uint8Array(readFileSync(`${DOWNLOADS}/${name}`));
      // Shared-formula dependents: the old importer translates them, the stream path keeps values.
      const { cells, diffs } = await expectParity(bytes, { ignoreFormulaText: true });
      expect(cells).toBeGreaterThan(0);
      expect(diffs.slice(0, 5)).toEqual([]);
    });
  }
});
