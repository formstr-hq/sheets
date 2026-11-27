import { describe, expect, it } from "vitest";
import ExcelJS from "exceljs";
import { importXlsxToBlossom } from "./streamImport";
import { BigSheetStore } from "./bigSheet";
import { argb, cellXml, esc, exportBigSheet, sheetNames, StyleTable, validationXml } from "./bigExport";
import { fakeBlossom, localSigner } from "./testUtils";
import { buildLargeXlsx } from "../xlsxStream/testFixtures";
import { createWorkerHandler } from "../xlsxStream/workerCore";
import type { WorkerLike } from "../xlsxStream/client";
import type { MainToWorker } from "../xlsxStream/protocol";

function inProcessWorker(): WorkerLike {
  const worker: WorkerLike = {
    onmessage: null,
    onerror: null,
    onmessageerror: null,
    postMessage(message: MainToWorker) {
      queueMicrotask(() => handle(message));
    },
    terminate() {
      handle({ t: "cancel" });
    },
  };
  const handle = createWorkerHandler((message) => setTimeout(() => worker.onmessage?.({ data: message }), 0));
  return worker;
}

const START = { maxChunkBytes: 64 * 1024, headRawBytes: 100_000, rawCap: 1_000_000 };

async function open(rows = 3000, tabs = 1) {
  const bytes = await buildLargeXlsx({ tabs, rows, cols: 8, sharedStrings: false });
  const fake = fakeBlossom();
  const signer = localSigner();
  const servers = ["https://export.test"];
  const env = { fetch: fake.fetch, sleep: async () => {} };
  const { root } = await importXlsxToBlossom(new Blob([bytes as BlobPart]), { title: "Export", servers, signer, env, createWorker: inProcessWorker, start: START });
  const store = new BigSheetStore(root, servers, env);
  await store.open();
  return store;
}

async function read(blob: Blob): Promise<ExcelJS.Workbook> {
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.load(await blob.arrayBuffer());
  return workbook;
}

describe("xlsx pieces", () => {
  it("escapes text and picks colours and sheet names", () => {
    expect(esc('a<b>&"c"\u0001')).toBe("a&lt;b&gt;&amp;&quot;c&quot;");
    expect(argb("#ff8800")).toBe("FFFF8800");
    expect(argb("red")).toBeUndefined();
    expect(sheetNames(["Data", "data", "A/B:C", "x".repeat(40), ""])).toEqual(["Data", "data 2", "A_B_C", "x".repeat(31), "Sheet"]);
  });

  it("writes each kind of cell", () => {
    expect(cellXml("A1", { value: 5 }, 0)).toBe('<c r="A1"><v>5</v></c>');
    expect(cellXml("A1", { value: "hi <b>" }, 3)).toBe('<c r="A1" s="3" t="inlineStr"><is><t xml:space="preserve">hi &lt;b&gt;</t></is></c>');
    expect(cellXml("A1", { formula: "SUM(B1:B2)", value: 3 }, 0)).toBe('<c r="A1"><f>SUM(B1:B2)</f><v>3</v></c>');
    expect(cellXml("A1", { formula: "1/0", value: "#DIV/0!", type: "error" }, 0)).toContain('t="e"');
    expect(cellXml("A1", { value: true }, 0)).toContain('t="b"');
    expect(cellXml("A1", {}, 0)).toBeNull();
    expect(cellXml("A1", undefined, 7)).toBe('<c r="A1" s="7"/>');
  });

  it("hands out one style number per distinct look", () => {
    const styles = new StyleTable();
    const bold = styles.id({ bold: true });
    expect(bold).toBeGreaterThan(0);
    expect(styles.id({ bold: true })).toBe(bold);
    expect(styles.id({ bold: true, background: "#ff0000" })).not.toBe(bold);
    expect(styles.id(undefined)).toBe(0);
    expect(styles.xml()).toContain('patternType="solid"');
    expect(styles.xml()).toContain("<b/>");
  });

  it("maps validation rules", () => {
    expect(validationXml("A1", { type: "list", options: ["a", "b"] })).toContain('<formula1>"a,b"</formula1>');
    expect(validationXml("A1", { type: "number", params: { op: "between", a: "1", b: "5" } })).toContain('type="decimal" operator="between"');
    expect(validationXml("A1", { type: "date", params: { op: "gt", a: "2026-01-01", b: "" } })).toContain('type="date" operator="greaterThan"><formula1>46023</formula1>');
    expect(validationXml("A1", { type: "checkbox" })).toBeNull();
  });
});

describe("exporting a big sheet", () => {
  it("writes every row of every tab, and a real xlsx reader sees values, formulas, styles and layout", async () => {
    const store = await open(3000, 2);
    expect(store.tabs[0].chunks.length).toBeGreaterThan(3);
    await store.load(0, 0);
    await store.load(1, 0);
    const last = store.tabs[0].rows - 1;
    await store.load(0, store.chunkFor(0, last));
    await store.edit(0, 2, 1, 2, 1, () => ({ value: "Styled", style: { bold: true, background: "#ffcc00", color: "#112233", numberFormat: "#,##0.00", align: "center", wrap: true } }));
    await store.edit(0, 3, 1, 3, 1, () => ({ value: 1234.5, style: { numberFormat: "#,##0.00" } }));
    await store.edit(0, 4, 1, 4, 1, () => ({ formula: "B4*2", value: 2469 }));
    await store.edit(0, last, 1, last, 1, () => ({ value: "the very last row", link: { url: "https://example.com/x" }, validation: { type: "list", options: ["a", "b"] } }));
    await store.setLayout(0, {
      columnWidths: { 1: 210 },
      rowHeights: { 5: 60 },
      hiddenRows: [7],
      hiddenColumns: [4],
      frozen: { rows: 1, columns: 1 },
      merges: [{ r: 9, c: 0, rs: 2, cs: 2 }],
      borders: [{ r1: 12, c1: 2, r2: 13, c2: 3, type: "outside", color: "#ff0000", style: "medium" }],
      conditionalFormats: [{ ranges: [{ r1: 0, c1: 3, r2: 50, c2: 3 }], condition: { type: "gt", values: ["100"] }, format: { cellColor: "#ffeeee", textColor: "#990000" } }],
    });
    store.setTabColor(0, "#00ff00");
    store.renameTab(1, "Second: sheet");

    const progress: number[] = [];
    const blob = await exportBigSheet(store, { onProgress: (done) => progress.push(done) });
    expect(blob.size).toBeGreaterThan(1000);
    expect(progress.length).toBe(store.tabs[0].chunks.length + store.tabs[1].chunks.length);

    const workbook = await read(blob);
    expect(workbook.worksheets.map((w) => w.name)).toEqual([store.tabs[0].name, "Second_ sheet"]);
    const sheet = workbook.worksheets[0];
    // Rows from the first and the very last chunk are all there.
    expect(sheet.getCell(1, 1).value).toBe(store.cell(0, 0, 0)?.value);
    expect(sheet.getCell(last + 1, 2).value).toMatchObject({ text: "the very last row" });
    expect(sheet.rowCount).toBeGreaterThanOrEqual(last + 1);

    const styled = sheet.getCell(3, 2);
    expect(styled.value).toBe("Styled");
    expect(styled.font?.bold).toBe(true);
    expect(styled.font?.color?.argb).toBe("FF112233");
    expect(styled.fill).toMatchObject({ type: "pattern", fgColor: { argb: "FFFFCC00" } });
    expect(styled.alignment).toMatchObject({ horizontal: "center", wrapText: true });
    expect(sheet.getCell(4, 2).value).toBe(1234.5);
    expect(sheet.getCell(4, 2).numFmt).toBe("#,##0.00");
    expect(sheet.getCell(5, 2).value).toMatchObject({ formula: "B4*2", result: 2469 });
    expect(sheet.getCell(last + 1, 2).dataValidation).toMatchObject({ type: "list", formulae: ['"a,b"'] });

    expect(sheet.getColumn(2).width).toBeCloseTo(210 / 7.5, 1);
    expect(sheet.getColumn(5).hidden).toBe(true);
    expect(sheet.getRow(6).height).toBeCloseTo(45, 1);
    expect(sheet.getRow(8).hidden).toBe(true);
    expect(sheet.views[0]).toMatchObject({ state: "frozen", xSplit: 1, ySplit: 1 });
    expect(Object.keys((sheet as unknown as { _merges: object })._merges)).toContain("A10");
    expect(sheet.properties.tabColor?.argb).toBe("FF00FF00");
    const border = sheet.getCell(13, 3).border;
    expect(border?.left?.style).toBe("medium");
    expect(border?.top?.color?.argb).toBe("FFFF0000");
    expect(sheet.getCell(13, 4).border?.right?.style).toBe("medium");
    expect((sheet as unknown as { conditionalFormattings: unknown[] }).conditionalFormattings.length).toBeGreaterThan(0);
  }, 240_000);

  it("can be cancelled", async () => {
    const store = await open(3000, 1);
    await expect(exportBigSheet(store, { cancelled: () => true })).rejects.toThrow(/cancelled/);
  }, 120_000);
});
