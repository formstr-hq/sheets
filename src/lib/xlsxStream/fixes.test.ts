import { describe, expect, it } from "vitest";
import { fortuneSheetsToWorkbook, type WorkbookCell } from "../workbook";
import { parseSpreadsheetFile } from "../xlsx";
import { CREDIT_WINDOW, type MainToWorker, type WorkerToMain } from "./protocol";
import { runImport, type EmittedChunk, type ImportSink } from "./runImport";
import { readSheet } from "./sheetStream";
import { decodeEscapes, loadSharedStrings } from "./strings";
import { buildLargeXlsx } from "./testFixtures";
import { buildZip, zipBlob, type TestEntry } from "./testZip";
import { readWorkbookMeta } from "./workbookMeta";
import { CreditGate, createWorkerHandler } from "./workerCore";
import { openZip } from "./zip";

// Regression tests for the evaluator findings on the streaming import.

const NS = 'xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"';
const REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const STYLES =
  `<styleSheet ${NS}><numFmts count="1"><numFmt numFmtId="164" formatCode="yyyy-mm-dd"/></numFmts>` +
  '<fonts count="1"><font><sz val="11"/><name val="Calibri"/></font></fonts>' +
  '<fills count="2"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill></fills>' +
  '<borders count="1"><border/></borders><cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>' +
  '<cellXfs count="2"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>' +
  '<xf numFmtId="164" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/></cellXfs></styleSheet>';

function mini(sheetXml: string, extra: TestEntry[] = [], workbookPr = ""): TestEntry[] {
  return [
    { name: "[Content_Types].xml", data: '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"/>' },
    { name: "xl/workbook.xml", data: `<workbook ${NS} xmlns:r="${REL}"><workbookPr${workbookPr}/><sheets><sheet name="S" sheetId="1" r:id="rId1"/></sheets></workbook>` },
    { name: "xl/_rels/workbook.xml.rels", data: `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="${REL}/worksheet" Target="worksheets/sheet1.xml"/></Relationships>` },
    { name: "xl/styles.xml", data: STYLES },
    { name: "xl/worksheets/sheet1.xml", data: `<worksheet ${NS}>${sheetXml}</worksheet>` },
    ...extra,
  ];
}

async function streamCells(entries: TestEntry[]) {
  const zip = await openZip(zipBlob(entries));
  const meta = await readWorkbookMeta(zip);
  const strings = await loadSharedStrings(zip);
  const stream = readSheet(zip, meta.sheets[0], { strings, styles: meta.styles, date1904: meta.date1904 });
  const cells: Record<string, WorkbookCell> = {};
  for await (const row of stream.rows()) for (const [col, cell] of row.cells) cells[`${row.row}:${col}`] = cell;
  return { cells, scan: stream.scan() };
}

async function oldCells(entries: TestEntry[]): Promise<Record<string, WorkbookCell>> {
  const bytes = buildZip(entries);
  const imported = await parseSpreadsheetFile(new File([bytes as BlobPart], "x.xlsx"), "t");
  return fortuneSheetsToWorkbook({ title: "t", sheets: imported.sheets }).sheets[0].cells;
}

describe("H1: error values", () => {
  const sheet =
    '<sheetData><row r="1"><c r="A1" t="e"><f>NA()</f><v>#N/A</v></c><c r="B1" t="e"><f>1/0</f><v>#DIV/0!</v></c><c r="C1" t="e"><v>#REF!</v></c></row></sheetData>';

  it("keeps the error text of formula cells (was imported as the number 0)", async () => {
    const { cells } = await streamCells(mini(sheet));
    expect(cells["0:0"]).toMatchObject({ value: "#N/A", display: "#N/A", formula: "=NA()" });
    expect(cells["0:1"]).toMatchObject({ value: "#DIV/0!", formula: "=1/0" });
  });

  it("imports plain error cells as their text instead of [object Object]", async () => {
    const { cells } = await streamCells(mini(sheet));
    expect(cells["0:2"]).toMatchObject({ value: "#REF!", display: "#REF!", type: "text" });
  });

  it("matches the in-memory importer", async () => {
    const entries = mini(sheet);
    const old = await oldCells(entries);
    const { cells } = await streamCells(entries);
    expect(cells).toEqual(old);
    expect(old["0:2"].value).toBe("#REF!");
  });
});

describe("M1/M2: inline strings and empty values", () => {
  it("joins all runs of an inline string and skips phonetic text", async () => {
    const { cells } = await streamCells(
      mini(
        '<sheetData><row r="1"><c r="A1" t="inlineStr"><is><r><t>in</t></r><r><t xml:space="preserve"> line</t></r></is></c>' +
          '<c r="B1" t="inlineStr"><is><t>main</t><rPh sb="0" eb="1"><t>PHON</t></rPh></is></c>' +
          '<c r="C1" t="inlineStr"><is><t>plain</t></is></c></row></sheetData>',
      ),
    );
    expect(cells["0:0"].value).toBe("in line");
    expect(cells["0:1"].value).toBe("main");
    expect(cells["0:2"].value).toBe("plain");
  });

  it("drops cells with an empty value instead of importing NaN", async () => {
    const { cells } = await streamCells(
      mini('<sheetData><row r="1"><c r="A1"><v></v></c><c r="B1"><v/></c><c r="C1" t="s"><v></v></c><c r="D1"><v>3</v></c></row></sheetData>'),
    );
    expect(Object.keys(cells)).toEqual(["0:3"]);
  });
});

describe("M4: rows and cells without references, row order", () => {
  it("numbers rows and cells that have no r attribute", async () => {
    const { cells } = await streamCells(
      mini('<sheetData><row><c><v>1</v></c><c><v>2</v></c></row><row r="5"><c r="C5"><v>3</v></c><c><v>4</v></c></row><row><c><v>5</v></c></row></sheetData>'),
    );
    expect(Object.fromEntries(Object.entries(cells).map(([k, c]) => [k, c.value]))).toEqual({ "0:0": 1, "0:1": 2, "4:2": 3, "4:3": 4, "5:0": 5 });
  });

  it("reports unordered rows with the CORRUPT code and a plain message", async () => {
    await expect(streamCells(mini('<sheetData><row r="5"><c r="A5"><v>1</v></c></row><row r="2"><c r="A2"><v>1</v></c></row></sheetData>'))).rejects.toMatchObject({
      code: "CORRUPT",
      message: expect.stringMatching(/not in order/),
    });
  });
});

describe("M3: damaged files", () => {
  it("maps XML parse failures to CORRUPT with the parser text kept as detail", async () => {
    const noop: ImportSink = { chunk: async () => {}, tabDone: () => {}, progress: () => {} };
    const entries = mini('<sheetData><row r="1"><c r="A1"><v>1</v></c></row></sheetData></oops>');
    const failure = await runImport(zipBlob(entries), noop, { maxChunkBytes: 64 * 1024 }).catch((e) => e);
    expect(failure).toMatchObject({ code: "CORRUPT" });
    expect(failure.message).not.toMatch(/^\d+:\d+:/);
    expect(failure.detail).toMatch(/^\d+:\d+: /);
  });
});

describe("M5: sheets that cannot be imported are reported", () => {
  const wb = (sheets: string, rels: string) => [
    { name: "xl/workbook.xml", data: `<workbook ${NS} xmlns:r="${REL}"><sheets>${sheets}</sheets></workbook>` },
    { name: "xl/_rels/workbook.xml.rels", data: `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${rels}</Relationships>` },
  ];

  it("finds a part whose name differs only in case or uses backslashes", async () => {
    const zip = await openZip(
      zipBlob([
        ...wb(
          '<sheet name="A" sheetId="1" r:id="rId1"/><sheet name="B" sheetId="2" r:id="rId2"/>',
          `<Relationship Id="rId1" Type="${REL}/worksheet" Target="worksheets/SHEET1.xml"/><Relationship Id="rId2" Type="${REL}/worksheet" Target="worksheets\\sheet2.xml"/>`,
        ),
        { name: "xl/worksheets/sheet1.xml", data: "<worksheet/>" },
        { name: "xl/worksheets/sheet2.xml", data: "<worksheet/>" },
      ]),
    );
    const meta = await readWorkbookMeta(zip);
    expect(meta.sheets.map((s) => s.path)).toEqual(["xl/worksheets/sheet1.xml", "xl/worksheets/sheet2.xml"]);
    expect(meta.notes).toEqual([]);
  });

  it("adds a note for every skipped sheet and shows it in the import summary", async () => {
    const entries: TestEntry[] = [
      ...wb(
        '<sheet name="Good" sheetId="1" r:id="rId1"/><sheet name="Chart" sheetId="2" r:id="rId2"/><sheet name="Gone" sheetId="3" r:id="rId3"/><sheet name="NoRel" sheetId="4" r:id="rId9"/>',
        `<Relationship Id="rId1" Type="${REL}/worksheet" Target="worksheets/sheet1.xml"/><Relationship Id="rId2" Type="${REL}/chartsheet" Target="chartsheets/sheet1.xml"/><Relationship Id="rId3" Type="${REL}/worksheet" Target="worksheets/missing.xml"/>`,
      ),
      { name: "xl/worksheets/sheet1.xml", data: `<worksheet ${NS}><sheetData><row r="1"><c r="A1"><v>1</v></c></row></sheetData></worksheet>` },
      { name: "xl/chartsheets/sheet1.xml", data: "<chartsheet/>" },
    ];
    const summary = await runImport(zipBlob(entries), { chunk: async () => {}, tabDone: () => {}, progress: () => {} }, { maxChunkBytes: 64 * 1024 });
    expect(summary.tabs.map((t) => t.name)).toEqual(["Good"]);
    expect(summary.dropped).toHaveLength(3);
    expect(summary.dropped.join("\n")).toMatch(/"Chart".*not a normal worksheet/);
    expect(summary.dropped.join("\n")).toMatch(/"Gone".*missing/);
    expect(summary.dropped.join("\n")).toMatch(/"NoRel".*no data/);
  });
});

describe("L2/L3: dates in 1904 workbooks, _x escapes", () => {
  it("shifts date-formatted formula results by the 1904 epoch like the in-memory importer", async () => {
    const entries = mini('<sheetData><row r="1"><c r="A1" s="1"><f>B1+1</f><v>1</v></c></row></sheetData>', [], ' date1904="1"');
    const old = await oldCells(entries);
    const { cells } = await streamCells(entries);
    expect(cells["0:0"].value).toBe(1463);
    expect(cells).toEqual(old);
  });

  it("leaves 1900-system formula dates as they were", async () => {
    const { cells } = await streamCells(mini('<sheetData><row r="1"><c r="A1" s="1"><f>B1+1</f><v>45000</v></c></row></sheetData>'));
    expect(cells["0:0"].value).toBe(45000);
  });

  it("decodes _xHHHH_ escapes in shared strings", async () => {
    expect(decodeEscapes("a_x000D_b_x005F_x0041_c")).toBe("a\rb_x0041_c");
    const zip = await openZip(zipBlob([{ name: "xl/sharedStrings.xml", data: `<sst ${NS}><si><t>line_x000D__x000A_two</t></si></sst>` }]));
    expect((await loadSharedStrings(zip))!.get(0)).toBe("line\r\ntwo");
  });
});

describe("zip: comments that contain an end-of-directory signature", () => {
  it("finds the real record", async () => {
    const fake = new Uint8Array([0x50, 0x4b, 0x05, 0x06, ...new Array(18).fill(0)]);
    const comment = new TextDecoder("latin1").decode(fake) + " and more text";
    const blob = new Blob([buildZip([{ name: "x", data: "ok" }], { comment }) as BlobPart]);
    const zip = await openZip(blob);
    expect(zip.has("x")).toBe(true);
  });
});

describe("L7: layout sizes are clamped", () => {
  it("limits absurd row numbers and notes it", async () => {
    const { scan } = await streamCells(mini('<sheetData><row r="99999999999"><c r="A99999999999"><v>1</v></c></row></sheetData>'));
    expect(scan.layout.rowCount).toBe(1_048_576);
    expect(scan.notes.join(" ")).toMatch(/limited/);
  });
});

describe("columns and shared formulas", () => {
  it("gives a column without a width ExcelJS's default of 9 characters", async () => {
    const { scan } = await streamCells(mini('<cols><col min="1" max="1"/></cols><sheetData/>'));
    expect(scan.layout.columnWidths).toEqual({ 0: 68 });
  });

  it("skips a shared-formula dependent that has neither formula nor cached value", async () => {
    const { cells, scan } = await streamCells(
      mini(
        '<sheetData><row r="1"><c r="A1"><f t="shared" ref="A1:A3" si="0">1+1</f><v>2</v></c></row><row r="2"><c r="A2"><f t="shared" si="0"/></c></row><row r="3"><c r="A3" t="str"><f t="shared" si="0"/><v>x</v></c></row></sheetData>',
      ),
    );
    expect(cells["1:0"]).toBeUndefined();
    expect(cells["2:0"]).toMatchObject({ value: "x" });
    expect(cells["2:0"].formula).toBeUndefined();
    expect(scan.sharedFormulaValues).toBe(1);
    expect(scan.notes.join(" ")).toMatch(/shared with other cells/);
  });
});

describe("surrogate pairs at text piece boundaries", () => {
  it("never splits a pair when a piece boundary falls inside one", async () => {
    const text = "a" + "\u{1F600}".repeat(400_000); // pairs start on odd indexes, so 512K lands mid-pair
    const zip = await openZip(zipBlob([{ name: "t", data: text }]));
    const pieces: string[] = [];
    for await (const piece of zip.text("t")) pieces.push(piece);
    expect(pieces.length).toBeGreaterThan(1);
    for (const piece of pieces) {
      const first = piece.charCodeAt(0);
      const last = piece.charCodeAt(piece.length - 1);
      expect(first >= 0xdc00 && first <= 0xdfff).toBe(false);
      expect(last >= 0xd800 && last <= 0xdbff).toBe(false);
    }
    expect(pieces.join("")).toBe(text);
  });
});

describe("worker", () => {
  it("applies a learned smaller size limit to later cuts (limit message)", async () => {
    const bytes = await buildLargeXlsx({ tabs: 1, rows: 12000, cols: 10, sharedStrings: true });
    const frames: number[] = [];
    let handle: (message: MainToWorker) => void = () => {};
    const done = new Promise<void>((resolve) => {
      handle = createWorkerHandler((message: WorkerToMain) => {
        if (message.t === "chunk") {
          frames.push(message.chunk.frame.byteLength);
          queueMicrotask(() => handle({ t: "credit", n: 1 }));
        } else if (message.t === "done" || message.t === "error") resolve();
      });
    });
    handle({ t: "start", file: new Blob([bytes as BlobPart]), options: { maxChunkBytes: 512 * 1024, headRawBytes: 50_000, rawCap: 8 * 1024 * 1024 } });
    handle({ t: "limit", maxChunkBytes: 64 * 1024 });
    await done;
    expect(frames.length).toBeGreaterThan(4);
    expect(Math.max(...frames)).toBeLessThanOrEqual(64 * 1024);
  });

  it("a larger limit message does not raise the limit", async () => {
    const bytes = await buildLargeXlsx({ tabs: 1, rows: 6000, cols: 10, sharedStrings: true });
    const frames: number[] = [];
    let handle: (message: MainToWorker) => void = () => {};
    const done = new Promise<void>((resolve) => {
      handle = createWorkerHandler((message: WorkerToMain) => {
        if (message.t === "chunk") {
          frames.push(message.chunk.frame.byteLength);
          queueMicrotask(() => handle({ t: "credit", n: 1 }));
        } else if (message.t === "done" || message.t === "error") resolve();
      });
    });
    handle({ t: "start", file: new Blob([bytes as BlobPart]), options: { maxChunkBytes: 64 * 1024, headRawBytes: 50_000 } });
    handle({ t: "limit", maxChunkBytes: 8 * 1024 * 1024 });
    await done;
    expect(Math.max(...frames)).toBeLessThanOrEqual(64 * 1024);
  });

  it("cancel while the worker waits for a credit ends the import with CANCELLED", async () => {
    const bytes = await buildLargeXlsx({ tabs: 1, rows: 8000, cols: 10, sharedStrings: true });
    const sent: WorkerToMain[] = [];
    let chunks = 0;
    const settled = new Promise<WorkerToMain>((resolve) => {
      const handle = createWorkerHandler((message) => {
        sent.push(message);
        if (message.t === "chunk" && ++chunks === CREDIT_WINDOW) setTimeout(() => handle({ t: "cancel" }), 50); // no credits ever returned
        if (message.t === "error" || message.t === "done") resolve(message);
      });
      handle({ t: "start", file: new Blob([bytes as BlobPart]), options: { maxChunkBytes: 64 * 1024, headRawBytes: 50_000 } });
    });
    expect(await settled).toMatchObject({ t: "error", code: "CANCELLED" });
  });

  it("CreditGate.acquire rejects when aborted while waiting", async () => {
    const gate = new CreditGate(0);
    const controller = new AbortController();
    const waiting = gate.acquire(controller.signal);
    controller.abort();
    await expect(waiting).rejects.toMatchObject({ code: "CANCELLED" });
    gate.release(); // the aborted waiter must not swallow this credit
    await gate.acquire();
  });
});

describe("chunk sink contract", () => {
  it("passes chunks to the sink in order within a tab with the head last", async () => {
    const bytes = await buildLargeXlsx({ tabs: 1, rows: 3000, cols: 10, sharedStrings: false });
    const seen: EmittedChunk[] = [];
    await runImport(new Blob([bytes as BlobPart]), { chunk: async (c) => void seen.push(c), tabDone: () => {}, progress: () => {} }, { maxChunkBytes: 64 * 1024, headRawBytes: 50_000 });
    expect(seen.map((c) => c.seq)).toEqual(seen.map((_, i) => i));
    expect(seen.at(-1)!.isHead).toBe(true);
  });
});
