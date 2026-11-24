import { describe, expect, it } from "vitest";
import type { WorkbookCell } from "../workbook";
import { loadSharedStrings } from "./strings";
import { readSheet } from "./sheetStream";
import { readWorkbookMeta } from "./workbookMeta";
import { openZip } from "./zip";
import { zipBlob, type TestEntry } from "./testZip";

const NS = 'xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"';
const REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";

const STYLES =
  `<styleSheet ${NS}><numFmts count="1"><numFmt numFmtId="164" formatCode="yyyy-mm-dd"/></numFmts>` +
  '<fonts count="2"><font><sz val="11"/><name val="Calibri"/></font><font><b/><sz val="11"/><name val="Calibri"/></font></fonts>' +
  '<fills count="2"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill></fills>' +
  '<borders count="1"><border/></borders><cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>' +
  '<cellXfs count="3"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>' +
  '<xf numFmtId="164" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/>' +
  '<xf numFmtId="0" fontId="1" fillId="0" borderId="0" xfId="0" applyFont="1"/></cellXfs></styleSheet>';

function mini(sheetXml: string, extra: TestEntry[] = [], workbookPr = ""): TestEntry[] {
  return [
    { name: "xl/workbook.xml", data: `<workbook ${NS} xmlns:r="${REL}"><workbookPr${workbookPr}/><sheets><sheet name="S" sheetId="1" r:id="rId1"/></sheets></workbook>` },
    { name: "xl/_rels/workbook.xml.rels", data: `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="${REL}/worksheet" Target="worksheets/sheet1.xml"/></Relationships>` },
    { name: "xl/styles.xml", data: STYLES },
    { name: "xl/worksheets/sheet1.xml", data: `<worksheet ${NS}>${sheetXml}</worksheet>` },
    ...extra,
  ];
}

async function readAll(entries: TestEntry[]): Promise<Record<string, WorkbookCell>> {
  const zip = await openZip(zipBlob(entries));
  const meta = await readWorkbookMeta(zip);
  const strings = await loadSharedStrings(zip);
  const stream = readSheet(zip, meta.sheets[0], { strings, styles: meta.styles, date1904: meta.date1904 });
  const cells: Record<string, WorkbookCell> = {};
  for await (const row of stream.rows()) for (const [col, cell] of row.cells) cells[`${row.row}:${col}`] = cell;
  return cells;
}

describe("row mapping", () => {
  it("resolves shared strings (including index 0) and inline strings", async () => {
    const cells = await readAll(
      mini(
        '<sheetData><row r="1"><c r="A1" t="s"><v>0</v></c><c r="B1" t="s"><v>1</v></c><c r="C1" t="inlineStr"><is><t>inline</t></is></c><c r="D1" t="str"><v>cached</v></c></row></sheetData>',
        [{ name: "xl/sharedStrings.xml", data: `<sst ${NS}><si><t>first</t></si><si><t>Fish &amp; Chips</t></si></sst>` }],
      ),
    );
    expect(cells["0:0"]).toMatchObject({ value: "first", display: "first", type: "text" });
    expect(cells["0:1"]).toMatchObject({ value: "Fish & Chips" });
    expect(cells["0:2"]).toMatchObject({ value: "inline" });
    expect(cells["0:3"]).toMatchObject({ value: "cached" });
  });

  it("maps numbers, booleans, errors, dates and bold text", async () => {
    const cells = await readAll(
      mini(
        '<sheetData><row r="2" ht="30"><c r="A2"><v>12.5</v></c><c r="B2" t="b"><v>1</v></c><c r="C2" t="e"><v>#DIV/0!</v></c>' +
          '<c r="D2" s="1"><v>45000</v></c><c r="E2" s="2" t="inlineStr"><is><t>bold</t></is></c><c r="F2" s="2"/></row></sheetData>',
      ),
    );
    expect(cells["1:0"]).toMatchObject({ value: 12.5, type: "number" });
    expect(cells["1:1"]).toMatchObject({ value: true, type: "boolean", display: "TRUE" });
    // Error cells go through the shared converter, which (like the old importer) stringifies the error object.
    expect(cells["1:2"]).toBeDefined();
    expect(cells["1:3"]).toMatchObject({ value: 45000, type: "date", display: "2023-03-15" });
    expect(cells["1:4"]).toMatchObject({ value: "bold", style: { bold: true } });
    expect(cells["1:5"]).toBeUndefined(); // styled but empty
  });

  it("reads dates in the 1904 system", async () => {
    const sheet = '<sheetData><row r="1"><c r="A1" s="1"><v>1000</v></c></row></sheetData>';
    const a = await readAll(mini(sheet));
    const b = await readAll(mini(sheet, [], ' date1904="1"'));
    expect(a["0:0"].value).toBe(1000);
    expect(b["0:0"].display).not.toBe(a["0:0"].display);
  });

  it("keeps formulas with their cached results", async () => {
    const cells = await readAll(
      mini(
        '<sheetData><row r="1"><c r="A1"><v>2</v></c><c r="B1"><f>A1*3</f><v>6</v></c><c r="C1" t="str"><f>"x"&amp;"y"</f><v>xy</v></c><c r="D1"><f>NOW()</f></c></row></sheetData>',
      ),
    );
    expect(cells["0:1"]).toMatchObject({ formula: "=A1*3", value: 6 });
    expect(cells["0:2"]).toMatchObject({ formula: '="x"&"y"', value: "xy" });
    expect(cells["0:3"]).toMatchObject({ formula: "=NOW()" });
  });

  it("keeps the value of shared-formula dependents but not a bogus formula", async () => {
    const cells = await readAll(
      mini(
        '<sheetData><row r="1"><c r="A1"><v>1</v></c><c r="B1"><f t="shared" ref="B1:B3" si="0">A1*2</f><v>2</v></c></row>' +
          '<row r="2"><c r="A2"><v>2</v></c><c r="B2"><f t="shared" si="0"/><v>4</v></c></row>' +
          '<row r="3"><c r="A3"><v>3</v></c><c r="B3"><f t="shared" si="0"/></c></row></sheetData>',
      ),
    );
    expect(cells["0:1"]).toMatchObject({ formula: "=A1*2", value: 2 });
    expect(cells["1:1"]).toMatchObject({ value: 4 });
    expect(cells["1:1"].formula).toBeUndefined();
    expect(cells["2:1"]).toBeUndefined(); // no value and no usable formula
  });

  it("works when the shared strings part is missing", async () => {
    const cells = await readAll(mini('<sheetData><row r="1"><c r="A1" t="s"><v>3</v></c></row></sheetData>'));
    expect(cells["0:0"]).toMatchObject({ value: "" });
  });

  it("handles a self-closing sheetData, columns and merges after the data", async () => {
    const zip = await openZip(
      zipBlob(
        mini(
          '<cols><col min="2" max="3" width="20" customWidth="1"/></cols><sheetData/><mergeCells count="1"><mergeCell ref="A1:C2"/></mergeCells>',
        ),
      ),
    );
    const meta = await readWorkbookMeta(zip);
    const stream = readSheet(zip, meta.sheets[0], { strings: null, styles: meta.styles, date1904: false });
    for await (const row of stream.rows()) void row;
    expect(stream.scan().layout).toMatchObject({
      columnWidths: { 1: 150, 2: 150 },
      merges: [{ r: 0, c: 0, rs: 2, cs: 3 }],
      rowCount: 200,
      colCount: 26,
    });
  });

  it("finds </sheetData> and merges even when the text is read in tiny slices", async () => {
    const rows = Array.from({ length: 40 }, (_, i) => `<row r="${i + 1}"><c r="A${i + 1}"><v>${i}</v></c></row>`).join("");
    const entries = mini(`<cols><col min="1" max="1" width="10"/></cols><sheetData>${rows}</sheetData><mergeCells count="1"><mergeCell ref="B2:C3"/></mergeCells>`);
    const zip = await openZip(zipBlob(entries), { readSize: 3 });
    const meta = await readWorkbookMeta(zip);
    const stream = readSheet(zip, meta.sheets[0], { strings: null, styles: meta.styles, date1904: false });
    let count = 0;
    for await (const row of stream.rows()) count += row.cells.length;
    expect(count).toBe(40);
    expect(stream.scan().layout.merges).toEqual([{ r: 1, c: 1, rs: 2, cs: 2 }]);
    expect(stream.scan().layout.columnWidths).toEqual({ 0: 75 });
  });

  it("rejects rows that are out of order", async () => {
    const entries = mini('<sheetData><row r="5"><c r="A5"><v>1</v></c></row><row r="2"><c r="A2"><v>1</v></c></row></sheetData>');
    await expect(readAll(entries)).rejects.toMatchObject({ code: "CORRUPT" });
  });
});
