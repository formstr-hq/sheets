import { describe, expect, it } from "vitest";
import { readWorkbookMeta, resolveTarget } from "./workbookMeta";
import { openZip } from "./zip";
import { zipBlob, type TestEntry } from "./testZip";

const REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";

function workbookParts(options: {
  sheets: { name: string; rId: string; state?: string }[];
  rels: { id: string; type: string; target: string }[];
  date1904?: boolean;
}): TestEntry[] {
  const sheets = options.sheets
    .map((s) => `<sheet name="${s.name}" sheetId="${s.rId.replace(/\D/g, "")}"${s.state ? ` state="${s.state}"` : ""} r:id="${s.rId}"/>`)
    .join("");
  const rels = options.rels.map((r) => `<Relationship Id="${r.id}" Type="${REL}/${r.type}" Target="${r.target}"/>`).join("");
  return [
    {
      name: "xl/workbook.xml",
      data: `<?xml version="1.0"?><workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="${REL}"><workbookPr${options.date1904 ? ' date1904="1"' : ""}/><sheets>${sheets}</sheets></workbook>`,
    },
    { name: "xl/_rels/workbook.xml.rels", data: `<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${rels}</Relationships>` },
  ];
}

describe("readWorkbookMeta", () => {
  it("lists worksheets in workbook order with hidden state and odd part names", async () => {
    const zip = await openZip(
      zipBlob([
        ...workbookParts({
          sheets: [
            { name: "Second &amp; more", rId: "rId7" },
            { name: "Hidden", rId: "rId2", state: "hidden" },
            { name: "Chart", rId: "rId5" },
            { name: "Very", rId: "rId9", state: "veryHidden" },
          ],
          rels: [
            { id: "rId7", type: "worksheet", target: "worksheets/data-main.xml" },
            { id: "rId2", type: "worksheet", target: "/xl/worksheets/sheet9.xml" },
            { id: "rId5", type: "chartsheet", target: "chartsheets/sheet1.xml" },
            { id: "rId9", type: "worksheet", target: "../xl/worksheets/v.xml" },
          ],
        }),
        { name: "xl/worksheets/data-main.xml", data: "<worksheet/>" },
        { name: "xl/worksheets/sheet9.xml", data: "<worksheet/>" },
        { name: "xl/chartsheets/sheet1.xml", data: "<chartsheet/>" },
        { name: "xl/worksheets/v.xml", data: "<worksheet/>" },
      ]),
    );
    const meta = await readWorkbookMeta(zip);
    expect(meta.sheets).toEqual([
      { name: "Second & more", hidden: false, path: "xl/worksheets/data-main.xml", order: 0 },
      { name: "Hidden", hidden: true, path: "xl/worksheets/sheet9.xml", order: 1 },
      { name: "Very", hidden: true, path: "xl/worksheets/v.xml", order: 2 },
    ]);
    expect(meta.date1904).toBe(false);
  });

  it("reads the 1904 date system and works without styles.xml", async () => {
    const zip = await openZip(
      zipBlob([
        ...workbookParts({ sheets: [{ name: "S", rId: "rId1" }], rels: [{ id: "rId1", type: "worksheet", target: "worksheets/sheet1.xml" }], date1904: true }),
        { name: "xl/worksheets/sheet1.xml", data: "<worksheet/>" },
      ]),
    );
    const meta = await readWorkbookMeta(zip);
    expect(meta.date1904).toBe(true);
    expect(meta.styles.getStyleModel(0)).toBeNull();
  });

  it("parses styles.xml when present", async () => {
    const styles =
      '<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><numFmts count="1"><numFmt numFmtId="164" formatCode="0.00%"/></numFmts>' +
      '<fonts count="2"><font><sz val="11"/><name val="Calibri"/></font><font><b/><sz val="12"/><name val="Arial"/></font></fonts>' +
      '<fills count="2"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill></fills>' +
      '<borders count="1"><border/></borders><cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>' +
      '<cellXfs count="2"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/><xf numFmtId="164" fontId="1" fillId="0" borderId="0" xfId="0" applyFont="1"/></cellXfs></styleSheet>';
    const zip = await openZip(
      zipBlob([
        ...workbookParts({ sheets: [{ name: "S", rId: "rId1" }], rels: [{ id: "rId1", type: "worksheet", target: "worksheets/sheet1.xml" }] }),
        { name: "xl/worksheets/sheet1.xml", data: "<worksheet/>" },
        { name: "xl/styles.xml", data: styles },
      ]),
    );
    const { styles: manager } = await readWorkbookMeta(zip);
    expect(manager.getStyleModel(1)).toMatchObject({ numFmt: "0.00%", font: { bold: true, name: "Arial" } });
  });

  it("rejects a zip that is not a workbook", async () => {
    const zip = await openZip(zipBlob([{ name: "hello.txt", data: "hi" }]));
    await expect(readWorkbookMeta(zip)).rejects.toMatchObject({ code: "NOT_XLSX" });
  });

  it("returns no sheets when there are no worksheet relationships", async () => {
    const zip = await openZip(zipBlob(workbookParts({ sheets: [{ name: "S", rId: "rId1" }], rels: [] })));
    expect((await readWorkbookMeta(zip)).sheets).toEqual([]);
  });
});

describe("resolveTarget", () => {
  it("handles relative, absolute and parent paths", () => {
    expect(resolveTarget("worksheets/sheet1.xml")).toBe("xl/worksheets/sheet1.xml");
    expect(resolveTarget("/xl/worksheets/sheet1.xml")).toBe("xl/worksheets/sheet1.xml");
    expect(resolveTarget("../xl/worksheets/s.xml")).toBe("xl/worksheets/s.xml");
    expect(resolveTarget("./worksheets//s.xml")).toBe("xl/worksheets/s.xml");
  });
});
