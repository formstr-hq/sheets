/// <reference types="node" />
import { createRequire } from "node:module";
import { describe, expect, it } from "vitest";
import {
  createSheetReader,
  createStyleManager,
  parseRelationships,
  parseSax,
  parseWorkbookXml,
} from "./excelInternals";

// excelInternals.ts deep-imports private ExcelJS files. This fails loudly if an exceljs upgrade
// moves or reshapes them, instead of breaking the import at runtime.
describe("exceljs internals", () => {
  it("is the pinned version", () => {
    const require = createRequire(import.meta.url);
    expect(require("exceljs/package.json").version).toBe("4.4.0");
  });

  it("exposes the facade functions", () => {
    for (const fn of [createSheetReader, createStyleManager, parseRelationships, parseSax, parseWorkbookXml]) {
      expect(typeof fn).toBe("function");
    }
    const styles = createStyleManager();
    expect(typeof styles.parseStream).toBe("function");
    expect(typeof styles.getStyleModel).toBe("function");
  });

  it("parses xml from a string iterator", async () => {
    async function* source() {
      yield '<Relationships><Relationship Id="rId1" Type="t" Target="worksheets/sheet1.xml"/></Relationships>';
    }
    const rels = await parseRelationships(source());
    expect(rels).toEqual([expect.objectContaining({ Id: "rId1", Target: "worksheets/sheet1.xml" })]);
  });

  it("reads rows from a worksheet string stream", async () => {
    async function* source() {
      yield '<worksheet><sheetData><row r="1"><c r="A1"><v>7</v></c><c r="B1" t="s"><v>3</v></c></row></sheetData>';
      yield "</worksheet>";
    }
    const reader = createSheetReader({
      id: "1",
      source: source(),
      context: { styles: createStyleManager(), properties: { model: {} } },
    });
    const rows = [];
    for await (const row of reader) rows.push(row);
    expect(rows).toHaveLength(1);
    expect(rows[0].getCell(1).value).toBe(7);
    expect(rows[0].getCell(2).value).toEqual({ sharedString: 3 });
  });
});
