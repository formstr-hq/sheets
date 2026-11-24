// Consumer of the built package, with NO bundler aliases: proves a bundled worker and the ExcelJS
// deep imports ship from a package and run in a browser.
import { zipSync, strToU8 } from "fflate";
import { probeXlsx } from "@formstr/xlsx-stream";

const ns = 'xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"';
const rel = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const bytes = zipSync({
  "[Content_Types].xml": strToU8('<?xml version="1.0"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="xml" ContentType="application/xml"/></Types>'),
  "xl/workbook.xml": strToU8(`<?xml version="1.0"?><workbook ${ns} xmlns:r="${rel}"><workbookPr date1904="1"/><sheets><sheet name="Data" sheetId="1" r:id="rId1"/><sheet name="Hidden &amp; more" sheetId="2" state="hidden" r:id="rId2"/></sheets></workbook>`),
  "xl/_rels/workbook.xml.rels": strToU8(`<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="${rel}/worksheet" Target="worksheets/sheet1.xml"/><Relationship Id="rId2" Type="${rel}/worksheet" Target="worksheets/sheet2.xml"/></Relationships>`),
  "xl/worksheets/sheet1.xml": strToU8(`<?xml version="1.0"?><worksheet ${ns}><sheetData><row r="1"><c r="A1" t="s"><v>0</v></c></row></sheetData></worksheet>`),
  "xl/worksheets/sheet2.xml": strToU8(`<?xml version="1.0"?><worksheet ${ns}><sheetData/></worksheet>`),
  "xl/sharedStrings.xml": strToU8(`<?xml version="1.0"?><sst ${ns} count="2" uniqueCount="2"><si><t>héllo &amp; wörld</t></si><si><r><t>rich</t></r><r><t> text</t></r></si></sst>`),
});

const out = document.getElementById("out")!;
probeXlsx(new Blob([bytes as BlobPart]))
  .then((probe) => {
    (window as unknown as { RESULT: unknown }).RESULT = probe;
    out.textContent = JSON.stringify(probe, null, 2);
  })
  .catch((error) => {
    (window as unknown as { RESULT: unknown }).RESULT = { error: String(error) };
    out.textContent = `ERROR ${error}`;
  });
