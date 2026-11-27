// Independent cell/row count of an xlsx using ExcelJS's own Node streaming reader (not our code).
//   node tools/count-xlsx.mjs <file.xlsx>
// Prints JSON: { sheets: [{ name, rows, cells }], cells }. Used to check the streaming importer.
import { createRequire } from "node:module";
import { createReadStream } from "node:fs";

const require = createRequire(new URL("../package.json", import.meta.url));
const ExcelJS = require("exceljs");

const file = process.argv[2];
if (!file) {
  console.error("usage: node tools/count-xlsx.mjs <file.xlsx>");
  process.exit(2);
}
const reader = new ExcelJS.stream.xlsx.WorkbookReader(createReadStream(file), {
  sharedStrings: "cache",
  hyperlinks: "ignore",
  worksheets: "emit",
  styles: "ignore",
});
const sheets = [];
let total = 0;
for await (const sheet of reader) {
  let rows = 0;
  let cells = 0;
  for await (const row of sheet) {
    if (!row.hasValues) continue;
    rows++;
    row.eachCell({ includeEmpty: false }, () => cells++);
  }
  sheets.push({ name: sheet.name, rows, cells });
  total += cells;
}
console.log(JSON.stringify({ sheets, cells: total }));
