// Generates a big demo spreadsheet (no row cap) with ExcelJS's streaming writer, so memory stays flat.
//   node tools/make-big-xlsx.mjs <targetMB> [cols] [tabs] [outFile] [flags]
//   node tools/make-big-xlsx.mjs 100 30 4 big-demo-100mb.xlsx
//   node tools/make-big-xlsx.mjs 40 20 4 text-heavy.xlsx --shared-strings --styles
//
// Flags:
//   --shared-strings  write xl/sharedStrings.xml (ExcelJS writes it AFTER the sheets, like Excel
//                     does) and draw text from a pool of 20,000 values, so the table dedups.
//   --unique-strings  with --shared-strings: every text cell is a distinct string (worst case for
//                     the shared-strings table). Without --shared-strings text is inline.
//   --styles          bold header, number formats, a date column, column widths, and the last
//                     two header cells merged.
// Rows are added round-robin across tabs. Prints the exact cell count when done.
import { createRequire } from "node:module";
import { statSync } from "node:fs";

const require = createRequire(new URL("../package.json", import.meta.url));
const ExcelJS = require("exceljs");

const flags = new Set(process.argv.slice(2).filter((a) => a.startsWith("--")));
const args = process.argv.slice(2).filter((a) => !a.startsWith("--"));
const targetMB = Number(args[0] ?? 100);
const cols = Number(args[1] ?? 30);
const tabs = Number(args[2] ?? 4);
const out = args[3] ?? `big-demo-${targetMB}mb.xlsx`;
const sharedStrings = flags.has("--shared-strings");
const uniqueStrings = flags.has("--unique-strings");
const styles = flags.has("--styles");

let seed = 42;
const rnd = () => (seed = (seed * 1664525 + 1013904223) >>> 0) / 2 ** 32;
const words = ["alpha", "bravo", "delta", "omega", "sigma", "kappa", "zeta", "theta"];
const POOL = 20000;
let unique = 0;

const wb = new ExcelJS.stream.xlsx.WorkbookWriter({ filename: out, useStyles: styles, useSharedStrings: sharedStrings });
const sheets = Array.from({ length: tabs }, (_, t) => wb.addWorksheet(`Tab ${t + 1}`));
if (styles) {
  for (const ws of sheets) {
    ws.columns = Array.from({ length: cols }, (_, c) => ({ width: 10 + (c % 5) * 4 }));
  }
}
sheets.forEach((ws) => {
  const header = ws.addRow(Array.from({ length: cols }, (_, c) => `Column ${c + 1}`));
  if (styles) {
    header.font = { bold: true };
    // A stream sheet only accepts merges while the rows involved are uncommitted.
    if (cols >= 3) ws.mergeCells(1, cols - 1, 1, cols);
  }
  header.commit();
});

const bytesPerRowGuess = cols * (sharedStrings ? 8 : 9.4); // compressed bytes per row, rough
const rowsPerTab = Math.ceil((targetMB * 1024 * 1024) / bytesPerRowGuess / tabs);
const text = () =>
  sharedStrings && !uniqueStrings
    ? `${words[Math.floor(rnd() * 8)]}-${Math.floor(rnd() * POOL).toString(36)}`
    : uniqueStrings
      ? `${words[Math.floor(rnd() * 8)]}-${(unique++).toString(36)}-${Math.floor(rnd() * 1e9).toString(36)}`
      : `${words[Math.floor(rnd() * 8)]}-${Math.floor(rnd() * 1e9).toString(36)}`;
const DAY = 86400000;
const row = (r) =>
  Array.from({ length: cols }, (_, c) => {
    if (styles && c % 7 === 6) return new Date(Date.UTC(2020, 0, 1) + ((r * 31 + c) % 2000) * DAY);
    return c % 3 === 0 ? Math.round(rnd() * 1e6) / 100 : c % 3 === 1 ? text() : r * (c + 1);
  });

for (let r = 1; r <= rowsPerTab; r++) {
  for (const ws of sheets) {
    const added = ws.addRow(row(r));
    if (styles) {
      for (let c = 0; c < cols; c += 3) added.getCell(c + 1).numFmt = "#,##0.00";
      for (let c = 6; c < cols; c += 7) added.getCell(c + 1).numFmt = "yyyy-mm-dd";
    }
    added.commit();
  }
  if (r % 50000 === 0) console.log(`  ${r} rows per tab`);
}
sheets.forEach((ws) => ws.commit());
await wb.commit();
const mb = statSync(out).size / 1024 / 1024;
const cells = tabs * (rowsPerTab + 1) * cols;
console.log(`wrote ${out}: ${tabs} tabs x ${rowsPerTab + 1} rows (incl. header) x ${cols} cols = ${cells} cells (${(cells / 1e6).toFixed(1)}M), ${mb.toFixed(1)} MB`);
