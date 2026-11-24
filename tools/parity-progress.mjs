// Prints how close the big-sheet grid is to FortuneSheet parity, from parity/matrix.json.
// Usage: pnpm parity            (add --list to see every open item)
import { readFileSync } from "node:fs";

const matrix = JSON.parse(readFileSync(new URL("../parity/matrix.json", import.meta.url), "utf8"));
const score = { done: 1, partial: 0.5, todo: 0 };
const bar = (fraction, width = 30) => {
  const filled = Math.round(fraction * width);
  return `[${"█".repeat(filled)}${"░".repeat(width - filled)}]`;
};

let total = 0;
let points = 0;
const open = [];
for (const [group, items] of Object.entries(matrix.groups)) {
  const entries = Object.entries(items);
  const got = entries.reduce((sum, [, status]) => sum + score[status], 0);
  total += entries.length;
  points += got;
  console.log(`${group.padEnd(26)} ${bar(got / entries.length, 20)} ${Math.round((got / entries.length) * 100)}%`);
  for (const [name, status] of entries) if (status !== "done") open.push(`${status === "partial" ? "~" : " "} ${group}: ${name}`);
}
const fraction = points / total;
console.log("");
console.log(`PARITY  ${bar(fraction)} ${Math.round(fraction * 100)}%  (${open.length} of ${total} items still open)`);
if (open.length === 0) console.log("\n🎉 Full parity: every item is done. Time to run the automated check and remove FortuneSheet.");
if (process.argv.includes("--list")) console.log("\n" + open.join("\n"));
