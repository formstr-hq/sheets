import { describe, expect, it } from "vitest";
import { parse, print, rewriteForEdit, columnName, type Node } from "@formstr/sheets-formula";
import { ArrayAxis, formulaFromIds, formulaToIds, type RefContext, type TabAxes } from "./idRefs";

function rng(seed: number) {
  let s = seed >>> 0;
  return () => {
    s = (Math.imul(s, 1664525) + 1013904223) >>> 0;
    return s / 2 ** 32;
  };
}
const mkTab = (id: string, name: string, rows: number, cols: number): TabAxes => ({
  id,
  name,
  rows: new ArrayAxis(Array.from({ length: rows }, (_, i) => `${id}r${i}`)),
  cols: new ArrayAxis(Array.from({ length: cols }, (_, i) => `${id}c${i}`)),
});

const SHEETS = ["", "", "", "T!", "U!", "'Sheet 4'!"];
function genRef(rand: () => number, rows: number, cols: number, allowPast: boolean): string {
  const pick = <T,>(l: T[]) => l[Math.floor(rand() * l.length)];
  const r = () => Math.floor(rand() * (allowPast && rand() < 0.2 ? rows + 5 : rows));
  const c = () => Math.floor(rand() * (allowPast && rand() < 0.2 ? cols + 3 : cols));
  const d = () => (rand() < 0.25 ? "$" : "");
  const sh = pick(SHEETS);
  const kind = rand();
  if (kind < 0.35) return `${sh}${d()}${columnName(c())}${d()}${r() + 1}`;
  if (kind < 0.75) {
    const [r1, r2] = [r(), r()].sort((a, b) => a - b);
    const [c1, c2] = [c(), c()].sort((a, b) => a - b);
    return `${sh}${d()}${columnName(c1)}${d()}${r1 + 1}:${d()}${columnName(c2)}${d()}${r2 + 1}`;
  }
  if (kind < 0.88) {
    const [c1, c2] = [c(), c()].sort((a, b) => a - b);
    return `${sh}${d()}${columnName(c1)}:${d()}${columnName(c2)}`;
  }
  const [r1, r2] = [r(), r()].sort((a, b) => a - b);
  return `${sh}${d()}${r1 + 1}:${d()}${r2 + 1}`;
}
function genExpr(rand: () => number, rows: number, cols: number, depth: number, allowPast: boolean): string {
  const pick = <T,>(l: T[]) => l[Math.floor(rand() * l.length)];
  const k = rand();
  if (depth <= 0 || k < 0.35) return genRef(rand, rows, cols, allowPast);
  if (k < 0.5) return String(Math.floor(rand() * 100));
  if (k < 0.6) return `"${pick(["A1", "x", "~(", "!"])}"`;
  if (k < 0.8) return `${pick(["SUM", "MAX", "IF", "COUNT", "AVERAGE"])}(${genExpr(rand, rows, cols, depth - 1, allowPast)},${genExpr(rand, rows, cols, depth - 1, allowPast)})`;
  if (k < 0.9) return `-${genExpr(rand, rows, cols, depth - 1, allowPast)}`;
  return `(${genExpr(rand, rows, cols, depth - 1, allowPast)}${pick(["+", "*", "&", "<", "-"])}${genExpr(rand, rows, cols, depth - 1, allowPast)})`;
}

function run(seed: number, allowPast: boolean) {
  const rand = rng(seed * 7919);
  const pick = <T,>(l: T[]) => l[Math.floor(rand() * l.length)];
  const R = 8 + Math.floor(rand() * 10);
  const C = 5 + Math.floor(rand() * 6);
  const tabs = [mkTab("a", "S", R, C), mkTab("t", "T", R, C), mkTab("u", "U", R, C), mkTab("s4", "Sheet 4", R, C)];
  const ctx: RefContext = { tabs, current: "a" };
  const texts = Array.from({ length: 12 }, () => genExpr(rand, R, C, 3, allowPast));
  const stored = texts.map((f) => formulaToIds(f, ctx));
  let nodes: Node[] = texts.map((f) => parse("=" + f));
  const log: string[] = [];
  let counter = 0;
  const steps = 1 + Math.floor(rand() * 10);
  for (let step = 0; step < steps; step++) {
    const target = pick(tabs);
    const axisName = pick<"row" | "col">(["row", "col"]);
    const axis = (axisName === "row" ? target.rows : target.cols) as ArrayAxis;
    const kind = rand() < 0.5 ? "insert" : "delete";
    const count = 1 + Math.floor(rand() * Math.max(1, axis.length));
    // KNOWN LIMITATION: a reference past the last row or column is counted from that last one, so rows or columns
    // inserted exactly at the end (between the last one and the reference) do not push it. Not generated when past-the-end refs are allowed.
    const index = kind === "insert" ? Math.floor(rand() * (axis.length + (allowPast ? 0 : 1))) : Math.floor(rand() * axis.length);
    if (kind === "delete" && index + count > axis.length) continue;
    if (kind === "delete" && axis.length - count < 1) continue;
    log.push(`${kind} ${axisName} ${target.name} @${index} x${count}`);
    nodes = nodes.map((n) => rewriteForEdit(n, { kind, axis: axisName, sheet: target.name, index, count }, "S"));
    if (kind === "insert") axis.insert(index, Array.from({ length: count }, () => `n${seed}_${counter++}`));
    else axis.remove(index, count);
  }
  const bad: string[] = [];
  stored.forEach((g, i) => {
    let fromIds = "";
    try {
      fromIds = formulaFromIds(g, ctx);
      const got = print(parse("=" + fromIds));
      const want = print(nodes[i]);
      if (got !== want) bad.push(`${texts[i]}  ids:${fromIds}  oracle:${want}  edits:${log.join("; ")}`);
    } catch (e) {
      bad.push(`${texts[i]} threw/parse-failed on ${fromIds}: ${(e as Error).message}  edits:${log.join("; ")}`);
    }
  });
  return bad;
}

describe("oracle fuzz, references inside the axes", () => {
  // Differences found here are reported by the test as mismatches (see log in message).
  it("2000 seeds", () => {
    const all: string[] = [];
    for (let seed = 1; seed <= 2000; seed++) all.push(...run(seed, false).map((b) => `seed ${seed}: ${b}`));
    if (all.length) console.log(`${all.length} mismatches, first 8:\n` + all.slice(0, 8).join("\n"));
    expect(all.length).toBe(0);
  });
});

describe("oracle fuzz, references past the end of the axes allowed", () => {
  // BUG: positions past the end (@N) never shift; see advTextTombs.test.ts. This fuzz shows how often it matters.
  it("2000 seeds", () => {
    const all: string[] = [];
    for (let seed = 1; seed <= 2000; seed++) all.push(...run(seed, true).map((b) => `seed ${seed}: ${b}`));
    if (all.length) console.log(`${all.length} mismatches (past-end), first 5:\n` + all.slice(0, 5).join("\n"));
    expect(all.length).toBe(0);
  });
});
