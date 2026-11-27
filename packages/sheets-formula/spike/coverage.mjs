// Spike B: which of FortuneSheet's function names do @sheetxl/formulas and @formulajs/formulajs cover?
// Sources: FortuneSheet's evaluable set = Object.keys(formulajs 2.9.3) (what formula-parser exposes);
// FortuneSheet's help popup = locale.functionlist in @fortune-sheet/core; sheetxl = its `declarations` export.
import { createRequire } from "node:module";
import { readFileSync, readdirSync } from "node:fs";
import { fileURLToPath } from "node:url";
import path from "node:path";

const root = path.resolve(path.dirname(fileURLToPath(import.meta.url)), "../../..");
const pnpm = path.join(root, "node_modules/.pnpm");
const find = (prefix) => readdirSync(pnpm).find((name) => name.startsWith(prefix));

// 1. FortuneSheet's evaluable names (formulajs 2.9.3, the version formula-parser installs).
const oldDir = find("@formulajs+formulajs@2.9.3");
const requireOld = createRequire(path.join(pnpm, oldDir, "node_modules/@formulajs/formulajs/package.json"));
const old = requireOld("@formulajs/formulajs");
const evaluable = Object.keys(old).filter((key) => typeof old[key] === "function" && /^[A-Z][A-Z0-9_.]*$/.test(key));

// 2. FortuneSheet's popup list.
const coreDir = find("@fortune-sheet+core@");
const core = readFileSync(path.join(pnpm, coreDir, "node_modules/@fortune-sheet/core/dist/index.esm.js"), "utf8");
const at = core.indexOf("functionlist");
const popup = new Set();
for (const match of core.slice(at, at + 400_000).matchAll(/\bn:\s*"([A-Z][A-Z0-9_.]*)"/g)) popup.add(match[1]);

// 3. sheetxl.
const decl = (await import("@sheetxl/formulas/declarations")).default;
const sheetxl = new Set();
for (const group of decl) for (const name of Object.keys(group.declarations ?? {})) sheetxl.add(name.toUpperCase());

// 4. formulajs current.
const fjs = await import("@formulajs/formulajs");
const fjsAll = fjs.default ?? fjs;
const formulajs = new Set(Object.keys(fjsAll).filter((key) => typeof fjsAll[key] === "function").map((k) => k.toUpperCase()));

const norm = (name) => name.replace(/_/g, ".");
const has = (set, name) => set.has(name) || set.has(norm(name)) || set.has(name.replace(/\./g, "_"));
const missing = (names, ...sets) => names.filter((name) => !sets.some((set) => has(set, name)));

const popupList = [...popup];
const report = {
  counts: { fortuneEvaluable: evaluable.length, fortunePopup: popupList.length, sheetxl: sheetxl.size, formulajsCurrent: formulajs.size },
  evaluableMissingFromSheetxl: missing(evaluable, sheetxl),
  evaluableMissingFromBoth: missing(evaluable, sheetxl, formulajs),
  popupMissingFromSheetxl: missing(popupList, sheetxl),
  popupMissingFromBoth: missing(popupList, sheetxl, formulajs),
  templates: ["SUM", "IF", "COUNTIF", "COUNTA", "DATE"].map((name) => [name, sheetxl.has(name)]),
  sample: [...sheetxl].slice(0, 12),
};
console.log(JSON.stringify(report, null, 1));
