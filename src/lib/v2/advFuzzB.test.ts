// Differential fuzzing, family B: a pure in-memory MODEL as the oracle. The model keeps formula text per cell and applies every
// insert/delete with rewriteForEdit (the same engine the positional code uses, but with no store, no chunks and no ids), and every
// typed value or formula. The id-based store (with chunks that are NOT in memory, reopen from the saved root, several saves)
// must end up with exactly the same formulas and plain values.
import { describe, expect, it } from "vitest";
import { DEFAULT_MIX, type GenState, apply, dump, generate, openSide, rng, saveSide, setup, world, describeOp, type Mix, type Op } from "./advFuzzHarness";
import { moveIndex } from "./layoutEdit";

interface MCell {
  formula?: string;
  value?: unknown;
}
interface MSheet {
  name: string;
  rows: number;
  cols: number;
  cells: Map<string, MCell>;
}

const STRUCT_ONLY: Mix = { ...DEFAULT_MIX, sort: 0, drag: 0, fill: 0, paste: 0, move: 0, replace: 0, layout: 0, tabs: 0, undo: 0, recalc: 0, struct: 60, type: 25, save: 10, nasty: 0.2 };

async function modelFrom(store: Awaited<ReturnType<typeof setup>>["ref"]["side"]["store"]): Promise<MSheet[]> {
  const d = await dump(store);
  return d.tabs.map((t) => {
    const cells = new Map<string, MCell>();
    for (const [k, v] of Object.entries(t.cells)) {
      const [value, formula] = JSON.parse(v) as [unknown, string | null];
      cells.set(k, formula ? { formula } : { value });
    }
    return { name: t.name, rows: t.rows, cols: t.cols, cells };
  });
}

async function modelStruct(sheets: MSheet[], tab: number, edit: { kind: "insert" | "delete"; axis: "row" | "col"; index: number; count: number }): Promise<void> {
  const engine = await import("@formstr/sheets-formula");
  const target = sheets[tab];
  const size = edit.axis === "row" ? target.rows : target.cols;
  let { index, count } = edit;
  if (edit.kind === "delete") {
    index = Math.min(index, size - 1);
    count = Math.min(count, size - index);
    if (count >= size) return;
  } else index = Math.min(index, size);
  if (count <= 0) return;
  const e = { kind: edit.kind, axis: edit.axis, index, count };
  for (const sheet of sheets) {
    for (const cell of sheet.cells.values()) {
      if (!cell.formula) continue;
      try {
        cell.formula = engine.print(engine.rewriteForEdit(engine.parse(cell.formula), { ...e, sheet: target.name }, sheet.name)).replace(/^=/, "");
      } catch {
        // unreadable text stays
      }
    }
  }
  const moved = new Map<string, MCell>();
  for (const [key, cell] of target.cells) {
    const at = key.indexOf(":");
    const r = Number(key.slice(0, at));
    const c = Number(key.slice(at + 1));
    const m = moveIndex(edit.axis === "row" ? r : c, e);
    if (m === null) continue;
    moved.set(edit.axis === "row" ? `${m}:${c}` : `${r}:${m}`, cell);
  }
  target.cells = moved;
  if (edit.axis === "row") target.rows += edit.kind === "insert" ? count : -count;
  else target.cols += edit.kind === "insert" ? count : -count;
}

function modelType(sheets: MSheet[], tab: number, row: number, col: number, text: string): void {
  const s = sheets[tab];
  row = Math.max(0, Math.min(row, s.rows - 1));
  col = Math.max(0, Math.min(col, s.cols - 1));
  const key = `${row}:${col}`;
  const trimmed = text.trim();
  if (text.startsWith("=")) s.cells.set(key, { formula: text.slice(1) });
  else if (trimmed === "") s.cells.delete(key);
  else s.cells.set(key, { value: /^[-+]?[\d,]*\.?\d+(e[-+]?\d+)?$/i.test(trimmed) ? Number(trimmed.replace(/,/g, "")) : text });
}

describe("adversarial differential fuzz B (pure model oracle)", () => {
  for (const seed of [201, 202, 203, 204, 205, 206]) {
    it(`id store equals the in-memory model (seed ${seed})`, async () => {
      const w = await world(2, 1500);
      // only the id store: the model replaces the positional reference
      const { ref, sub } = await setup(w);
      let side = { side: sub.side };
      const model = await modelFrom(ref.side.store);
      const rand = rng(seed);
      const state: GenState = { strayChunks: 0, names: 1 };
      const log: string[] = [];
      for (let i = 0; i < 30; i++) {
        const op: Op = generate(rand, side.side.store, side.side, state, STRUCT_ONLY);
        log.push(describeOp(op));
        if (op.k === "struct") {
          const tab = ((op.tab % model.length) + model.length) % model.length;
          try {
            await apply(w, side, op);
          } catch (error) {
            if (error instanceof Error && /too many parts/.test(error.message)) continue;
            if (error instanceof Error && error.constructor.name === "Skip") continue;
            throw error;
          }
          await modelStruct(model, tab, op);
        } else if (op.k === "type") {
          const tab = ((op.tab % model.length) + model.length) % model.length;
          await apply(w, side, op);
          modelType(model, tab, op.row, op.col, op.text);
        } else if (op.k === "save") {
          await apply(w, side, op);
        }
      }
      side.side = await saveSide(w, side.side, true);
      const got = await dump(side.side.store);
      const bad: string[] = [];
      got.tabs.forEach((t, ti) => {
        const m = model[ti];
        if (t.rows !== m.rows || t.cols !== m.cols) bad.push(`tab ${t.name} size ${t.rows}x${t.cols} model ${m.rows}x${m.cols}`);
        const keys = new Set([...Object.keys(t.cells), ...m.cells.keys()]);
        for (const k of keys) {
          const real = t.cells[k] ? (JSON.parse(t.cells[k]) as [unknown, string | null]) : null;
          const exp = m.cells.get(k);
          const realFormula = real?.[1] ?? undefined;
          if (realFormula !== exp?.formula) bad.push(`${t.name} ${k}: formula store=${realFormula} model=${exp?.formula}`);
          else if (!exp?.formula && JSON.stringify(real?.[0] ?? undefined) !== JSON.stringify(exp?.value ?? undefined)) bad.push(`${t.name} ${k}: value store=${JSON.stringify(real?.[0])} model=${JSON.stringify(exp?.value)}`);
        }
      });
      if (bad.length) console.log(`SEED ${seed} ops:\n${log.join("\n")}\nDIFFS (${bad.length}):\n${bad.slice(0, 12).join("\n")}`);
      expect(bad.slice(0, 12)).toEqual([]);
    }, 400_000);
  }
});

void openSide;
