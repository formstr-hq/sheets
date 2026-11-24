import { describe, expect, it } from "vitest";
import { addBorder, edgeKey, edgesIn, type Border } from "./borders";
import { MergeIndex } from "./merges";
import { single } from "./selection";

describe("merged cells", () => {
  const index = new MergeIndex();
  index.set([
    { r: 2, c: 1, rs: 3, cs: 2 },
    { r: 10, c: 0, rs: 1, cs: 4 },
    { r: 5, c: 5, rs: 1, cs: 1 }, // not a merge
  ]);

  it("finds the block a cell belongs to", () => {
    expect(index.at(3, 2)).toMatchObject({ r: 2, c: 1 });
    expect(index.at(2, 1)).toMatchObject({ r: 2, c: 1 });
    expect(index.at(5, 5)).toBeUndefined();
    expect(index.at(1, 1)).toBeUndefined();
  });

  it("finds blocks that start above the window", () => {
    expect(index.intersecting(4, 0, 6, 9)).toHaveLength(1);
    expect(index.intersecting(6, 0, 9, 9)).toHaveLength(0);
  });

  it("grows a selection to cover whole blocks", () => {
    expect(index.expand({ r1: 4, c1: 2, r2: 4, c2: 2 })).toEqual({ r1: 2, c1: 1, r2: 4, c2: 2 });
    expect(index.expand({ r1: 0, c1: 0, r2: 2, c2: 1 })).toEqual({ r1: 0, c1: 0, r2: 4, c2: 2 });
  });

  it("moves the active cell to the top-left of its block", () => {
    const next = index.expandSelection(single({ row: 4, col: 2 }));
    expect(next.active).toEqual({ row: 2, col: 1 });
    expect(next.ranges[0]).toEqual({ r1: 2, c1: 1, r2: 4, c2: 2 });
  });
});

describe("borders", () => {
  const all: Border = { r1: 1, c1: 1, r2: 2, c2: 2, type: "all", color: "#f00" };

  it("draws every line of a range for 'all'", () => {
    const edges = edgesIn([all], 0, 0, 5, 5);
    // 3 horizontal lines x 2 columns, 3 vertical lines x 2 rows.
    expect(edges.h.size).toBe(6);
    expect(edges.v.size).toBe(6);
    expect(edges.h.get(edgeKey(3, 2))?.color).toBe("#f00");
  });

  it("draws only the outline for 'outside' and only the middle for 'inside'", () => {
    const outside = edgesIn([{ ...all, type: "outside" }], 0, 0, 5, 5);
    expect(outside.h.size).toBe(4);
    expect(outside.v.size).toBe(4);
    const inside = edgesIn([{ ...all, type: "inside" }], 0, 0, 5, 5);
    expect(inside.h.size).toBe(2);
    expect(inside.v.size).toBe(2);
  });

  it("draws one side", () => {
    const edges = edgesIn([{ ...all, type: "bottom" }], 0, 0, 5, 5);
    expect([...edges.h.keys()]).toEqual([edgeKey(3, 1), edgeKey(3, 2)]);
    expect(edges.v.size).toBe(0);
  });

  it("'none' removes lines drawn before it, and a later border draws again", () => {
    const erased = edgesIn([all, { r1: 1, c1: 1, r2: 1, c2: 1, type: "none" }], 0, 0, 5, 5);
    expect(erased.h.has(edgeKey(1, 1))).toBe(false);
    expect(erased.v.has(edgeKey(1, 1))).toBe(false);
    expect(erased.h.has(edgeKey(3, 2))).toBe(true);
    const again = edgesIn([all, { ...all, type: "none" }, { ...all, type: "top", color: "#00f" }], 0, 0, 5, 5);
    expect(again.h.size).toBe(2);
    expect(again.h.get(edgeKey(1, 1))?.color).toBe("#00f");
  });

  it("only looks at the window, so a whole-column border is cheap", () => {
    const edges = edgesIn([{ r1: 0, c1: 0, r2: 24_999_999, c2: 0, type: "outside" }], 1_000_000, 0, 1_000_020, 3);
    expect(edges.v.size).toBe(2 * 21);
  });

  it("drops requests a 'none' fully covers", () => {
    const list = addBorder([all], { r1: 0, c1: 0, r2: 9, c2: 9, type: "none" });
    expect(list).toEqual([]);
    const partial = addBorder([all], { r1: 1, c1: 1, r2: 1, c2: 1, type: "none" });
    expect(partial).toHaveLength(2);
  });
});
