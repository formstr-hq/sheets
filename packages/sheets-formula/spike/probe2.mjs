const decl = (await import("@sheetxl/formulas/declarations")).default;
for (const group of decl) {
  const d = group.declarations;
  const keys = Object.keys(d ?? {});
  console.log(group.category.padEnd(20), typeof d, Array.isArray(d) ? "array" : "obj", keys.length, JSON.stringify(keys.slice(0, 4)));
}
const fin = decl.find((g) => g.category === "Financial").declarations;
const first = Object.keys(fin)[0];
console.log("first financial:", first, JSON.stringify(fin[first]).slice(0, 400));
