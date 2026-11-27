const mod = await import("@sheetxl/formulas");
const F = mod.default ?? mod;
console.log("keys of default export:", Object.keys(F).length, Object.keys(F).slice(0, 15));
const decl = await import("@sheetxl/formulas/declarations");
const D = decl.default ?? decl;
console.log("declarations keys:", Object.keys(D).length, Object.keys(D).slice(0, 12));
console.log("typeof SUM", typeof F.SUM, "typeof Formulas", typeof F.Formulas);
