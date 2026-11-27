const mod = await import("@sheetxl/formulas");
const names = Object.keys(mod);
console.log("named exports:", names.length, names.slice(0, 20));
const F = mod.default;
console.log("default:", typeof F, F && Object.getOwnPropertyNames(F).slice(0, 10));
for (const key of names.slice(0, 30)) console.log(key, typeof mod[key]);
