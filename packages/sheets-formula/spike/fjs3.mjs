import * as f from "@formulajs/formulajs";
const t = (label, fn) => { try { console.log(label.padEnd(26), fn()); } catch (e) { console.log(label.padEnd(26), "THROWS", e.message); } };
for (const n of ["NORMSDIST","LOGNORMDIST","TDIST","NORMDIST","STDEV","RANK","MODE","VAR","PERCENTILE","CHIDIST","FDIST","BINOMDIST","POISSON","WEIBULL","NORMSINV","TINV"]) console.log(n, typeof f[n]);
t("NORMSDIST(0)", () => f.NORMSDIST(0));
t("NORM.S.DIST(0,true)", () => f.NORM.S.DIST(0, true));
t("LOGNORMDIST(1,0,1)", () => f.LOGNORMDIST && f.LOGNORMDIST(1, 0, 1));
t("LOGNORM.DIST(1,0,1,true)", () => f.LOGNORM.DIST(1, 0, 1, true));
t("TDIST(2,10,2)", () => f.TDIST && f.TDIST(2, 10, 2));
t("T.DIST.2T(2,10)", () => f.T.DIST["2T"] ? f.T.DIST["2T"](2, 10) : Object.keys(f.T.DIST));
