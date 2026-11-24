import { describe, expect, it } from "vitest";
import { BOND_FUNCTIONS as F } from "./bonds";
import { FormulaError } from "./types";

/** Serial number of a date, 1900 system (valid for dates after 1900-02-28). */
const D = (y: number, m: number, d: number) => Math.round((Date.UTC(y, m - 1, d) - Date.UTC(1899, 11, 30)) / 86400000);
const near = (actual: unknown, expected: number, rel = 1e-6) => {
  expect(typeof actual).toBe("number");
  expect(Math.abs((actual as number) - expected)).toBeLessThanOrEqual(Math.abs(expected) * rel + 1e-12);
};
const err = (v: unknown, code: string) => {
  expect(v).toBeInstanceOf(FormulaError);
  expect((v as FormulaError).code).toBe(code);
};

describe("date helper", () => {
  it("matches Excel serials", () => {
    expect(D(2024, 3, 15)).toBe(45366);
  });
});

describe("Microsoft documentation examples", () => {
  it("COUPNUM = 4", () => expect(F.COUPNUM([D(2007, 1, 25), D(2008, 11, 15), 2, 1])).toBe(4));
  it("COUPPCD = 15-Nov-10", () => expect(F.COUPPCD([D(2011, 1, 25), D(2011, 11, 15), 2, 1])).toBe(D(2010, 11, 15)));
  it("COUPNCD = 15-May-11", () => expect(F.COUPNCD([D(2011, 1, 25), D(2011, 11, 15), 2, 1])).toBe(D(2011, 5, 15)));
  it("COUPDAYBS = 71", () => expect(F.COUPDAYBS([D(2011, 1, 25), D(2011, 11, 15), 2, 1])).toBe(71));
  it("COUPDAYS = 181", () => expect(F.COUPDAYS([D(2011, 1, 25), D(2011, 11, 15), 2, 1])).toBe(181));
  it("COUPDAYSNC = 110", () => expect(F.COUPDAYSNC([D(2011, 1, 25), D(2011, 11, 15), 2, 1])).toBe(110));
  it("PRICE = 94.63436162", () =>
    near(F.PRICE([D(2008, 2, 15), D(2017, 11, 15), 0.0575, 0.065, 100, 2, 0]), 94.63436162));
  it("YIELD = 0.065", () =>
    near(F.YIELD([D(2008, 2, 15), D(2016, 11, 15), 0.0575, 95.04287, 100, 2, 0]), 0.065, 1e-6));
  it("DURATION (2018..2048) = 10.9191453", () =>
    near(F.DURATION([D(2018, 7, 1), D(2048, 1, 1), 0.08, 0.09, 2, 1]), 10.9191453));
  it("DURATION (2008..2016) = 5.993775", () =>
    near(F.DURATION([D(2008, 1, 1), D(2016, 1, 1), 0.08, 0.09, 2, 1]), 5.993775));
  it("MDURATION (2008..2016) = 5.73567", () =>
    near(F.MDURATION([D(2008, 1, 1), D(2016, 1, 1), 0.08, 0.09, 2, 1]), 5.73567, 1e-5));
  it("PRICEMAT = 99.98449888", () =>
    near(F.PRICEMAT([D(2008, 2, 15), D(2008, 4, 13), D(2007, 11, 11), 0.061, 0.061, 0]), 99.98449888));
  it("YIELDMAT = 0.060954", () =>
    near(F.YIELDMAT([D(2008, 3, 15), D(2008, 11, 3), D(2007, 11, 8), 0.0625, 100.0123, 0]), 0.060954, 1e-5));
  it("ACCRINTM = 20.54794521 (serial inputs 39539 / 39614)", () => {
    near(F.ACCRINTM([39539, 39614, 0.1, 1000, 3]), 20.54794521);
    near(F.ACCRINTM([D(2008, 4, 1), D(2008, 6, 15), 0.1, 1000, 3]), 20.54794521);
  });
  it("INTRATE = 0.05768", () =>
    near(F.INTRATE([D(2008, 2, 15), D(2008, 5, 15), 1000000, 1014420, 2]), 0.05768));
  it("RECEIVED = 1014584.654", () =>
    near(F.RECEIVED([D(2008, 2, 15), D(2008, 5, 15), 1000000, 0.0575, 2]), 1014584.654));
  it("YIELDDISC = 0.052823", () =>
    near(F.YIELDDISC([D(2008, 2, 16), D(2008, 3, 1), 99.795, 100, 2]), 0.052823, 1e-5));
});

describe("odd-period and VDB documentation examples", () => {
  it("ODDFPRICE = 113.597717 (docs: $113.60)", () =>
    near(F.ODDFPRICE([D(2008, 11, 11), D(2021, 3, 1), D(2008, 10, 15), D(2009, 3, 1), 0.0785, 0.0625, 100, 2, 1]), 113.597717, 1e-6));
  it("ODDFYIELD = 7.72%", () =>
    near(F.ODDFYIELD([D(2008, 11, 11), D(2021, 3, 1), D(2008, 10, 15), D(2009, 3, 1), 0.0575, 84.5, 100, 2, 0]), 0.0772, 1e-3));
  it("ODDFYIELD inverts ODDFPRICE", () => {
    const a = [D(2008, 11, 11), D(2021, 3, 1), D(2008, 10, 15), D(2009, 3, 1), 0.0785];
    const p = F.ODDFPRICE([...a, 0.0625, 100, 2, 1]) as number;
    near(F.ODDFYIELD([...a, p, 100, 2, 1]), 0.0625, 1e-9);
  });
  it("ODDFPRICE with a long first period inverts through ODDFYIELD", () => {
    for (const basis of [0, 1, 2, 3, 4]) {
      const a = [D(2008, 11, 11), D(2021, 3, 1), D(2008, 1, 15), D(2009, 3, 1), 0.0785];
      const p = F.ODDFPRICE([...a, 0.0625, 100, 2, basis]) as number;
      expect(p).toBeGreaterThan(100);
      near(F.ODDFYIELD([...a, p, 100, 2, basis]), 0.0625, 1e-9);
    }
  });
  it("ODDFPRICE equals PRICE when the first period is regular", () => {
    // issue at the start of a regular coupon period: identical cash flows to PRICE
    near(
      F.ODDFPRICE([D(2010, 4, 10), D(2015, 7, 15), D(2010, 1, 15), D(2010, 7, 15), 0.06, 0.05, 100, 2, 0]),
      F.PRICE([D(2010, 4, 10), D(2015, 7, 15), 0.06, 0.05, 100, 2, 0]) as number,
      1e-12,
    );
  });
  it("ODDLPRICE = 99.87829 (docs: $99.88)", () =>
    near(F.ODDLPRICE([D(2008, 2, 7), D(2008, 6, 15), D(2007, 10, 15), 0.0375, 0.0405, 100, 2, 0]), 99.87829, 1e-6));
  it("ODDLYIELD = 0.0451919 (docs: 0.04519)", () =>
    near(F.ODDLYIELD([D(2008, 4, 20), D(2008, 6, 15), D(2007, 12, 24), 0.0375, 99.875, 100, 2, 0]), 0.045192, 1e-4));
  it("ODDLYIELD inverts ODDLPRICE", () => {
    const a = [D(2008, 2, 7), D(2008, 6, 15), D(2007, 10, 15), 0.0375];
    const p = F.ODDLPRICE([...a, 0.0405, 100, 2, 0]) as number;
    near(F.ODDLYIELD([...a, p, 100, 2, 0]), 0.0405, 1e-9);
  });
  it("VDB documentation examples", () => {
    near(F.VDB([2400, 300, 3650, 0, 1]), 1.32, 5e-3);
    near(F.VDB([2400, 300, 120, 0, 1]), 40, 1e-9);
    near(F.VDB([2400, 300, 10, 0, 1]), 480, 1e-9);
    near(F.VDB([2400, 300, 120, 6, 18]), 396.31, 1e-5);
    near(F.VDB([2400, 300, 120, 6, 18, 1.5]), 311.81, 1e-5);
    near(F.VDB([2400, 300, 10, 0, 0.875, 1.5]), 315, 1e-9);
  });
  it("VDB no_switch keeps declining balance", () => {
    const sw = F.VDB([2400, 300, 10, 8, 10]) as number;
    const ns = F.VDB([2400, 300, 10, 8, 10, 2, true]) as number;
    expect(sw).toBeGreaterThan(ns);
  });
  it("odd-period and VDB errors", () => {
    const a = [D(2008, 11, 11), D(2021, 3, 1), D(2008, 10, 15), D(2009, 3, 1)];
    err(F.ODDFPRICE([a[0], a[1], a[0], a[3], 0.05, 0.05, 100, 2, 0]), "#NUM!"); // issue = settlement
    err(F.ODDFPRICE([a[0], a[3], a[2], a[1], 0.05, 0.05, 100, 2, 0]), "#NUM!"); // first coupon after maturity
    err(F.ODDFPRICE([...a, -0.01, 0.05, 100, 2, 0]), "#NUM!");
    err(F.ODDFPRICE([...a, 0.05, 0.05, 100, 3, 0]), "#NUM!");
    err(F.ODDFPRICE([...a, 0.05, 0.05, 100, 2, 9]), "#NUM!");
    err(F.ODDFYIELD([...a, 0.05, 0, 100, 2, 0]), "#NUM!");
    err(F.ODDFPRICE(["x", ...a.slice(1), 0.05, 0.05, 100, 2, 0]), "#VALUE!");
    err(F.ODDLPRICE([D(2007, 10, 15), D(2008, 6, 15), D(2008, 2, 7), 0.05, 0.05, 100, 2, 0]), "#NUM!");
    err(F.ODDLYIELD([D(2008, 4, 20), D(2008, 6, 15), D(2007, 12, 24), 0.05, 0, 100, 2, 0]), "#NUM!");
    err(F.VDB([2400, 300, 10, 5, 3]), "#NUM!");
    err(F.VDB([2400, 300, 10, 0, 11]), "#NUM!");
    err(F.VDB([-1, 300, 10, 0, 1]), "#NUM!");
    err(F.VDB([2400, 300, 10, 0, 1, 0]), "#NUM!");
    err(F.VDB([2400, 300, 10, 0]), "#VALUE!");
  });
});

describe("coupon schedule edge cases", () => {
  it("settlement on a coupon date", () => {
    const a = [D(2011, 5, 15), D(2011, 11, 15), 2, 0];
    expect(F.COUPPCD(a)).toBe(D(2011, 5, 15));
    expect(F.COUPNCD(a)).toBe(D(2011, 11, 15));
    expect(F.COUPDAYBS(a)).toBe(0);
    expect(F.COUPDAYSNC(a)).toBe(180);
    expect(F.COUPNUM(a)).toBe(1);
  });
  it("end-of-month maturity keeps end-of-month coupons", () => {
    expect(F.COUPPCD([D(2010, 10, 1), D(2011, 2, 28), 2, 0])).toBe(D(2010, 8, 31));
    expect(F.COUPNCD([D(2010, 10, 1), D(2011, 2, 28), 2, 0])).toBe(D(2011, 2, 28));
  });
  it("basis 3, 2 and 0 period lengths", () => {
    expect(F.COUPDAYS([D(2011, 1, 25), D(2011, 11, 15), 4, 3])).toBe(365 / 4);
    expect(F.COUPDAYS([D(2011, 1, 25), D(2011, 11, 15), 2, 2])).toBe(180);
    expect(F.COUPDAYS([D(2011, 1, 25), D(2011, 11, 15), 1, 0])).toBe(360);
  });
  it("truncates arguments", () => {
    expect(F.COUPNUM([D(2011, 1, 25) + 0.9, D(2011, 11, 15) + 0.9, 2.7, 1.9])).toBe(2);
  });
  it("uses the first element of an array", () => {
    expect(F.COUPNUM([[[D(2011, 1, 25)]], D(2011, 11, 15), 2, 1])).toBe(2);
  });
  it("PRICE at par when coupon = yield on a coupon date", () => {
    near(F.PRICE([D(2020, 1, 15), D(2025, 1, 15), 0.05, 0.05, 100, 2, 0]), 100, 1e-12);
  });
  it("YIELD inverts PRICE across bases and frequencies", () => {
    for (const basis of [0, 1, 2, 3, 4]) {
      for (const f of [1, 2, 4]) {
        const p = F.PRICE([D(2020, 3, 10), D(2031, 9, 30), 0.045, 0.071, 100, f, basis]) as number;
        near(F.YIELD([D(2020, 3, 10), D(2031, 9, 30), 0.045, p, 100, f, basis]), 0.071, 1e-9);
      }
    }
  });
  it("YIELD with a single remaining coupon uses the closed form", () => {
    const p = F.PRICE([D(2020, 1, 1), D(2020, 6, 30), 0.06, 0.08, 100, 2, 0]) as number;
    near(F.YIELD([D(2020, 1, 1), D(2020, 6, 30), 0.06, p, 100, 2, 0]), 0.08, 1e-9);
  });
  it("YIELD handles zero and very high yields", () => {
    near(F.PRICE([D(2020, 1, 15), D(2030, 1, 15), 0.05, 0, 100, 2, 0]), 150);
    expect(Math.abs(F.YIELD([D(2020, 1, 15), D(2030, 1, 15), 0.05, 150, 100, 2, 0]) as number)).toBeLessThan(1e-9);
    const hi = F.PRICE([D(2020, 1, 15), D(2030, 1, 15), 0.05, 5, 100, 2, 0]) as number;
    near(F.YIELD([D(2020, 1, 15), D(2030, 1, 15), 0.05, hi, 100, 2, 0]), 5, 1e-8);
  });
  it("MDURATION = DURATION / (1 + yld/freq)", () => {
    const a = [D(2008, 1, 1), D(2016, 1, 1), 0.08, 0.09, 2, 1];
    near(F.MDURATION(a), (F.DURATION(a) as number) / 1.045, 1e-12);
  });
});

describe("errors", () => {
  const s = D(2008, 2, 15);
  const m = D(2017, 11, 15);
  it("settlement on or after maturity", () => {
    err(F.COUPNUM([m, s, 2, 0]), "#NUM!");
    err(F.COUPNUM([s, s, 2, 0]), "#NUM!");
    err(F.PRICE([m, s, 0.05, 0.05, 100, 2, 0]), "#NUM!");
    err(F.YIELD([m, m, 0.05, 99, 100, 2, 0]), "#NUM!");
    err(F.DURATION([m, s, 0.05, 0.05, 2, 0]), "#NUM!");
    err(F.INTRATE([m, s, 1, 2, 0]), "#NUM!");
  });
  it("bad frequency or basis", () => {
    err(F.COUPNUM([s, m, 3, 0]), "#NUM!");
    err(F.COUPNUM([s, m, 0, 0]), "#NUM!");
    err(F.COUPNUM([s, m, 2, 5]), "#NUM!");
    err(F.COUPNUM([s, m, 2, -1]), "#NUM!");
    err(F.PRICE([s, m, 0.05, 0.05, 100, 12, 0]), "#NUM!");
    err(F.YIELDDISC([s, m, 99, 100, 7]), "#NUM!");
  });
  it("negative or zero amounts", () => {
    err(F.PRICE([s, m, -0.01, 0.05, 100, 2, 0]), "#NUM!");
    err(F.PRICE([s, m, 0.05, -0.05, 100, 2, 0]), "#NUM!");
    err(F.PRICE([s, m, 0.05, 0.05, 0, 2, 0]), "#NUM!");
    err(F.YIELD([s, m, 0.05, 0, 100, 2, 0]), "#NUM!");
    err(F.ACCRINTM([s, m, 0, 1000, 0]), "#NUM!");
    err(F.ACCRINTM([s, m, 0.1, 0, 0]), "#NUM!");
    err(F.INTRATE([s, m, 0, 100, 0]), "#NUM!");
    err(F.RECEIVED([s, m, 100, 0, 0]), "#NUM!");
    err(F.YIELDDISC([s, m, 0, 100, 0]), "#NUM!");
    err(F.DURATION([s, m, 0.05, -1, 2, 0]), "#NUM!");
    err(F.YIELDMAT([s, m, s - 100, 0.05, 0, 0]), "#NUM!");
  });
  it("RECEIVED with a discount that exhausts the investment", () => {
    err(F.RECEIVED([s, s + 400, 100, 5, 3]), "#NUM!");
  });
  it("#VALUE! for non-date text and missing arguments", () => {
    err(F.COUPNUM(["not a date", m, 2, 0]), "#VALUE!");
    err(F.PRICE([s, m, 0.05, 0.05, 100, 2, "x"]), "#VALUE!");
    err(F.COUPNUM([s, m]), "#VALUE!");
    err(F.PRICE([s, m, 0.05, 0.05, 100]), "#VALUE!");
    err(F.COUPNUM([s, m, 2, 0, 1]), "#VALUE!");
  });
  it("negative serial gives #NUM!", () => {
    err(F.COUPNUM([-5, m, 2, 0]), "#NUM!");
  });
  it("input errors propagate", () => {
    const e = new FormulaError("#N/A");
    expect(F.PRICE([e, m, 0.05, 0.05, 100, 2, 0])).toBe(e);
    expect(F.COUPNUM([s, m, 2, e])).toBe(e);
  });
  it("omitted and empty basis default to 0", () => {
    const a = [D(2011, 1, 25), D(2011, 11, 15), 2];
    expect(F.COUPDAYBS(a)).toBe(F.COUPDAYBS([...a, 0]));
    expect(F.COUPDAYBS([...a, null])).toBe(F.COUPDAYBS([...a, 0]));
  });
  it("ACCRINTM defaults par to 1000", () => {
    near(F.ACCRINTM([D(2008, 4, 1), D(2008, 6, 15), 0.1, undefined, 3]), 20.54794521);
  });
});
