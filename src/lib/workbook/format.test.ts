import { describe, expect, it } from "vitest";
import { formatNumberForDisplay } from "./format";

// formatNumberForDisplay caches its Intl.NumberFormat instances; the output must stay what
// toLocaleString("en-US", ...) gives.
describe("formatNumberForDisplay", () => {
  const reference = (value: number, decimals: number, grouping: boolean) =>
    value.toLocaleString("en-US", { minimumFractionDigits: decimals, maximumFractionDigits: decimals, useGrouping: grouping });

  it("matches toLocaleString for the formats the importer sees", () => {
    const values = [0, 1, -1, 0.5, 1234.5678, -98765.4321, 1e9 + 0.129, 0.000123];
    for (const value of values) {
      expect(formatNumberForDisplay(value, "#,##0.00")).toBe(reference(value, 2, true));
      expect(formatNumberForDisplay(value, "0.0")).toBe(reference(value, 1, false));
      expect(formatNumberForDisplay(value, "0")).toBe(reference(value, 0, false));
      expect(formatNumberForDisplay(value, "0.00%")).toBe(`${reference(value * 100, 2, false)}%`);
      expect(formatNumberForDisplay(value, "$#,##0.000")).toBe(`$${reference(value, 3, true)}`);
    }
  });

  it("handles General, text and date formats without the formatter", () => {
    expect(formatNumberForDisplay(12.5, undefined)).toBe("12.5");
    expect(formatNumberForDisplay(12.5, "General")).toBe("12.5");
    expect(formatNumberForDisplay(45000, "yyyy-mm-dd")).toBe("2023-03-15");
  });
});
