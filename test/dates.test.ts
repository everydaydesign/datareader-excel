import { describe, expect, test } from "bun:test";
import { serialToDate } from "../src/dates";

const iso = (d: Date) => d.toISOString();

describe("serialToDate (1900 system)", () => {
  test("serial 1 = 1900-01-01 (fixes ExcelJS #1928)", () => {
    expect(iso(serialToDate(1, false))).toBe("1900-01-01T00:00:00.000Z");
  });
  test("serial 59 = 1900-02-28 (before the phantom leap day)", () => {
    expect(iso(serialToDate(59, false))).toBe("1900-02-28T00:00:00.000Z");
  });
  test("serial 61 = 1900-03-01 (after the phantom leap day)", () => {
    expect(iso(serialToDate(61, false))).toBe("1900-03-01T00:00:00.000Z");
  });
  test("serial 45657 = 2024-12-31", () => {
    expect(iso(serialToDate(45657, false))).toBe("2024-12-31T00:00:00.000Z");
  });
  test("fractional part is time of day (0.5 = noon)", () => {
    expect(iso(serialToDate(45657.5, false))).toBe("2024-12-31T12:00:00.000Z");
  });
});

describe("serialToDate (1904 system)", () => {
  test("serial 0 = 1904-01-01", () => {
    expect(iso(serialToDate(0, true))).toBe("1904-01-01T00:00:00.000Z");
  });
  test("serial 1 = 1904-01-02", () => {
    expect(iso(serialToDate(1, true))).toBe("1904-01-02T00:00:00.000Z");
  });
});
