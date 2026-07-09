import { describe, expect, test } from "bun:test";
import { DEFAULT_LIMITS, XlsxError } from "../src/limits";

describe("limits", () => {
  test("XlsxError is a catchable Error subclass", () => {
    const e = new XlsxError("bad");
    expect(e).toBeInstanceOf(Error);
    expect(e).toBeInstanceOf(XlsxError);
    expect(e.message).toBe("bad");
  });
  test("DEFAULT_LIMITS", () => {
    expect(DEFAULT_LIMITS.maxCells).toBe(5_000_000);
    expect(DEFAULT_LIMITS.maxInflatedBytes).toBe(512 * 1024 * 1024);
  });
});
