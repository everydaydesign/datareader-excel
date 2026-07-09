import { describe, expect, test } from "bun:test";
import { colIndex, parseSheet } from "../src/sheet";
import type { SheetCtx } from "../src/sheet";

const ctx = (over: Partial<SheetCtx> = {}): SheetCtx => ({
  sharedStrings: [],
  dateStyles: new Set(),
  date1904: false,
  dates: true,
  mergedCells: "topLeft",
  maxCells: 5_000_000,
  ...over,
});

describe("colIndex", () => {
  test("A/Z/AA/AMJ", () => {
    expect(colIndex("A1")).toBe(0);
    expect(colIndex("Z9")).toBe(25);
    expect(colIndex("AA1")).toBe(26);
    expect(colIndex("AMJ1")).toBe(1023);
  });
});

describe("parseSheet", () => {
  test("shared strings, numbers, booleans, errors, and gaps", () => {
    const xml =
      "<worksheet><sheetData>" +
      '<row r="1"><c r="A1" t="s"><v>0</v></c><c r="C1"><v>3.5</v></c></row>' +
      '<row r="2"><c r="A2" t="b"><v>1</v></c><c r="B2" t="e"><v>#N/A</v></c></row>' +
      "</sheetData></worksheet>";
    const grid = parseSheet(xml, ctx({ sharedStrings: ["hi"] }));
    expect(grid).toEqual([
      ["hi", null, 3.5],
      [true, null, null],
    ]);
  });

  test("a date-styled number decodes to a Date; dates:false keeps the serial", () => {
    const xml =
      '<worksheet><sheetData><row r="1"><c r="A1" s="1"><v>45657</v></c></row></sheetData></worksheet>';
    const asDate = parseSheet(xml, ctx({ dateStyles: new Set([1]) }))[0][0];
    expect(asDate instanceof Date).toBe(true);
    expect((asDate as Date).toISOString()).toBe("2024-12-31T00:00:00.000Z");
    const raw = parseSheet(xml, ctx({ dateStyles: new Set([1]), dates: false }))[0][0];
    expect(raw).toBe(45657);
  });

  test("inlineStr and formula cached value", () => {
    const xml =
      "<worksheet><sheetData>" +
      '<row r="1"><c r="A1" t="inlineStr"><is><t>x</t></is></c><c r="B1"><f>1+2</f><v>3</v></c></row>' +
      "</sheetData></worksheet>";
    expect(parseSheet(xml, ctx())).toEqual([["x", 3]]);
  });

  test('mergedCells "fill" propagates the top-left value across the range', () => {
    const xml =
      "<worksheet>" +
      '<sheetData><row r="1"><c r="A1" t="s"><v>0</v></c></row></sheetData>' +
      '<mergeCells><mergeCell ref="A1:C1"/></mergeCells>' +
      "</worksheet>";
    expect(parseSheet(xml, ctx({ sharedStrings: ["Region"], mergedCells: "fill" }))).toEqual([
      ["Region", "Region", "Region"],
    ]);
    expect(parseSheet(xml, ctx({ sharedStrings: ["Region"] }))).toEqual([["Region", null, null]]);
  });
});
