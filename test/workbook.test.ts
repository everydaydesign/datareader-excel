import { describe, expect, test } from "bun:test";
import { parseWorkbook } from "../src/workbook";

const WB =
  '<workbook><workbookPr date1904="1"/><sheets>' +
  '<sheet name="First" sheetId="1" r:id="rId1"/>' +
  '<sheet name="Second" sheetId="2" r:id="rId2"/>' +
  "</sheets></workbook>";
const RELS =
  "<Relationships>" +
  '<Relationship Id="rId1" Target="worksheets/sheet1.xml"/>' +
  '<Relationship Id="rId2" Target="worksheets/sheet2.xml"/>' +
  "</Relationships>";

describe("parseWorkbook", () => {
  test("ordered sheets resolved via rels + date1904 flag", () => {
    const wb = parseWorkbook(WB, RELS);
    expect(wb.date1904).toBe(true);
    expect(wb.sheets).toEqual([
      { name: "First", path: "xl/worksheets/sheet1.xml" },
      { name: "Second", path: "xl/worksheets/sheet2.xml" },
    ]);
  });

  test("tolerates missing rels + workbookPr (ExcelJS #1329)", () => {
    const wb = parseWorkbook(
      '<workbook><sheets><sheet name="Only" r:id="rId1"/></sheets></workbook>',
      undefined,
    );
    expect(wb.date1904).toBe(false);
    expect(wb.sheets).toEqual([{ name: "Only", path: "xl/worksheets/sheet1.xml" }]);
  });
});
