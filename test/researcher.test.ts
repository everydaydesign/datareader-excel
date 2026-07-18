import { describe, expect, test } from "bun:test";
import ExcelJS from "exceljs";
import { readXlsx } from "../src/index";
import { makeXlsx } from "./helpers/make-xlsx";

const SHEET = (rows: string) => `<worksheet><sheetData>${rows}</sheetData></worksheet>`;
const wbParts = (sheetXml: string, extra: Record<string, string> = {}) =>
  makeXlsx({
    "[Content_Types].xml": "<Types/>",
    "xl/_rels/workbook.xml.rels":
      '<Relationships><Relationship Id="rId1" Target="worksheets/sheet1.xml"/></Relationships>',
    "xl/workbook.xml": '<workbook><sheets><sheet name="S" r:id="rId1"/></sheets></workbook>',
    "xl/worksheets/sheet1.xml": sheetXml,
    ...extra,
  });

describe("researcher guarantees", () => {
  test("a percent-formatted cell reads the RAW value 0.45, not '45%'", async () => {
    const wb = new ExcelJS.Workbook();
    const ws = wb.addWorksheet("P");
    ws.getCell("A1").value = 0.45;
    ws.getCell("A1").numFmt = "0%";
    const buf = new Uint8Array(await wb.xlsx.writeBuffer());
    const { sheets } = await readXlsx(buf);
    expect(sheets[0]!.rows[0]![0]).toBe(0.45);
  });

  test("full double precision is preserved", async () => {
    const buf = wbParts(SHEET('<row r="1"><c r="A1"><v>0.1234567890123</v></c></row>'));
    const { sheets } = await readXlsx(buf);
    expect(sheets[0]!.rows[0]![0]).toBe(0.1234567890123);
  });

  test("blank and #N/A error cells both read null", async () => {
    const buf = wbParts(
      SHEET('<row r="1"><c r="A1"/><c r="B1" t="e"><v>#N/A</v></c><c r="C1"><v>1</v></c></row>'),
    );
    const { sheets } = await readXlsx(buf);
    expect(sheets[0]!.rows[0]).toEqual([null, null, 1]);
  });

  test("wide sheet: a multi-letter column ref (AMJ1) reads correctly", async () => {
    const buf = wbParts(SHEET('<row r="1"><c r="A1"><v>1</v></c><c r="AMJ1"><v>2</v></c></row>'));
    const { sheets } = await readXlsx(buf);
    expect(sheets[0]!.rows[0]![0]).toBe(1);
    expect(sheets[0]!.rows[0]![1023]).toBe(2);
  });

  test("unicode (Hebrew/CJK) cell values round-trip", async () => {
    const buf = wbParts(
      SHEET('<row r="1"><c r="A1" t="inlineStr"><is><t>שלום 世界</t></is></c></row>'),
    );
    const { sheets } = await readXlsx(buf);
    expect(sheets[0]!.rows[0]![0]).toBe("שלום 世界");
  });

  test("a workbook missing styles/sharedStrings/workbookPr reads without throwing (#1329)", async () => {
    const buf = wbParts(SHEET('<row r="1"><c r="A1"><v>7</v></c></row>'));
    const { sheets } = await readXlsx(buf);
    expect(sheets[0]!.rows[0]![0]).toBe(7);
  });
});
