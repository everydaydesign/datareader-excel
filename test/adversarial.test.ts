import { describe, expect, test } from "bun:test";
import ExcelJS from "exceljs";
import { XlsxError, readXlsx } from "../src/index";
import { makeXlsx } from "./helpers/make-xlsx";

const SHEET = (rows: string) => `<worksheet><sheetData>${rows}</sheetData></worksheet>`;
const wbParts = (sheetXml: string) =>
  makeXlsx({
    "[Content_Types].xml": "<Types/>",
    "xl/workbook.xml": '<workbook><sheets><sheet name="S" r:id="rId1"/></sheets></workbook>',
    "xl/_rels/workbook.xml.rels":
      '<Relationships><Relationship Id="rId1" Target="worksheets/sheet1.xml"/></Relationships>',
    "xl/worksheets/sheet1.xml": sheetXml,
  });

describe("adversarial input → XlsxError (no hang/OOM)", () => {
  test("non-zip garbage", async () => {
    await expect(readXlsx(new Uint8Array(64).fill(0x41))).rejects.toBeInstanceOf(XlsxError);
  });
  test("OLE/encrypted container", async () => {
    const cfb = new Uint8Array(64);
    cfb.set([0xd0, 0xcf, 0x11, 0xe0]);
    await expect(readXlsx(cfb)).rejects.toThrow(/encrypted/i);
  });
  test("inflate byte-budget bomb", async () => {
    const wb = new ExcelJS.Workbook();
    const ws = wb.addWorksheet("B");
    for (let i = 0; i < 2000; i++) ws.addRow([i, i, i, i, i]);
    const buf = new Uint8Array(await wb.xlsx.writeBuffer());
    await expect(readXlsx(buf, { maxInflatedBytes: 1024 })).rejects.toBeInstanceOf(XlsxError);
  });
  test("cell budget", async () => {
    const wb = new ExcelJS.Workbook();
    const ws = wb.addWorksheet("C");
    for (let i = 0; i < 50; i++) ws.addRow([1, 2, 3, 4, 5, 6, 7, 8, 9, 10]);
    const buf = new Uint8Array(await wb.xlsx.writeBuffer());
    await expect(readXlsx(buf, { maxCells: 100 })).rejects.toThrow(/maxCells/);
  });
  // A tiny file whose cell/merge refs blow up maxRow/maxCol must be rejected BEFORE the dense grid
  // is allocated — else it OOMs. These must return FAST (no multi-GB allocation).
  test("huge row ref (maxRow≈2e9) rejects fast, no OOM", async () => {
    const buf = wbParts(SHEET('<row r="2000000000"><c r="A2000000000"><v>1</v></c></row>'));
    await expect(readXlsx(buf)).rejects.toBeInstanceOf(XlsxError);
  });
  test("huge column ref (ZZZZZZ1, maxCol≈3.2e8) rejects fast, no OOM", async () => {
    const buf = wbParts(SHEET('<row r="1"><c r="ZZZZZZ1"><v>1</v></c></row>'));
    await expect(readXlsx(buf)).rejects.toBeInstanceOf(XlsxError);
  });
});
