import { describe, expect, test } from "bun:test";
import ExcelJS from "exceljs";
import { XlsxError } from "../src/limits";
import { unzip } from "../src/zip";
import { makeXlsx } from "./helpers/make-xlsx";

const LIMIT = 512 * 1024 * 1024;
const dec = (b: Uint8Array) => new TextDecoder().decode(b);

describe("unzip", () => {
  test("reads stored (uncompressed) entries", async () => {
    const zip = makeXlsx({ "a.txt": "hello", "dir/b.xml": "<r/>" });
    const map = await unzip(zip, LIMIT);
    expect(dec(map.get("a.txt")!)).toBe("hello");
    expect(dec(map.get("dir/b.xml")!)).toBe("<r/>");
  });

  test("inflates a real deflated .xlsx from ExcelJS (method 8)", async () => {
    const wb = new ExcelJS.Workbook();
    wb.addWorksheet("Data").addRow(["a", 1]);
    const buf = new Uint8Array(await wb.xlsx.writeBuffer());
    const map = await unzip(buf, LIMIT);
    expect(map.has("xl/workbook.xml")).toBe(true);
    expect(dec(map.get("xl/workbook.xml")!)).toContain("<workbook");
  });

  test("rejects a non-ZIP input", async () => {
    await expect(unzip(new Uint8Array([1, 2, 3, 4]), LIMIT)).rejects.toBeInstanceOf(XlsxError);
  });

  test("rejects an OLE/encrypted container (CFB signature)", async () => {
    const cfb = new Uint8Array([0xd0, 0xcf, 0x11, 0xe0, 0xa1, 0xb1, 0x1a, 0xe1]);
    await expect(unzip(cfb, LIMIT)).rejects.toThrow(/encrypted/i);
  });

  test("rejects inflation past the byte budget", async () => {
    const wb = new ExcelJS.Workbook();
    const ws = wb.addWorksheet("Big");
    for (let i = 0; i < 500; i++) ws.addRow([i, i, i, i]);
    const buf = new Uint8Array(await wb.xlsx.writeBuffer());
    await expect(unzip(buf, 256)).rejects.toBeInstanceOf(XlsxError);
  });
});
