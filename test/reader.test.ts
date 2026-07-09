import { describe, expect, test } from "bun:test";
import ExcelJS from "exceljs";
import { XlsxError, readXlsx } from "../src/index";

async function build(fn: (wb: ExcelJS.Workbook) => void): Promise<Uint8Array> {
  const wb = new ExcelJS.Workbook();
  fn(wb);
  return new Uint8Array(await wb.xlsx.writeBuffer());
}

describe("readXlsx", () => {
  test("reads a single sheet into a dense grid", async () => {
    const buf = await build((wb) => {
      const ws = wb.addWorksheet("Data");
      ws.addRow(["a", "b"]);
      ws.addRow([1, 2]);
    });
    const { sheets } = await readXlsx(buf);
    expect(sheets).toHaveLength(1);
    expect(sheets[0].name).toBe("Data");
    expect(sheets[0].rows).toEqual([
      ["a", "b"],
      [1, 2],
    ]);
  });

  test("enumerates multiple sheets in order", async () => {
    const buf = await build((wb) => {
      wb.addWorksheet("One").addRow([1]);
      wb.addWorksheet("Two").addRow([2]);
    });
    const { sheets } = await readXlsx(buf);
    expect(sheets.map((s) => s.name)).toEqual(["One", "Two"]);
  });

  test("a Date cell round-trips to a Date", async () => {
    const buf = await build((wb) => {
      wb.addWorksheet("D").addRow([new Date(Date.UTC(2024, 8, 13)), 5]);
    });
    const { sheets } = await readXlsx(buf);
    expect(sheets[0].rows[0][0] instanceof Date).toBe(true);
    expect(sheets[0].rows[0][1]).toBe(5);
  });

  test("rejects a non-xlsx input with XlsxError", async () => {
    await expect(readXlsx(new Uint8Array([1, 2, 3]))).rejects.toBeInstanceOf(XlsxError);
  });
});
