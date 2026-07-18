import type { CellValue } from "../src/index";
import { describe, expect, test } from "bun:test";
import ExcelJS from "exceljs";
import { readXlsx } from "../src/index";

function isPrimitiveCell(v: unknown): v is boolean | number | string {
  return typeof v === "number" || typeof v === "string" || typeof v === "boolean";
}

/** Normalize an ExcelJS object cell (formula result / hyperlink text / richText runs). */
function normalizeObject(v: object): CellValue {
  if ("result" in v) return normalize(v.result);
  if ("text" in v && typeof v.text === "string") return v.text;
  if ("richText" in v && Array.isArray(v.richText))
    return v.richText.map((r: { text?: string }) => r.text ?? "").join("");
  return null;
}

/** ExcelJS cell value → our CellValue (normalizing formula/richText/hyperlink; dates excluded here). */
function normalize(v: unknown): CellValue {
  if (v === null || v === undefined) return null;
  if (isPrimitiveCell(v)) return v;
  if (v instanceof Date) return v;
  if (typeof v === "object") return normalizeObject(v);
  return null;
}

// Return type is inferred from writeBuffer() (ExcelJS's own Buffer, a Uint8Array
// subclass) rather than annotated: ExcelJS's .d.ts references its bundled @types/node
// Buffer, which is a distinct type from this file's global @types/bun Buffer. Inferring
// keeps the value feeding wb.xlsx.load (wants that same Buffer) and readXlsx (wants a
// Uint8Array) without a cast; annotating either way reintroduces the cross-@types skew.
async function build(rowsPerSheet: unknown[][][]) {
  const wb = new ExcelJS.Workbook();
  rowsPerSheet.forEach((rows, i) => {
    const ws = wb.addWorksheet(`S${i + 1}`);
    rows.forEach((r) => ws.addRow(r));
  });
  return wb.xlsx.writeBuffer();
}

describe("differential vs ExcelJS (structure)", () => {
  test("strings, numbers, booleans, multi-sheet, gaps agree", async () => {
    const fixtures: unknown[][][][] = [
      [
        [
          ["name", "score", "ok"],
          ["alice", 12.5, true],
          ["bob", -3, false],
        ],
      ],
      [[["a"], [1], [1000000], [0.000123]]],
      [[["s1"]], [["s2"], [2]]],
    ];
    for (const rows of fixtures) {
      const buf = await build(rows);
      const wb = new ExcelJS.Workbook();
      await wb.xlsx.load(buf);
      const ours = await readXlsx(buf);
      expect(ours.sheets.length).toEqual(wb.worksheets.length);
      wb.eachSheet((ws, si) => {
        const sheet = ours.sheets[si - 1]!;
        ws.eachRow({ includeEmpty: true }, (row, rn) => {
          for (let c = 1; c <= (sheet.rows[0]?.length ?? 0); c++) {
            expect(sheet.rows[rn - 1]?.[c - 1] ?? null).toEqual(normalize(row.getCell(c).value));
          }
        });
      });
    }
  });
});
