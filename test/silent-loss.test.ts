import type { SheetCtx } from "../src/sheet";
import { describe, expect, test } from "bun:test";
import { readXlsx, XlsxError } from "../src/index";
import { parseSharedStrings } from "../src/sharedstrings";
import { parseSheet } from "../src/sheet";
import { parseDateStyles } from "../src/styles";
import { makeXlsx } from "./helpers/make-xlsx";

// The silent-data-loss classes the resource-focused adversarial suite doesn't cover (audit 2026-07-18).
const ctx = (over: Partial<SheetCtx> = {}): SheetCtx => ({
  date1904: false,
  dates: true,
  dateStyles: new Set(),
  maxCells: 5_000_000,
  mergedCells: "topLeft",
  sharedStrings: [],
  ...over,
});

describe("silent-loss fixes", () => {
  test("M1: a leading <!DOCTYPE> no longer zeroes the part", () => {
    expect(parseSharedStrings("<!DOCTYPE sst><sst><si><t>hello</t></si></sst>")).toEqual(["hello"]);
    // A DOCTYPE internal subset [ … ] can itself contain '>' — skipDeclaration must close past the ']'.
    expect(
      parseSharedStrings('<!DOCTYPE sst [<!ENTITY a "b">]><sst><si><t>hi</t></si></sst>'),
    ).toEqual(["hi"]);
  });

  test("M2: cells without an `r` follow the previous cell (not all column 0)", () => {
    const xml =
      "<worksheet><sheetData><row><c><v>1</v></c><c><v>2</v></c><c><v>3</v></c></row></sheetData></worksheet>";
    expect(parseSheet(xml, ctx())).toEqual([[1, 2, 3]]);
  });

  test("L1: East-Asian builtin (numFmtId 30) and system [$-F800] are detected as dates", () => {
    const styles =
      "<styleSheet>" +
      '<numFmts><numFmt numFmtId="164" formatCode="[$-F800]"/></numFmts>' +
      '<cellXfs><xf numFmtId="0"/><xf numFmtId="30"/><xf numFmtId="164"/></cellXfs>' +
      "</styleSheet>";
    const dateStyles = parseDateStyles(styles);
    expect(dateStyles.has(0)).toBe(false); // General — not a date
    expect(dateStyles.has(1)).toBe(true); // numFmtId 30 — East-Asian builtin date
    expect(dateStyles.has(2)).toBe(true); // custom [$-F800] — system long date
  });

  test("L2: a malformed (lowercase) cell ref throws instead of vanishing", () => {
    const xml = '<worksheet><sheetData><row><c r="a1"><v>42</v></c></row></sheetData></worksheet>';
    expect(() => parseSheet(xml, ctx())).toThrow(XlsxError);
  });

  test("L2: a workbook referencing a missing sheet part throws, naming the sheet", async () => {
    const buf = makeXlsx({
      "[Content_Types].xml": "<Types/>",
      "xl/_rels/workbook.xml.rels":
        '<Relationships><Relationship Id="rId1" Target="worksheets/sheet1.xml"/><Relationship Id="rId2" Target="worksheets/sheet2.xml"/></Relationships>',
      "xl/workbook.xml":
        '<workbook><sheets><sheet name="Real" r:id="rId1"/><sheet name="Ghost" r:id="rId2"/></sheets></workbook>',
      "xl/worksheets/sheet1.xml":
        '<worksheet><sheetData><row r="1"><c r="A1"><v>1</v></c></row></sheetData></worksheet>',
      // xl/worksheets/sheet2.xml intentionally absent — the ghost sheet.
    });
    let err: unknown;
    try {
      await readXlsx(buf);
    } catch (e) {
      err = e;
    }
    expect(err).toBeInstanceOf(XlsxError);
    expect(err instanceof Error && err.message.includes("Ghost")).toBe(true);
  });
});
