import type { CellValue, ParsedFile, Sheet, XlsxOptions } from "./types";
import type { SheetCtx } from "./sheet";
import { DEFAULT_LIMITS, XlsxError } from "./limits";
import { parseSharedStrings } from "./sharedstrings";
import { parseSheet } from "./sheet";
import { parseDateStyles } from "./styles";
import { parseWorkbook } from "./workbook";
import { unzip } from "./zip";

function textOf(files: Map<string, Uint8Array>, path: string): string | undefined {
  const bytes = files.get(path);
  return bytes ? new TextDecoder("utf-8").decode(bytes) : undefined;
}

/** Read a .xlsx/.xlsm (OOXML) file into a typed, multi-sheet grid. Async (ZIP inflation uses
 * DecompressionStream). Structure is faithful; date-styled serials decode to Date (opt out with
 * `dates: false`); merged values fill with `mergedCells: "fill"`. */
export async function readXlsx(
  input: ArrayBuffer | Uint8Array,
  opts?: XlsxOptions,
): Promise<ParsedFile> {
  const bytes = input instanceof Uint8Array ? input : new Uint8Array(input);
  const dates = opts?.dates ?? true;
  const mergedCells = opts?.mergedCells ?? "topLeft";
  const maxCells = opts?.maxCells ?? DEFAULT_LIMITS.maxCells;
  const maxInflatedBytes = opts?.maxInflatedBytes ?? DEFAULT_LIMITS.maxInflatedBytes;

  const files = await unzip(bytes, maxInflatedBytes);
  const wbXml = textOf(files, "xl/workbook.xml");
  if (!wbXml) throw new XlsxError("not a valid .xlsx (missing xl/workbook.xml)");
  const { sheets: sheetRefs, date1904 } = parseWorkbook(
    wbXml,
    textOf(files, "xl/_rels/workbook.xml.rels"),
  );

  const sharedStrings = parseSharedStrings(textOf(files, "xl/sharedStrings.xml") ?? "<sst/>");
  const dateStyles = parseDateStyles(textOf(files, "xl/styles.xml") ?? "<styleSheet/>");

  const sheets: Sheet[] = [];
  let cellCount = 0;
  for (const ref of sheetRefs) {
    const xml = textOf(files, ref.path);
    if (xml === undefined) continue;
    const ctx: SheetCtx = { sharedStrings, dateStyles, date1904, dates, mergedCells, maxCells };
    const rows: CellValue[][] = parseSheet(xml, ctx);
    cellCount += rows.length * (rows[0]?.length ?? 0);
    if (cellCount > maxCells) {
      throw new XlsxError(
        `workbook exceeds the ${maxCells}-cell limit — pass a larger maxCells to raise it`,
      );
    }
    sheets.push({ name: ref.name, rows });
  }
  return { format: "xlsx", sheets };
}
