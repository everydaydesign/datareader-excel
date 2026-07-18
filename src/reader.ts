import type { SheetCtx } from "./sheet";
import type { CellValue, ParsedFile, Sheet, XlsxOptions } from "./types";
import { DEFAULT_LIMITS, XlsxError } from "./limits";
import { parseSharedStrings } from "./sharedstrings";
import { parseSheet } from "./sheet";
import { parseDateStyles } from "./styles";
import { parseWorkbook } from "./workbook";
import { unzip } from "./zip";

type ResolvedXlsxOptions = {
  dates: boolean;
  maxCells: number;
  maxInflatedBytes: number;
  mergedCells: "fill" | "topLeft";
};

type SheetRef = { name: string; path: string };

function textOf(files: Map<string, Uint8Array>, path: string): string | undefined {
  const bytes = files.get(path);
  return bytes ? new TextDecoder("utf-8").decode(bytes) : undefined;
}

/** Read `opts` once so each field's `??` default doesn't also pay for an optional-chain on `opts`. */
function resolveXlsxOptions(opts?: XlsxOptions): ResolvedXlsxOptions {
  const o = opts ?? {};
  return {
    dates: o.dates ?? true,
    maxCells: o.maxCells ?? DEFAULT_LIMITS.maxCells,
    maxInflatedBytes: o.maxInflatedBytes ?? DEFAULT_LIMITS.maxInflatedBytes,
    mergedCells: o.mergedCells ?? "topLeft",
  };
}

/** Parse each referenced worksheet into a grid, enforcing the cumulative cell budget as it goes. */
function buildSheets(
  files: Map<string, Uint8Array>,
  sheetRefs: SheetRef[],
  ctx: SheetCtx,
): Sheet[] {
  const sheets: Sheet[] = [];
  let cellCount = 0;
  for (const ref of sheetRefs) {
    const xml = textOf(files, ref.path);
    // A worksheet the workbook references but whose part is absent is a malformed file — surface it
    // loudly (naming the sheet) rather than silently returning a workbook missing a whole sheet.
    if (xml === undefined) {
      throw new XlsxError(`worksheet "${ref.name}" is missing its part (${ref.path})`);
    }
    const rows: CellValue[][] = parseSheet(xml, ctx);
    cellCount += rows.length * (rows[0]?.length ?? 0);
    if (cellCount > ctx.maxCells) {
      throw new XlsxError(
        `workbook exceeds the ${ctx.maxCells}-cell limit — pass a larger maxCells to raise it`,
      );
    }
    sheets.push({ name: ref.name, rows });
  }
  return sheets;
}

/** Read a .xlsx/.xlsm (OOXML) file into a typed, multi-sheet grid. Async (ZIP inflation uses
 * DecompressionStream). Structure is faithful; date-styled serials decode to Date (opt out with
 * `dates: false`); merged values fill with `mergedCells: "fill"`. */
export async function readXlsx(
  input: ArrayBuffer | Uint8Array,
  opts?: XlsxOptions,
): Promise<ParsedFile> {
  const bytes = input instanceof Uint8Array ? input : new Uint8Array(input);
  const { dates, maxCells, maxInflatedBytes, mergedCells } = resolveXlsxOptions(opts);

  const files = await unzip(bytes, maxInflatedBytes);
  const wbXml = textOf(files, "xl/workbook.xml");
  if (!wbXml) throw new XlsxError("not a valid .xlsx (missing xl/workbook.xml)");
  const { date1904, sheets: sheetRefs } = parseWorkbook(
    wbXml,
    textOf(files, "xl/_rels/workbook.xml.rels"),
  );

  const sharedStrings = parseSharedStrings(textOf(files, "xl/sharedStrings.xml") ?? "<sst/>");
  const dateStyles = parseDateStyles(textOf(files, "xl/styles.xml") ?? "<styleSheet/>");

  const ctx: SheetCtx = { date1904, dates, dateStyles, maxCells, mergedCells, sharedStrings };
  return { format: "xlsx", sheets: buildSheets(files, sheetRefs, ctx) };
}
