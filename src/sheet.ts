import type { CellValue } from "./types";
import type { XmlNode } from "./xml";
import { serialToDate } from "./dates";
import { XlsxError } from "./limits";
import { parseXml } from "./xml";

export type SheetCtx = {
  date1904: boolean;
  dates: boolean;
  dateStyles: Set<number>;
  maxCells: number;
  mergedCells: "topLeft" | "fill";
  sharedStrings: string[];
};

type MergeRange = { cA: number; cB: number; rA: number; rB: number };

type RowCells = { cells: CellValue[]; maxCol: number };

type SparseRows = { maxCol: number; maxRow: number; rowsByIndex: CellValue[][] };

/** Column letters of a cell ref (e.g. "AMJ12") → 0-based column index. */
export function colIndex(ref: string): number {
  let n = 0;
  for (let i = 0; i < ref.length; i++) {
    const c = ref.charCodeAt(i);
    if (c < 65 || c > 90) break; // stop at the first digit
    n = n * 26 + (c - 64);
  }
  return n - 1;
}

/** Row digits of a cell ref → 0-based row index ("C12" → 11). */
function rowIndex(ref: string): number {
  return Number(ref.replace(/^[A-Z]+/, "")) - 1;
}

/** Concatenate <t> text under a node (for <is> inline strings). Iterative pre-order walk (no
 * recursion, so deeply nested markup can't stack-overflow); skips <rPh>/<phoneticPr> so furigana
 * phonetic runs never corrupt the value (漢字, not 漢字かんじ). */
function inlineText(node: XmlNode): string {
  let out = "";
  const stack: XmlNode[] = [node];
  while (stack.length > 0) {
    const n = stack.pop();
    if (n === undefined) continue;
    if (n.name === "rPh" || n.name === "phoneticPr") continue;
    if (n.name === "t") {
      out += n.text;
      continue;
    }
    // Push children in reverse so pop yields left-to-right document order.
    for (let i = n.children.length - 1; i >= 0; i--) stack.push(n.children[i]!);
  }
  return out;
}

/** Decode a numeric <c> (t absent or "n"): a date-styled serial becomes a Date unless dates are off. */
function decodeNumeric(raw: string, c: XmlNode, ctx: SheetCtx): CellValue {
  const num = Number(raw);
  if (!Number.isFinite(num)) return raw;
  const styleIdx = Number(c.attrs.s ?? "0");
  if (ctx.dateStyles.has(styleIdx)) {
    return ctx.dates ? serialToDate(num, ctx.date1904) : num;
  }
  return num;
}

/** Decode a <c> that carries a <v> value: shared-string, "str", boolean, or numeric/date. */
function decodeValueCell(c: XmlNode, v: XmlNode, ctx: SheetCtx): CellValue {
  const t = c.attrs.t ?? "n";
  const raw = v.text;
  if (t === "s") {
    const idx = Number(raw);
    return ctx.sharedStrings[idx] ?? null;
  }
  if (t === "str") return raw;
  if (t === "b") return raw === "1";
  return decodeNumeric(raw, c, ctx);
}

/** Decode one <c> element to a CellValue. */
function decodeCell(c: XmlNode, ctx: SheetCtx): CellValue {
  const t = c.attrs.t ?? "n";
  if (t === "inlineStr") {
    const is = c.children.find((n) => n.name === "is");
    return is ? inlineText(is) : null;
  }
  if (t === "e") return null;
  const v = c.children.find((n) => n.name === "v");
  if (!v) return null;
  return decodeValueCell(c, v, ctx);
}

/** Parse <mergeCells> into 0-based inclusive ranges. */
function mergeRanges(sheet: XmlNode): MergeRange[] {
  const merges = sheet.children.find((n) => n.name === "mergeCells");
  const ranges: MergeRange[] = [];
  for (const mc of merges?.children ?? []) {
    const ref = mc.attrs.ref;
    if (!ref) continue;
    const [a, b] = ref.split(":");
    if (!a || !b) continue;
    ranges.push({ cA: colIndex(a), cB: colIndex(b), rA: rowIndex(a), rB: rowIndex(b) });
  }
  return ranges;
}

/** Apply merged-cell "fill": copy each range's top-left value across the range. */
function applyMerges(grid: CellValue[][], merges: MergeRange[]): void {
  for (const m of merges) {
    const val = grid[m.rA]?.[m.cA] ?? null;
    for (let r = m.rA; r <= m.rB; r++) {
      const gr = grid[r];
      if (!gr) continue;
      for (let c = m.cA; c <= m.cB; c++) gr[c] = val;
    }
  }
}

/** Row digits of a cell ref, defaulting to the next sequential row when the `r` attr is absent. */
function rowIndexOf(row: XmlNode, nextDefault: number): number {
  return (Number(row.attrs.r) || nextDefault) - 1;
}

/** Decode one <row>'s <c> children into a column-indexed sparse cell array + its column extent. A
 * cell with no `r` follows the previous one (ECMA-376: an omitted ref means the next column, mirroring
 * the sequential row default) instead of all collapsing onto column 0; a PRESENT ref that doesn't
 * parse to a valid column (e.g. lowercase "a1" → -1) is a malformed file, so it throws rather than
 * silently discarding the value at `cells[-1]`. */
function readRowCells(row: XmlNode, ctx: SheetCtx): RowCells {
  const cells: CellValue[] = [];
  let maxCol = 0;
  let nextCol = 0;
  for (const c of row.children) {
    if (c.name !== "c") continue;
    let ci = nextCol;
    if (c.attrs.r !== undefined) {
      ci = colIndex(c.attrs.r);
      if (ci < 0) throw new XlsxError(`malformed cell reference "${c.attrs.r}"`);
    }
    cells[ci] = decodeCell(c, ctx);
    nextCol = ci + 1;
    if (ci + 1 > maxCol) maxCol = ci + 1;
  }
  return { cells, maxCol };
}

/** Walk <sheetData> into row-indexed sparse cells, tracking the maximum row/column seen. */
function readSparseRows(sheet: XmlNode, ctx: SheetCtx): SparseRows {
  const sheetData = sheet.children.find((n) => n.name === "sheetData");
  const rowsByIndex: CellValue[][] = [];
  let maxCol = 0;
  let maxRow = 0;
  for (const row of sheetData?.children ?? []) {
    if (row.name !== "row") continue;
    const rIdx = rowIndexOf(row, rowsByIndex.length + 1);
    const { cells, maxCol: rowMaxCol } = readRowCells(row, ctx);
    if (rowMaxCol > maxCol) maxCol = rowMaxCol;
    rowsByIndex[rIdx] = cells;
    if (rIdx + 1 > maxRow) maxRow = rIdx + 1;
  }
  return { maxCol, maxRow, rowsByIndex };
}

/** Extend the (row, column) extent so a merge over otherwise-empty cells still materializes them. */
function mergedExtent(
  merges: MergeRange[],
  maxRow: number,
  maxCol: number,
): { cols: number; rows: number } {
  let rows = maxRow;
  let cols = maxCol;
  for (const m of merges) {
    if (m.rB + 1 > rows) rows = m.rB + 1;
    if (m.cB + 1 > cols) cols = m.cB + 1;
  }
  return { cols, rows };
}

/** Materialize the sparse rows into a dense `rows × cols` grid, filling gaps with null. */
function densify(rowsByIndex: CellValue[][], rows: number, cols: number): CellValue[][] {
  const grid: CellValue[][] = [];
  for (let r = 0; r < rows; r++) {
    const src = rowsByIndex[r] ?? [];
    const dense: CellValue[] = [];
    for (let c = 0; c < cols; c++) dense[c] = src[c] ?? null;
    grid[r] = dense;
  }
  return grid;
}

/** Parse a worksheet XML into a dense, row-major CellValue grid. */
export function parseSheet(xml: string, ctx: SheetCtx): CellValue[][] {
  const sheet = parseXml(xml);
  const { maxCol, maxRow, rowsByIndex } = readSparseRows(sheet, ctx);
  // A merge over otherwise-empty cells (e.g. B1:C1) makes those columns/rows exist, so it must
  // extend the grid extent even in "topLeft" mode — else the null "others" would be dropped.
  const merges = mergeRanges(sheet);
  const { cols, rows } = mergedExtent(merges, maxRow, maxCol);
  // Bound the extent BEFORE materializing the dense grid — maxRow/maxCol come from attacker-
  // controlled cell/merge refs, so a tiny file (e.g. <c r="A2000000000">) could otherwise OOM
  // allocating billions of slots before reader.ts's post-parse cumulative check runs.
  if (rows * cols > ctx.maxCells) {
    throw new XlsxError(
      `sheet extent ${rows}×${cols} exceeds the ${ctx.maxCells}-cell limit — pass a larger maxCells to raise it`,
    );
  }
  const grid = densify(rowsByIndex, rows, cols);
  if (ctx.mergedCells === "fill") applyMerges(grid, merges);
  return grid;
}
