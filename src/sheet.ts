import type { CellValue } from "./types";
import type { XmlNode } from "./xml";
import { serialToDate } from "./dates";
import { XlsxError } from "./limits";
import { parseXml } from "./xml";

export type SheetCtx = {
  sharedStrings: string[];
  dateStyles: Set<number>;
  date1904: boolean;
  dates: boolean;
  mergedCells: "topLeft" | "fill";
  maxCells: number;
};

type MergeRange = { rA: number; cA: number; rB: number; cB: number };

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
    for (let i = n.children.length - 1; i >= 0; i--) stack.push(n.children[i]);
  }
  return out;
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
  const raw = v.text;
  if (t === "s") {
    const idx = Number(raw);
    return ctx.sharedStrings[idx] ?? null;
  }
  if (t === "str") return raw;
  if (t === "b") return raw === "1";
  const num = Number(raw);
  if (!Number.isFinite(num)) return raw;
  const styleIdx = Number(c.attrs.s ?? "0");
  if (ctx.dateStyles.has(styleIdx)) {
    return ctx.dates ? serialToDate(num, ctx.date1904) : num;
  }
  return num;
}

/** Parse <mergeCells> into 0-based inclusive ranges. */
function mergeRanges(sheet: XmlNode): MergeRange[] {
  const merges = sheet.children.find((n) => n.name === "mergeCells");
  const ranges: MergeRange[] = [];
  for (const mc of merges?.children ?? []) {
    const ref = mc.attrs.ref;
    if (!ref) continue;
    const [a, b] = ref.split(":");
    if (!b) continue;
    ranges.push({ rA: rowIndex(a), cA: colIndex(a), rB: rowIndex(b), cB: colIndex(b) });
  }
  return ranges;
}

/** Apply merged-cell "fill": copy each range's top-left value across the range. */
function applyMerges(grid: CellValue[][], merges: MergeRange[]): void {
  for (const m of merges) {
    const val = grid[m.rA]?.[m.cA] ?? null;
    for (let r = m.rA; r <= m.rB; r++) {
      for (let c = m.cA; c <= m.cB; c++) {
        if (grid[r]) grid[r][c] = val;
      }
    }
  }
}

/** Parse a worksheet XML into a dense, row-major CellValue grid. */
export function parseSheet(xml: string, ctx: SheetCtx): CellValue[][] {
  const sheet = parseXml(xml);
  const sheetData = sheet.children.find((n) => n.name === "sheetData");
  const rowsByIndex: CellValue[][] = [];
  let maxCol = 0;
  let maxRow = 0;
  for (const row of sheetData?.children ?? []) {
    if (row.name !== "row") continue;
    const rIdx = (Number(row.attrs.r) || rowsByIndex.length + 1) - 1;
    const cells: CellValue[] = [];
    for (const c of row.children) {
      if (c.name !== "c") continue;
      const ci = colIndex(c.attrs.r ?? "A");
      cells[ci] = decodeCell(c, ctx);
      if (ci + 1 > maxCol) maxCol = ci + 1;
    }
    rowsByIndex[rIdx] = cells;
    if (rIdx + 1 > maxRow) maxRow = rIdx + 1;
  }
  // A merge over otherwise-empty cells (e.g. B1:C1) makes those columns/rows exist, so it must
  // extend the grid extent even in "topLeft" mode — else the null "others" would be dropped.
  const merges = mergeRanges(sheet);
  for (const m of merges) {
    if (m.rB + 1 > maxRow) maxRow = m.rB + 1;
    if (m.cB + 1 > maxCol) maxCol = m.cB + 1;
  }
  // Bound the extent BEFORE materializing the dense grid — maxRow/maxCol come from attacker-
  // controlled cell/merge refs, so a tiny file (e.g. <c r="A2000000000">) could otherwise OOM
  // allocating billions of slots before reader.ts's post-parse cumulative check runs.
  if (maxRow * maxCol > ctx.maxCells) {
    throw new XlsxError(
      `sheet extent ${maxRow}×${maxCol} exceeds the ${ctx.maxCells}-cell limit — pass a larger maxCells to raise it`,
    );
  }
  const grid: CellValue[][] = [];
  for (let r = 0; r < maxRow; r++) {
    const src = rowsByIndex[r] ?? [];
    const dense: CellValue[] = [];
    for (let c = 0; c < maxCol; c++) dense[c] = src[c] ?? null;
    grid[r] = dense;
  }
  if (ctx.mergedCells === "fill") applyMerges(grid, merges);
  return grid;
}
