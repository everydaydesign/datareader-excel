import type { XmlNode } from "./xml";
import { parseXml } from "./xml";

// numFmtIds whose builtin format renders as a date/time (ECMA-376 §18.8.30; matches SheetJS/ExcelJS).
// 14–22 + 45–47 are the Western date/time formats; 27–36 and 50–58 are the East-Asian (CJK) locale
// date/time builtins — omitting them mis-typed those files' date cells as raw serial numbers.
const BUILTIN_DATE = new Set<number>([
  14, 15, 16, 17, 18, 19, 20, 21, 22, 27, 28, 29, 30, 31, 32, 33, 34, 35, 36, 45, 46, 47, 50, 51,
  52, 53, 54, 55, 56, 57, 58,
]);

/** True if a custom number-format code denotes a date/time. */
function isDateFormatCode(code: string): boolean {
  // The system long-date / time formats put their date meaning in a bracketed token the strip below
  // would remove — [$-F800] = system long date, [$-F400] = system time — so detect them up front.
  if (/\[\$-F[48]00\]/i.test(code)) return true;
  // Otherwise strip quoted literals ("…"), bracketed tokens ([$-409], [Red], …), and backslash-escaped
  // chars, then look for a date/time token.
  const stripped = code
    .replace(/"[^"]*"/g, "")
    .replace(/\[[^\]]*\]/g, "")
    .replace(/\\./g, "");
  return /[ymdhs]/i.test(stripped);
}

/** Collect the numFmtIds of a <numFmts> section whose format code renders as a date. */
function addCustomDateFmtIds(section: XmlNode, dateFmtIds: Set<number>): void {
  for (const nf of section.children) {
    if (nf.name === "numFmt" && isDateFormatCode(nf.attrs.formatCode ?? "")) {
      const id = Number(nf.attrs.numFmtId);
      if (Number.isFinite(id)) dateFmtIds.add(id);
    }
  }
}

/** Collect the per-xf numFmtId of a <cellXfs> section, in style-index order. */
function addCellXfFmtIds(section: XmlNode, cellXfs: string[]): void {
  for (const xf of section.children) {
    if (xf.name === "xf") cellXfs.push(xf.attrs.numFmtId ?? "0");
  }
}

/** Parse xl/styles.xml → the set of cellXfs indices that render as dates. */
export function parseDateStyles(xml: string): Set<number> {
  const root = parseXml(xml);
  const dateFmtIds = new Set(BUILTIN_DATE);
  const cellXfs: string[] = [];
  for (const section of root.children) {
    if (section.name === "numFmts") addCustomDateFmtIds(section, dateFmtIds);
    if (section.name === "cellXfs") addCellXfFmtIds(section, cellXfs);
  }
  const out = new Set<number>();
  cellXfs.forEach((numFmtId, idx) => {
    if (dateFmtIds.has(Number(numFmtId))) out.add(idx);
  });
  return out;
}
