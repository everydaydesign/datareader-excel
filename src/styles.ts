import { parseXml } from "./xml";

const BUILTIN_DATE = new Set([14, 15, 16, 17, 18, 19, 20, 21, 22, 45, 46, 47]);

/** True if a custom number-format code denotes a date/time. Strip quoted literals ("…"), bracketed
 * tokens ([$-409], [Red], …), and backslash-escaped chars, then look for a date/time token. */
function isDateFormatCode(code: string): boolean {
  const stripped = code
    .replace(/"[^"]*"/g, "")
    .replace(/\[[^\]]*\]/g, "")
    .replace(/\\./g, "");
  return /[ymdhs]/i.test(stripped);
}

/** Parse xl/styles.xml → the set of cellXfs indices that render as dates. */
export function parseDateStyles(xml: string): Set<number> {
  const root = parseXml(xml);
  const dateFmtIds = new Set(BUILTIN_DATE);
  const cellXfs: string[] = [];
  for (const section of root.children) {
    if (section.name === "numFmts") {
      for (const nf of section.children) {
        if (nf.name === "numFmt" && isDateFormatCode(nf.attrs.formatCode ?? "")) {
          const id = Number(nf.attrs.numFmtId);
          if (Number.isFinite(id)) dateFmtIds.add(id);
        }
      }
    }
    if (section.name === "cellXfs") {
      for (const xf of section.children) {
        if (xf.name === "xf") cellXfs.push(xf.attrs.numFmtId ?? "0");
      }
    }
  }
  const out = new Set<number>();
  cellXfs.forEach((numFmtId, idx) => {
    if (dateFmtIds.has(Number(numFmtId))) out.add(idx);
  });
  return out;
}
