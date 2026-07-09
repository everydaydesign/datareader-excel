import { parseXml } from "./xml";

/** rId → worksheet path (under xl/) from xl/_rels/workbook.xml.rels. */
function relMap(relsXml: string | undefined): Map<string, string> {
  const map = new Map<string, string>();
  if (!relsXml) return map;
  const root = parseXml(relsXml);
  for (const rel of root.children) {
    if (rel.name === "Relationship" && rel.attrs.Id && rel.attrs.Target) {
      const target = rel.attrs.Target.replace(/^\/?xl\//, "").replace(/^\//, "");
      map.set(rel.attrs.Id, `xl/${target}`);
    }
  }
  return map;
}

/** Parse xl/workbook.xml (+ its rels) → ordered sheets and the date1904 flag. Missing rels /
 * workbookPr are tolerated with sensible fallbacks (#1329). */
export function parseWorkbook(
  wbXml: string,
  relsXml: string | undefined,
): { sheets: { name: string; path: string }[]; date1904: boolean } {
  const root = parseXml(wbXml);
  const rels = relMap(relsXml);
  let date1904 = false;
  const sheets: { name: string; path: string }[] = [];
  for (const section of root.children) {
    if (section.name === "workbookPr") {
      const v = section.attrs.date1904;
      date1904 = v === "1" || v === "true";
    }
    if (section.name === "sheets") {
      section.children.forEach((sheet, i) => {
        if (sheet.name !== "sheet") return;
        const name = sheet.attrs.name ?? `Sheet${i + 1}`;
        const path = rels.get(sheet.attrs.id ?? "") ?? `xl/worksheets/sheet${i + 1}.xml`;
        sheets.push({ name, path });
      });
    }
  }
  return { sheets, date1904 };
}
