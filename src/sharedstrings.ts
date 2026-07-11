import type { XmlNode } from "./xml";
import { parseXml } from "./xml";

/** Concatenate the text of every <t> descendant of a node (handles rich-text <r><t> runs).
 * Iterative pre-order walk (no recursion, so deeply nested markup can't stack-overflow); skips
 * <rPh>/<phoneticPr> so furigana phonetic runs never corrupt the value (漢字, not 漢字かんじ). */
function collectText(node: XmlNode): string {
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

/** Parse xl/sharedStrings.xml → the shared-string table (one entry per <si>). */
export function parseSharedStrings(xml: string): string[] {
  const root = parseXml(xml);
  const out: string[] = [];
  for (const si of root.children) {
    if (si.name === "si") out.push(collectText(si));
  }
  return out;
}
