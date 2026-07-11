export type XmlNode = {
  attrs: Record<string, string>;
  children: XmlNode[];
  name: string;
  text: string;
};

const ENTITIES: Record<string, string> = {
  amp: "&",
  apos: "'",
  gt: ">",
  lt: "<",
  quot: '"',
};

/** Decode XML entities (&amp; &lt; &gt; &quot; &apos; and numeric &#dec; / &#xhex;). */
function decodeEntities(s: string): string {
  if (!s.includes("&")) return s;
  return s.replace(/&(#x?[0-9a-fA-F]+|[a-zA-Z]+);/g, (m, body: string) => {
    if (body[0] === "#") {
      const code =
        body[1] === "x" || body[1] === "X"
          ? parseInt(body.slice(2), 16)
          : parseInt(body.slice(1), 10);
      // String.fromCodePoint throws RangeError outside U+0000..U+10FFFF (e.g. &#x110000;); guard the
      // range and fall back to the literal so a hostile entity never crashes the parser.
      return Number.isFinite(code) && code >= 0 && code <= 0x10ffff
        ? String.fromCodePoint(code)
        : m;
    }
    const ent = ENTITIES[body];
    return ent === undefined ? m : ent;
  });
}

/** Strip a namespace prefix (`r:id` → `id`, `x:sheet` → `sheet`). */
function local(name: string): string {
  const i = name.indexOf(":");
  return i === -1 ? name : name.slice(i + 1);
}

/** Parse a start tag's attributes: name="value" (single or double quoted). */
function parseAttrs(src: string): Record<string, string> {
  const attrs: Record<string, string> = {};
  const re = /([^\s=/]+)\s*=\s*"([^"]*)"|([^\s=/]+)\s*=\s*'([^']*)'/g;
  let m: RegExpExecArray | null;
  while ((m = re.exec(src)) !== null) {
    const key = local(m[1] ?? m[3] ?? "");
    const val = decodeEntities(m[2] ?? m[4] ?? "");
    if (key) attrs[key] = val;
  }
  return attrs;
}

/** A scoped XML parser for well-formed OOXML — not a general engine. Builds an element tree with
 * decoded text/attributes and namespace prefixes stripped. Ignores <?…?>, comments, and CDATA fences
 * (their inner text is kept). */
export function parseXml(text: string): XmlNode {
  const root: XmlNode = { attrs: {}, children: [], name: "", text: "" };
  const stack: XmlNode[] = [root];
  let i = 0;
  const n = text.length;
  while (i < n) {
    const lt = text.indexOf("<", i);
    if (lt === -1) break;
    if (lt > i) {
      const chunk = text.slice(i, lt);
      const top = stack[stack.length - 1]!;
      top.text += decodeEntities(chunk);
    }
    if (text.startsWith("<?", lt)) {
      i = text.indexOf("?>", lt + 2);
      i = i === -1 ? n : i + 2;
      continue;
    }
    if (text.startsWith("<!--", lt)) {
      i = text.indexOf("-->", lt + 4);
      i = i === -1 ? n : i + 3;
      continue;
    }
    if (text.startsWith("<![CDATA[", lt)) {
      const end = text.indexOf("]]>", lt + 9);
      const stop = end === -1 ? n : end;
      stack[stack.length - 1]!.text += text.slice(lt + 9, stop);
      i = end === -1 ? n : end + 3;
      continue;
    }
    const gt = text.indexOf(">", lt);
    if (gt === -1) break;
    const raw = text.slice(lt + 1, gt);
    if (raw[0] === "/") {
      if (stack.length > 1) stack.pop();
      i = gt + 1;
      continue;
    }
    const selfClosing = raw.endsWith("/");
    const inner = selfClosing ? raw.slice(0, -1) : raw;
    const sp = inner.search(/\s/);
    const tag = local(sp === -1 ? inner : inner.slice(0, sp));
    const attrs = sp === -1 ? {} : parseAttrs(inner.slice(sp + 1));
    const node: XmlNode = { attrs, children: [], name: tag, text: "" };
    stack[stack.length - 1]!.children.push(node);
    if (!selfClosing) stack.push(node);
    i = gt + 1;
  }
  return root.children[0] ?? root;
}
