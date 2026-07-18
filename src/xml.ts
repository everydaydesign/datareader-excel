export type XmlNode = {
  attrs: Record<string, string>;
  children: XmlNode[];
  name: string;
  text: string;
};

type ParseCursor = { n: number; stack: XmlNode[]; text: string };

const ENTITIES: Record<string, string> = {
  amp: "&",
  apos: "'",
  gt: ">",
  lt: "<",
  quot: '"',
};

/** Decode a numeric character reference body (`#dec` or `#xhex`), falling back to `fallback` when the
 * code point is out of U+0000..U+10FFFF — String.fromCodePoint throws a RangeError past it (e.g.
 * &#x110000;), so a hostile entity must never crash the parser. */
function decodeNumericEntity(body: string, fallback: string): string {
  const code =
    body[1] === "x" || body[1] === "X" ? parseInt(body.slice(2), 16) : parseInt(body.slice(1), 10);
  return Number.isFinite(code) && code >= 0 && code <= 0x10ffff
    ? String.fromCodePoint(code)
    : fallback;
}

/** Decode XML entities (&amp; &lt; &gt; &quot; &apos; and numeric &#dec; / &#xhex;). */
function decodeEntities(s: string): string {
  if (!s.includes("&")) return s;
  return s.replace(/&(#x?[0-9a-fA-F]+|[a-zA-Z]+);/g, (m, body: string) => {
    if (body[0] === "#") return decodeNumericEntity(body, m);
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

/** Append decoded character data to the element currently on top of the stack. */
function appendText(stack: XmlNode[], chunk: string): void {
  stack[stack.length - 1]!.text += decodeEntities(chunk);
}

/** Index just past the next `close` delimiter at/after `from`, or the end of input if it's absent. */
function endAfter(cursor: ParseCursor, from: number, close: string): number {
  const j = cursor.text.indexOf(close, from);
  return j === -1 ? cursor.n : j + close.length;
}

/** Consume a <![CDATA[…]]> block: keep its inner text on the stack top, return the index past it. */
function consumeCData(cursor: ParseCursor, lt: number): number {
  const { n, stack, text } = cursor;
  const end = text.indexOf("]]>", lt + 9);
  const stop = end === -1 ? n : end;
  stack[stack.length - 1]!.text += text.slice(lt + 9, stop);
  return end === -1 ? n : end + 3;
}

/** Skip a `<!…>` declaration (DOCTYPE/ENTITY/…), starting at `lt`. A DOCTYPE's internal subset
 * `[ … ]` may itself contain `>`, so when a `[` opens before the next `>`, close past its `]` first.
 * Without this, `<!DOCTYPE sst>` parses as an element and becomes the document root, silently zeroing
 * every string/style/cell in the part (some third-party generators emit a DOCTYPE — it's legal XML). */
function skipDeclaration(cursor: ParseCursor, lt: number): number {
  const { text } = cursor;
  const gt = text.indexOf(">", lt);
  const bracket = text.indexOf("[", lt);
  if (bracket !== -1 && (gt === -1 || bracket < gt)) {
    const close = text.indexOf("]", bracket);
    return endAfter(cursor, close === -1 ? lt + 2 : close, ">");
  }
  return endAfter(cursor, lt + 2, ">");
}

/** Skip a non-element construct at `lt` (<?…?>, <!--…-->, <![CDATA[…]]>, <!DOCTYPE …>) and return the
 * next index, or null when `lt` begins an ordinary element tag. */
function skipSpecial(cursor: ParseCursor, lt: number): number | null {
  const { text } = cursor;
  if (text.startsWith("<?", lt)) return endAfter(cursor, lt + 2, "?>");
  if (text.startsWith("<!--", lt)) return endAfter(cursor, lt + 4, "-->");
  if (text.startsWith("<![CDATA[", lt)) return consumeCData(cursor, lt);
  if (text.startsWith("<!", lt)) return skipDeclaration(cursor, lt);
  return null;
}

/** Handle one element tag spanning `lt`..`gt`: pop on a close tag, else push a new node (unless
 * self-closing) carrying its decoded attributes. */
function processTag(cursor: ParseCursor, lt: number, gt: number): void {
  const { stack, text } = cursor;
  const raw = text.slice(lt + 1, gt);
  if (raw[0] === "/") {
    if (stack.length > 1) stack.pop();
    return;
  }
  const selfClosing = raw.endsWith("/");
  const inner = selfClosing ? raw.slice(0, -1) : raw;
  const sp = inner.search(/\s/);
  const tag = local(sp === -1 ? inner : inner.slice(0, sp));
  const attrs = sp === -1 ? {} : parseAttrs(inner.slice(sp + 1));
  const node: XmlNode = { attrs, children: [], name: tag, text: "" };
  stack[stack.length - 1]!.children.push(node);
  if (!selfClosing) stack.push(node);
}

/** A scoped XML parser for well-formed OOXML — not a general engine. Builds an element tree with
 * decoded text/attributes and namespace prefixes stripped. Ignores <?…?>, comments, and CDATA fences
 * (their inner text is kept). */
export function parseXml(text: string): XmlNode {
  const root: XmlNode = { attrs: {}, children: [], name: "", text: "" };
  const stack: XmlNode[] = [root];
  const n = text.length;
  const cursor: ParseCursor = { n, stack, text };
  let i = 0;
  while (i < n) {
    const lt = text.indexOf("<", i);
    if (lt === -1) break;
    if (lt > i) appendText(stack, text.slice(i, lt));
    const skipped = skipSpecial(cursor, lt);
    if (skipped !== null) {
      i = skipped;
      continue;
    }
    const gt = text.indexOf(">", lt);
    if (gt === -1) break;
    processTag(cursor, lt, gt);
    i = gt + 1;
  }
  return root.children[0] ?? root;
}
