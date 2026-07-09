import { describe, expect, test } from "bun:test";
import { parseXml } from "../src/xml";

describe("parseXml", () => {
  test("elements, attributes, and text", () => {
    const root = parseXml('<a x="1"><b>hi</b><c/></a>');
    expect(root.name).toBe("a");
    expect(root.attrs.x).toBe("1");
    expect(root.children.map((n) => n.name)).toEqual(["b", "c"]);
    expect(root.children[0].text).toBe("hi");
  });
  test("decodes entities in text and attributes", () => {
    const root = parseXml('<t a="x &amp; y">1 &lt; 2 &#65; &#x42;</t>');
    expect(root.attrs.a).toBe("x & y");
    expect(root.text).toBe("1 < 2 A B");
  });
  test("strips namespace prefixes on tags and attributes", () => {
    const root = parseXml('<x:sheet r:id="rId1" name="Data"/>');
    expect(root.name).toBe("sheet");
    expect(root.attrs.id).toBe("rId1");
    expect(root.attrs.name).toBe("Data");
  });
  test("ignores the XML declaration and comments", () => {
    const root = parseXml('<?xml version="1.0"?><!-- c --><r><v>5</v></r>');
    expect(root.name).toBe("r");
    expect(root.children[0].text).toBe("5");
  });
  test("preserves significant whitespace in text", () => {
    const root = parseXml('<t xml:space="preserve"> a </t>');
    expect(root.text).toBe(" a ");
  });
  test("an out-of-range numeric entity falls back to the literal (no RangeError)", () => {
    expect(parseXml("<t>&#x110000;</t>").text).toBe("&#x110000;");
  });
});
