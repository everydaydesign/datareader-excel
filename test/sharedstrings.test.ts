import { describe, expect, test } from "bun:test";
import { parseSharedStrings } from "../src/sharedstrings";

describe("parseSharedStrings", () => {
  test("plain <si><t> strings", () => {
    const xml = "<sst><si><t>alpha</t></si><si><t>beta</t></si></sst>";
    expect(parseSharedStrings(xml)).toEqual(["alpha", "beta"]);
  });
  test("rich text: concatenates all <r><t> runs (ExcelJS #2761)", () => {
    const xml = "<sst><si><r><t>Hello </t></r><r><t>world</t></r></si></sst>";
    expect(parseSharedStrings(xml)).toEqual(["Hello world"]);
  });
  test("preserves significant whitespace (xml:space)", () => {
    const xml = '<sst><si><t xml:space="preserve"> x </t></si></sst>';
    expect(parseSharedStrings(xml)).toEqual([" x "]);
  });
  test("skips <rPh> phonetic (furigana) runs so CJK values don't corrupt", () => {
    const xml = "<sst><si><r><t>漢字</t></r><rPh><t>かんじ</t></rPh></si></sst>";
    expect(parseSharedStrings(xml)).toEqual(["漢字"]);
  });
  test("deeply nested markup does not stack-overflow (iterative walk)", () => {
    const depth = 20000;
    const xml = `<sst><si>${"<r>".repeat(depth)}<t>x</t>${"</r>".repeat(depth)}</si></sst>`;
    expect(parseSharedStrings(xml)).toEqual(["x"]);
  });
});
