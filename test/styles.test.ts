import { describe, expect, test } from "bun:test";
import { parseDateStyles } from "../src/styles";

describe("parseDateStyles", () => {
  test("builtin date numFmtIds (14) are dates", () => {
    const xml =
      '<styleSheet><cellXfs count="2"><xf numFmtId="0"/><xf numFmtId="14"/></cellXfs></styleSheet>';
    const set = parseDateStyles(xml);
    expect(set.has(0)).toBe(false);
    expect(set.has(1)).toBe(true);
  });
  test("custom formatCode with date tokens is a date; a currency format is not", () => {
    const xml =
      "<styleSheet>" +
      '<numFmts><numFmt numFmtId="164" formatCode="yyyy-mm-dd"/><numFmt numFmtId="165" formatCode="&quot;$&quot;#,##0.00"/></numFmts>' +
      '<cellXfs count="2"><xf numFmtId="164"/><xf numFmtId="165"/></cellXfs>' +
      "</styleSheet>";
    const set = parseDateStyles(xml);
    expect(set.has(0)).toBe(true); // yyyy-mm-dd
    expect(set.has(1)).toBe(false); // currency
  });
  test("a currency format that literally contains a letter in a quoted string is not a date", () => {
    const xml =
      '<styleSheet><numFmts><numFmt numFmtId="166" formatCode="&quot;Day&quot;#,##0"/></numFmts>' +
      '<cellXfs count="1"><xf numFmtId="166"/></cellXfs></styleSheet>';
    expect(parseDateStyles(xml).has(0)).toBe(false);
  });
});
