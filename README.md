# datareader-excel

**Correct, zero-dependency Excel .xlsx/.xlsm reader for the browser and Node.**

`datareader-excel` reads modern Excel `.xlsx`/`.xlsm` (OOXML / SpreadsheetML)
files into a typed, row-major grid of plain JavaScript values: strings,
numbers, booleans, `Date`s, and `null` for blank or error cells. It is written
in pure TypeScript against Web platform APIs (`DecompressionStream`,
`TextDecoder`) — **no runtime dependencies**, and the same build runs in the
browser and in Node. Because ZIP inflation is async, `readXlsx` returns a
`Promise`.

## Why

The de-facto Excel reader, ExcelJS, is a ~1 MB read+write+styling+charts
library whose last release was October 2023 and last commit January 2025, and
which now carries a cluster of unpatched transitive-dependency CVEs (minimatch,
uuid, tmp, archiver/inflight). If all you need is to *read* the data out of a
spreadsheet, that is a large, stalled dependency and a standing attack surface
to take on.

`datareader-excel` reads the data and nothing else. Zero runtime dependencies
means no transitive-CVE class exists by construction; a hand-written ZIP +
OOXML reader keeps it small and strictly read-only; it is TypeScript-native (no
`@types/*` to chase); and it decodes Excel date serials to `Date` correctly —
including the boundary cases ExcelJS gets wrong — so it drops straight into a
data grid or a statistics pipeline. It deliberately targets the specific
read-side defects in ExcelJS's stalled backlog (see
[Better than ExcelJS](#better-than-exceljs)).

## Install

```bash
npm i datareader-excel
```

```bash
bun add datareader-excel
pnpm add datareader-excel
yarn add datareader-excel
```

## Usage

### Browser — a picked/dropped `File`

```ts
import { readXlsx } from "datareader-excel";

const input =
  document.querySelector<HTMLInputElement>("#file");

input.addEventListener("change", async () => {
  const file = input.files?.[0];
  if (!file) return;

  // File.arrayBuffer() and readXlsx are both async.
  const { sheets } = await readXlsx(
    await file.arrayBuffer(),
  );

  console.log(sheets.length); // worksheet count
  console.log(sheets[0].name); // first sheet name
  console.log(sheets[0].rows[0]); // its first row
});
```

### React — a file input

```tsx
import { useState } from "react";
import {
  readXlsx,
  type CellValue,
} from "datareader-excel";

export function XlsxImporter() {
  const [rows, setRows] = useState<CellValue[][]>([]);

  async function onFile(
    e: React.ChangeEvent<HTMLInputElement>,
  ) {
    const file = e.target.files?.[0];
    if (!file) return;
    const { sheets } = await readXlsx(
      await file.arrayBuffer(),
    );
    setRows(sheets[0]?.rows ?? []);
  }

  return (
    <>
      <input
        type="file"
        accept=".xlsx,.xlsm"
        onChange={onFile}
      />
      <p>{rows.length} rows in the first sheet</p>
    </>
  );
}
```

### Node — a file on disk

```ts
import { readFile } from "node:fs/promises";
import { readXlsx } from "datareader-excel";

// readFile returns a Buffer (a Uint8Array), which
// readXlsx accepts directly — no conversion needed.
const { sheets } = await readXlsx(
  await readFile("data.xlsx"),
);

for (const sheet of sheets) {
  console.log(`${sheet.name}: ${sheet.rows.length} rows`);
}
```

### Dates & merged cells

```ts
import { readXlsx } from "datareader-excel";

const buf = await file.arrayBuffer();

// Default: date-formatted cells decode to a Date, and
// a merged value lives only in its top-left cell.
const { sheets } = await readXlsx(buf);

// dates: false keeps the raw Excel serial number;
// mergedCells: "fill" copies a merged value across
// every cell of its range.
const raw = await readXlsx(buf, {
  dates: false,
  mergedCells: "fill",
});
```

## API

### `readXlsx(input, opts?)`

```ts
function readXlsx(
  input: ArrayBuffer | Uint8Array,
  opts?: XlsxOptions,
): Promise<ParsedFile>;
```

Reads a whole `.xlsx`/`.xlsm` end-to-end — unzip → parse the workbook, shared
strings, and styles → decode every worksheet into a dense, row-major grid — and
resolves to a `ParsedFile`. **Async**: ZIP inflation uses `DecompressionStream`,
so the call returns a `Promise` you `await`. Throws an `XlsxError` for a
malformed, encrypted, or oversized file (see
[Security & limits](#security--limits)).

### `XlsxOptions`

```ts
type XlsxOptions = {
  // Decode date-formatted serial numbers to Date.
  // false keeps the raw serial number. Default true.
  dates?: boolean;
  // "topLeft": a merged value stays only in the
  // top-left cell (faithful). "fill": propagate it
  // across the whole merge range. Default "topLeft".
  mergedCells?: "topLeft" | "fill";
  // Reject a workbook whose grid (rows × columns,
  // summed across sheets) exceeds this many cells.
  // Default 5,000,000.
  maxCells?: number;
  // Reject a workbook whose cumulative ZIP-inflate
  // output exceeds this many bytes (zip-bomb guard).
  // Default 536,870,912 (512 MiB).
  maxInflatedBytes?: number;
};
```

### `ParsedFile` and `Sheet`

```ts
type ParsedFile = {
  format: "xlsx";
  // Every worksheet, in workbook order.
  sheets: Sheet[];
};

type Sheet = {
  // The worksheet's name (its tab label).
  name: string;
  // Row-major cells; short rows are padded with
  // null so every row is the sheet's full width.
  rows: CellValue[][];
};
```

### `CellValue`

```ts
// A parsed cell: a string, a finite number, a
// boolean, a Date (date-formatted cells), or null
// (a blank or error cell).
type CellValue =
  | string
  | number
  | boolean
  | Date
  | null;
```

### `XlsxError`

```ts
class XlsxError extends Error {}
```

Thrown for a malformed, encrypted, or oversized file (a bad input, not a reader
bug) — "not a valid .xlsx", "encrypted workbooks are unsupported", or a limit
overflow. Because it extends `Error`, a plain `catch` still catches it; check
`err instanceof XlsxError` to tell a rejected file apart from an unexpected bug.

### `DEFAULT_LIMITS`

```ts
const DEFAULT_LIMITS: {
  maxCells: number;
  maxInflatedBytes: number;
} = {
  maxCells: 5_000_000, // rows × columns ceiling
  maxInflatedBytes: 512 * 1024 * 1024, // 512 MiB
};
```

The default ceilings `readXlsx` uses when you pass no `maxCells` /
`maxInflatedBytes`. No real spreadsheet is affected; pass stricter values where
memory is tight.

## Format coverage

| Area         | Handled                                                                                     |
| ------------ | ------------------------------------------------------------------------------------------- |
| Files        | `.xlsx` and `.xlsm` (OOXML / SpreadsheetML); **not** legacy `.xls` (BIFF/OLE)                |
| Sheets       | every worksheet, in workbook order (`ParsedFile.sheets`)                                     |
| Strings      | shared and inline strings; all `<r><t>` rich-text runs concatenated to the full value       |
| Numbers      | full stored double precision — the underlying value, never the rounded display              |
| Booleans     | `TRUE`/`FALSE` cells → `boolean`                                                             |
| Dates        | date-formatted serials → `Date` (spec-correct 1900/1904 systems, UTC); `dates: false` keeps the serial |
| Formulas     | the cached `<v>` result (no formula evaluation)                                              |
| Merged cells | value in the top-left cell, or propagated with `mergedCells: "fill"`                         |
| Missing data | blank cells and error cells (`#N/A`, `#DIV/0!`, …) both → `null`                             |
| Encoding     | UTF-8 via `TextDecoder`; non-Latin text (Hebrew/CJK/accents) reads correctly                |
| Not covered  | writing, styling/fonts/colors, charts, images, comments, data validation, hyperlink targets, encrypted workbooks |

## Better than ExcelJS

Read-side ExcelJS bugs cluster tightly, and `datareader-excel` was built to
cover the cluster. Each row links a real (open, at time of writing) ExcelJS
issue to how this reader behaves instead.

| Theme | ExcelJS issue(s) | `datareader-excel` |
| --- | --- | --- |
| **Dates** | #486 (timezone drift), #1928 (`1/1/1900` wrong), #2966 (date-formatted formula → `Invalid Date`) | Serial→`Date` per the OOXML spec (1900 leap-year bug, 1904 epoch, UTC-consistent), validated to **ground truth**, not to ExcelJS. |
| **Opt-out date coercion** | #522 | `dates: false` returns the raw serial `number` instead of a `Date`. |
| **Number stored as text** | #475, #1162 | A text-typed cell stays a `string` (the file's truth) — never mis-coerced to a number. |
| **Rich-text read** | #2761 (leading run dropped) | Concatenate *all* `<r><t>` runs → the correct full string. |
| **Merged cells** | #2729, #2640, #1567 | Parse `<mergeCells>`; value in the top-left cell (faithful), or `mergedCells: "fill"` across the range. |
| **Cryptic errors on bad input** | pervasive | A clear `XlsxError` ("not a valid .xlsx", "encrypted workbooks are unsupported") instead of a ZIP stacktrace. |
| **Transitive CVEs** | #3024, #3041, #3053, #2866, #2715 | **Zero dependencies** — a test asserts the package has no `dependencies`. No transitive-CVE class exists. |
| **Read OOM on big files** | #2925, #412, #1842 | Read-only + byte-budgeted; the `maxCells`/`maxInflatedBytes` caps throw a **bounded, actionable** `XlsxError` (naming the option to raise) instead of a heap crash. |
| **Crash on files missing optional parts** | #1329 | Graceful defaults (no `workbookPr` → 1900 system; no `styles`/`sharedStrings` → sensible) — reads files ExcelJS throws on. |
| **Unicode cell values** | #1473, #995 (signal) | Shared/inline strings via `TextDecoder("utf-8")` — non-Latin text (Hebrew/CJK/accents) reads correctly. |

**Date oracle caveat.** ExcelJS cannot be the oracle for dates — it *is* the
buggy party. Structure (strings, numbers, booleans, layout, sheet order) is
validated by an ExcelJS structural differential; **dates are validated to
ground truth**: write a known date via ExcelJS, read the resulting serial, and
assert we recover exactly that date via the spec-correct formula, including the
boundary cases ExcelJS gets wrong.

### Researcher-critical guarantees

Beyond the bug-fixes, `datareader-excel` guarantees the value-level properties
statistical analysis depends on — the places a *naive* reader tends to fail —
as explicit tested invariants:

- **Raw, analyzable values — never the formatted display string.** A cell
  displayed `45%` returns `0.45`; a currency cell returns the number. Every
  numeric cell is computable.
- **Full double precision** — the underlying stored value, not the rounded
  display.
- **Uniform missing data** — blank cells *and* error cells (`#N/A`, `#DIV/0!`,
  …) both → `null`, so missingness is consistent downstream.
- **Wide files** — multi-letter column references (`AA`, `CV1`, `AMJ1`) parse
  correctly; hundreds of variables per sheet are fine.

## Security & limits

`datareader-excel` bounds the two attacker-controllable allocations — the
inflated ZIP output and the parsed cell grid. While unzipping it caps
cumulative inflate output at `maxInflatedBytes` (default **512 MiB**) so a zip
bomb aborts before it materializes; after parsing it rejects any workbook whose
cell count (rows × columns, summed across sheets) exceeds `maxCells` (default
**5,000,000**). Either overflow throws a catchable `XlsxError` whose message
names the option to raise. Encrypted (password-protected) workbooks are OLE
containers, not ZIPs — they are detected by signature and rejected, not parsed.

```ts
import { readXlsx, XlsxError } from "datareader-excel";

try {
  const { sheets } = await readXlsx(buf, {
    maxCells: 200_000,
    maxInflatedBytes: 32 * 1024 * 1024,
  });
} catch (err) {
  if (err instanceof XlsxError) {
    // malformed, encrypted, or too large — reject it
  }
}
```

Recommendations for consumers:

- Still bound the **input** byte size before you hand a buffer to `readXlsx` —
  reject files larger than you expect. The caps bound the inflated output and
  the parsed grid, not the raw file you accept.
- Treat every sheet name and cell string as **untrusted** — don't use a
  file-derived string as a plain-object key without care; prefer a `Map` or a
  `null`-prototype object to avoid prototype-pollution surprises from
  adversarial names.

## Roadmap

`datareader-excel` reads `.xlsx`/`.xlsm` data correctly today and is
intentionally **read-only** (no writer — that is the whole reason it stays small
and dependency-free). On the map for future releases:

- **Streaming** — a chunked API for workbooks too large to inflate whole,
  alongside today's whole-file read.
- **More encodings** — XML parts are decoded as UTF-8 today (the OOXML norm);
  tolerating declared alternate encodings via `TextDecoder` is a natural
  extension.

Issues and contributions are welcome.

## License

MIT © 2026 everydaydesign
