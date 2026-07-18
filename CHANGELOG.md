# Changelog

All notable changes to `datareader-excel` are documented here.
The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.1.0/), and this
project adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

## [0.2.0] — 2026-07-18

Fixes a set of silent-data-loss classes on spec-legal input. One **behavior change** (M2, below)
shifts parsing of files that omit cell refs; pin `0.1.x` if you depend on the old behavior.

### Fixed
- **A `<!DOCTYPE>` in any XML part no longer zeroes that part.** It was parsed as an element and became
  the document root, silently emptying that part's strings/styles/cells; `<!…>` declarations are now
  skipped (internal-subset `[ … ]` aware). No entity expansion (XXE-safe).
- **Cells that omit `r` now follow the previous cell (behavior change).** ECMA-376: an omitted ref
  means the next column. The reader defaulted every ref-less cell to column 0, so three ref-less cells
  collapsed to one; now sequential. Files where every cell has an `r` (Excel always writes it) are
  unaffected.
- **More date formats detected.** Added the East-Asian builtin date `numFmtId`s (27–36, 50–58, matching
  SheetJS/ExcelJS) and the system `[$-F800]` / `[$-F400]` long-date/time codes, previously mis-typed as
  raw serial numbers.
- **Malformed input fails loud.** A cell ref that doesn't parse to a valid column (lowercase `a1`) and a
  worksheet whose part is absent now throw an `XlsxError` naming the ref / sheet, instead of silently
  dropping the value / whole sheet.

### Changed
- **`DEFAULT_LIMITS` doc corrected.** Neither ceiling bounds peak heap directly: `maxInflatedBytes` caps
  inflated bytes, but each part parses into a node tree ~13× its text (built before `maxCells`), so a
  512 MiB inflate ceiling can still drive several GB. Lower it where memory is tight.

## [0.1.2] — 2026-07-11

### Fixed
- Defensive guards for malformed input: merge ranges with an incomplete reference are skipped,
  and sparse/empty cells are handled instead of dereferenced.

### Changed
- Internal: non-null assertions for structurally-bounded indexing (`noUncheckedIndexedAccess`-safe).
  Pure type-safety — no API or parsed-output change.

## [0.1.1] — 2026-07-09

### Changed
- Docs: describe capabilities on their own terms.

## [0.1.0] — 2026-07-09

### Added
- Initial release: zero-dependency Excel `.xlsx`/`.xlsm` reader for the browser and Node —
  hand-written ZIP + minimal OOXML walker, correct dates (1900 leap-bug + 1904), merged cells,
  bounded memory.

[0.2.0]: https://github.com/everydaydesign/datareader-excel/releases/tag/v0.2.0
[0.1.2]: https://github.com/everydaydesign/datareader-excel/releases/tag/v0.1.2
[0.1.1]: https://github.com/everydaydesign/datareader-excel/releases/tag/v0.1.1
[0.1.0]: https://github.com/everydaydesign/datareader-excel/releases/tag/v0.1.0
