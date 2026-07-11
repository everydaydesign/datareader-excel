# Changelog

All notable changes to `datareader-excel` are documented here.
The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.1.0/), and this
project adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

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

[0.1.2]: https://github.com/everydaydesign/datareader-excel/releases/tag/v0.1.2
[0.1.1]: https://github.com/everydaydesign/datareader-excel/releases/tag/v0.1.1
[0.1.0]: https://github.com/everydaydesign/datareader-excel/releases/tag/v0.1.0
