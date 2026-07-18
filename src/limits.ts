/** Thrown for a malformed, hostile, oversized, or unsupported (encrypted) input — a bad file, not a
 * reader bug. Extends Error, so a plain catch still catches it; check `err instanceof XlsxError`. */
export class XlsxError extends Error {}

/** Resource ceilings for a hostile/huge .xlsx. IMPORTANT — neither bounds peak HEAP directly:
 * - `maxInflatedBytes` caps the cumulative INFLATED part bytes (the ZIP-bomb guard). Each XML part is
 *   then parsed into a node tree that measures ~13× its text, and that tree is built BEFORE the cell
 *   budget — so a 512 MiB inflate ceiling can still drive several GB of heap. Lower it where memory is
 *   tight (a streaming parser would make the tree the real bound; not built yet).
 * - `maxCells` caps the materialized grid AFTER parsing — it protects downstream consumers, not the
 *   parse itself.
 * Pass a stricter `readXlsx(input, opts)` where memory is tight; the throw names the option to raise. */
export const DEFAULT_LIMITS: { maxCells: number; maxInflatedBytes: number } = {
  maxCells: 5_000_000,
  maxInflatedBytes: 512 * 1024 * 1024,
};
