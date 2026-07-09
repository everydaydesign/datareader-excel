/** Thrown for a malformed, hostile, oversized, or unsupported (encrypted) input — a bad file, not a
 * reader bug. Extends Error, so a plain catch still catches it; check `err instanceof XlsxError`. */
export class XlsxError extends Error {}

/** Resource ceilings that keep a hostile/huge .xlsx from exhausting memory. Generous defaults; pass a
 * stricter `readXlsx(input, opts)` where memory is tight (the message names the option to raise). */
export const DEFAULT_LIMITS: { maxCells: number; maxInflatedBytes: number } = {
  maxCells: 5_000_000,
  maxInflatedBytes: 512 * 1024 * 1024,
};
