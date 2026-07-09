/** A parsed cell: a string, a finite number, a boolean, a Date, or null (blank/error). */
export type CellValue = string | number | boolean | Date | null;

/** Options for {@link readXlsx}. */
export type XlsxOptions = {
  /** Decode date-formatted serial numbers to Date. false → keep the raw serial number. Default true. */
  dates?: boolean;
  /** "topLeft": merged value only in the top-left cell (faithful). "fill": propagate across the
   *  merge range. Default "topLeft". */
  mergedCells?: "topLeft" | "fill";
  /** ncells (rows × cols, summed across sheets) ceiling → XlsxError. Default 5,000,000. */
  maxCells?: number;
  /** Cumulative ZIP-inflate output ceiling in bytes → XlsxError. Default 512 MiB. */
  maxInflatedBytes?: number;
};

/** One worksheet as a dense, row-major grid. */
export type Sheet = { name: string; rows: CellValue[][] };

/** The result of {@link readXlsx}: every worksheet, in workbook order. */
export type ParsedFile = { format: "xlsx"; sheets: Sheet[] };
