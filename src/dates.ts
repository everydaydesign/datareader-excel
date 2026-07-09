const DAY_MS = 86_400_000;

/** Convert an Excel date serial to a UTC Date.
 *
 * 1900 system: day 1 = 1900-01-01, and Excel wrongly counts a fictitious 1900-02-29 (serial 60), so
 * the phantom day is dropped only for serials at/after 61 to keep serial 1 (= 1900-01-01) correct —
 * the bug ExcelJS #1928/#486 get wrong. 1904 system: serial 0 = 1904-01-01. Fractional part is the
 * time of day. Always UTC (no local-timezone drift). */
export function serialToDate(serial: number, date1904: boolean): Date {
  if (date1904) return new Date(Date.UTC(1904, 0, 1) + serial * DAY_MS);
  const days = serial >= 61 ? serial - 1 : serial;
  return new Date(Date.UTC(1899, 11, 31) + days * DAY_MS);
}
