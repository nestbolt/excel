/**
 * Apply number formats to columns.
 *
 * Keys are column letters (`'A'`, `'B'`, ...) and values are
 * Excel number-format strings (e.g. `'#,##0.00'`, `'yyyy-mm-dd'`).
 */
export interface WithColumnFormatting {
  columnFormats(): Record<string, string>;
}
