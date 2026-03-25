/**
 * Use a row in the spreadsheet as column headings to derive object keys.
 *
 * Set `hasHeadingRow` to `true` as a marker. Optionally implement
 * `headingRow()` to specify a custom row number (defaults to 1).
 */
export interface WithHeadingRow {
  readonly hasHeadingRow: true;
  headingRow?(): number;
}
