/**
 * Add Excel auto-filter dropdowns to columns.
 *
 * Return a cell range like `'A1:D1'`, or `'auto'` to auto-detect
 * the range from headings.
 */
export interface WithAutoFilter {
  autoFilter(): string;
}
