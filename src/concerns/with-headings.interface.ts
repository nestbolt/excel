/**
 * Add a heading row to the sheet.
 *
 * Return a single array for one heading row, or a nested array for
 * multiple heading rows.
 */
export interface WithHeadings {
  headings(): string[] | string[][];
}
