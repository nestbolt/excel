/**
 * Start writing data at a specific cell instead of A1.
 *
 * Return a cell reference like `'C3'`.
 */
export interface WithCustomStartCell {
  startCell(): string;
}
