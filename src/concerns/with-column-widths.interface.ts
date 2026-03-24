/**
 * Set explicit column widths.
 *
 * Keys are column letters (`'A'`, `'B'`, ...) and values are
 * widths in character units.
 */
export interface WithColumnWidths {
  columnWidths(): Record<string, number>;
}
