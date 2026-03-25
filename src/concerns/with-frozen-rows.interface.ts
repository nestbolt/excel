/**
 * Freeze row(s) at the top of the sheet so they stay visible when
 * scrolling.
 *
 * Return the number of rows to freeze from the top (e.g. `1` to
 * freeze only the heading row).
 */
export interface WithFrozenRows {
  frozenRows(): number;
}

/**
 * Freeze column(s) at the left of the sheet so they stay visible
 * when scrolling horizontally.
 *
 * Return the number of columns to freeze from the left.
 */
export interface WithFrozenColumns {
  frozenColumns(): number;
}
