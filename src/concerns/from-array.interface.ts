/**
 * Provide export data as a two-dimensional array.
 *
 * Each inner array is one row of cell values.
 */
export interface FromArray {
  array(): any[][] | Promise<any[][]>;
}
