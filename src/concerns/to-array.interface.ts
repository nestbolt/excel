/**
 * Receive imported data as a two-dimensional array.
 */
export interface ToArray {
  handleArray(rows: any[][]): void | Promise<void>;
}
