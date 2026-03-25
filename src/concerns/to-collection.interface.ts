/**
 * Receive imported data as an array of objects.
 *
 * Requires {@link WithHeadingRow} or {@link WithColumnMapping} to derive
 * object keys.
 */
export interface ToCollection<T = Record<string, any>> {
  handleCollection(rows: T[]): void | Promise<void>;
}
