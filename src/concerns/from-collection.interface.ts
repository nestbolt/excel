/**
 * Provide export data as an array of objects or arrays.
 *
 * Each element represents one row. When combined with WithMapping,
 * the raw row is passed to `map()` before writing.
 */
export interface FromCollection<T = any> {
  collection(): T[] | Promise<T[]>;
}
