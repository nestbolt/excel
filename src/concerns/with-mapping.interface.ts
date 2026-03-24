/**
 * Transform each row before it is written to the sheet.
 *
 * The `map` method receives one raw item from `collection()` or
 * `array()` and must return a flat array of cell values.
 */
export interface WithMapping<T = any> {
  map(row: T): any[];
}
