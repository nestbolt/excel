/**
 * Map column letters or 1-based indices to named fields.
 *
 * Example: `{ name: 'A', email: 'C' }` or `{ name: 1, email: 3 }`.
 */
export interface WithColumnMapping {
  columnMapping(): Record<string, string | number>;
}
