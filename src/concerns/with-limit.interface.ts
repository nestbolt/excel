/**
 * Limit the number of data rows read during import.
 */
export interface WithLimit {
  limit(): number;
}
