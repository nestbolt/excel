/**
 * Skip rows before a given 1-based row number during import.
 */
export interface WithStartRow {
  startRow(): number;
}
