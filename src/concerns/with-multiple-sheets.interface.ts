/**
 * Export multiple sheets in a single workbook.
 *
 * Each element returned by `sheets()` is an independent exportable
 * object that can implement its own concerns (FromCollection,
 * WithHeadings, etc.).
 */
export interface WithMultipleSheets {
  sheets(): object[];
}
