// Data sources
export { FromCollection } from "./from-collection.interface";
export { FromArray } from "./from-array.interface";

// Headings & mapping
export { WithHeadings } from "./with-headings.interface";
export { WithMapping } from "./with-mapping.interface";

// Sheet structure
export { WithTitle } from "./with-title.interface";
export { WithMultipleSheets } from "./with-multiple-sheets.interface";
export { WithCustomStartCell } from "./with-custom-start-cell.interface";

// Formatting & styling
export { WithColumnWidths } from "./with-column-widths.interface";
export { WithColumnFormatting } from "./with-column-formatting.interface";
export {
  WithStyles,
  CellStyle,
  FontStyle,
  AlignmentStyle,
  FillStyle,
  BorderStyles,
  BorderStyle,
} from "./with-styles.interface";
export { ShouldAutoSize } from "./should-auto-size.interface";

// Auto-filter & freeze panes
export { WithAutoFilter } from "./with-auto-filter.interface";
export {
  WithFrozenRows,
  WithFrozenColumns,
} from "./with-frozen-rows.interface";

// Template
export { FromTemplate, WithTemplateData } from "./from-template.interface";

// Document properties
export { WithProperties, ExcelProperties } from "./with-properties.interface";

// CSV
export { WithCsvSettings, CsvSettings } from "./with-csv-settings.interface";

// Events
export {
  WithEvents,
  ExcelExportEvent,
  BeforeExportEventPayload,
  BeforeWritingEventPayload,
  BeforeSheetEventPayload,
  AfterSheetEventPayload,
} from "./with-events.interface";

// Import — data receivers
export { ToArray } from "./to-array.interface";
export { ToCollection } from "./to-collection.interface";

// Import — row processing
export { WithHeadingRow } from "./with-heading-row.interface";
export { WithImportMapping } from "./with-import-mapping.interface";
export { WithColumnMapping } from "./with-column-mapping.interface";

// Import — validation
export {
  WithValidation,
  ValidationRules,
  ValidationRule,
} from "./with-validation.interface";
export { SkipsOnError } from "./skips-on-error.interface";
export { SkipsEmptyRows } from "./skips-empty-rows.interface";

// Import — limits & batching
export { WithLimit } from "./with-limit.interface";
export { WithStartRow } from "./with-start-row.interface";
export { WithBatchInserts } from "./with-batch-inserts.interface";
