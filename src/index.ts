// Module
export { ExcelModule } from "./excel.module";

// Service
export { ExcelService } from "./excel.service";

// Constants & enums
export { EXCEL_OPTIONS, ExcelType } from "./excel.constants";

// Interfaces
export type { ExcelModuleOptions, ExcelAsyncOptions } from "./interfaces";
export type { ExcelDownloadResult } from "./interfaces";

// Concerns — data sources
export type { FromCollection } from "./concerns";
export type { FromArray } from "./concerns";

// Concerns — headings & mapping
export type { WithHeadings } from "./concerns";
export type { WithMapping } from "./concerns";

// Concerns — sheet structure
export type { WithTitle } from "./concerns";
export type { WithMultipleSheets } from "./concerns";
export type { WithCustomStartCell } from "./concerns";

// Concerns — formatting & styling
export type { WithColumnWidths } from "./concerns";
export type { WithColumnFormatting } from "./concerns";
export type {
  WithStyles,
  CellStyle,
  FontStyle,
  AlignmentStyle,
  FillStyle,
  BorderStyles,
  BorderStyle,
} from "./concerns";
export type { ShouldAutoSize } from "./concerns";

// Concerns — auto-filter & freeze panes
export type { WithAutoFilter } from "./concerns";
export type { WithFrozenRows, WithFrozenColumns } from "./concerns";

// Concerns — template
export type { FromTemplate, WithTemplateData } from "./concerns";

// Concerns — properties
export type { WithProperties, ExcelProperties } from "./concerns";

// Concerns — CSV
export type { WithCsvSettings, CsvSettings } from "./concerns";

// Concerns — events
export type {
  WithEvents,
  BeforeExportEventPayload,
  BeforeWritingEventPayload,
  BeforeSheetEventPayload,
  AfterSheetEventPayload,
} from "./concerns";
export { ExcelExportEvent } from "./concerns";

// Concerns — import data receivers
export type { ToArray } from "./concerns";
export type { ToCollection } from "./concerns";

// Concerns — import row processing
export type { WithHeadingRow } from "./concerns";
export type { WithImportMapping } from "./concerns";
export type { WithColumnMapping } from "./concerns";

// Concerns — import validation
export type {
  WithValidation,
  ValidationRules,
  ValidationRule,
} from "./concerns";
export type { SkipsOnError } from "./concerns";
export type { SkipsEmptyRows } from "./concerns";

// Concerns — import limits & batching
export type { WithLimit } from "./concerns";
export type { WithStartRow } from "./concerns";
export type { WithBatchInserts } from "./concerns";

// Import result types
export type {
  ImportResult,
  ImportValidationError,
  FieldError,
} from "./interfaces";

// Decorators — entity-based export
export { Exportable } from "./decorators";
export { ExportColumn } from "./decorators";
export { ExportIgnore } from "./decorators";
export { buildExportFromEntity } from "./decorators";
export type { ExportableOptions, ExportColumnOptions } from "./decorators";
