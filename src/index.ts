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
