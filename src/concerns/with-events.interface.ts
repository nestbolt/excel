import type { Workbook, Worksheet } from "exceljs";

/**
 * Register lifecycle event listeners for the export process.
 */
export interface WithEvents {
  registerEvents(): Partial<Record<ExcelExportEvent, (event: any) => void>>;
}

export enum ExcelExportEvent {
  BEFORE_EXPORT = "beforeExport",
  BEFORE_WRITING = "beforeWriting",
  BEFORE_SHEET = "beforeSheet",
  AFTER_SHEET = "afterSheet",
}

export interface BeforeExportEventPayload {
  exportable: object;
  workbook: Workbook;
}

export interface BeforeWritingEventPayload {
  exportable: object;
  workbook: Workbook;
}

export interface BeforeSheetEventPayload {
  exportable: object;
  worksheet: Worksheet;
}

export interface AfterSheetEventPayload {
  exportable: object;
  worksheet: Worksheet;
}
