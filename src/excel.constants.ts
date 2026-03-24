export const EXCEL_OPTIONS = "EXCEL_OPTIONS";

export enum ExcelType {
  XLSX = "xlsx",
  CSV = "csv",
}

export const CONTENT_TYPES: Record<ExcelType, string> = {
  [ExcelType.XLSX]:
    "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
  [ExcelType.CSV]: "text/csv",
};

export const EXTENSION_MAP: Record<string, ExcelType> = {
  xlsx: ExcelType.XLSX,
  csv: ExcelType.CSV,
};
