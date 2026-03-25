import type { ExcelModuleOptions } from "../interfaces";
import type { CsvSettings, WithCsvSettings } from "../concerns";

function isWithCsvSettings(obj: any): obj is WithCsvSettings {
  return typeof obj.csvSettings === "function";
}

export function resolveCsvSettings(
  exportable: object,
  options: ExcelModuleOptions,
): Required<CsvSettings> {
  const defaults: Required<CsvSettings> = {
    delimiter: ",",
    quoteChar: '"',
    lineEnding: "\n",
    useBom: false,
    encoding: "utf-8",
  };

  const global = options.csv ?? {};
  const perExport = isWithCsvSettings(exportable)
    ? exportable.csvSettings()
    : {};

  return { ...defaults, ...global, ...perExport };
}
