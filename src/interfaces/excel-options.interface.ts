import type { CsvSettings } from "../concerns";

export interface ExcelModuleOptions {
  /** Default export type when none can be inferred (default `'xlsx'`). */
  defaultType?: "xlsx" | "csv";
  /** Directory used for temporary files (default: OS temp dir). */
  tempDirectory?: string;
  /** Global CSV defaults applied when no per-export settings exist. */
  csv?: CsvSettings;
}

export interface ExcelAsyncOptions {
  imports?: any[];
  inject?: any[];
  useFactory: (
    ...args: any[]
  ) => Promise<ExcelModuleOptions> | ExcelModuleOptions;
}
