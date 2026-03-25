import { Workbook } from "exceljs";
import { Readable } from "stream";
import type { ExcelModuleOptions } from "./interfaces";
import type { ImportResult } from "./interfaces";
import { ExcelType } from "./excel.constants";
import { resolveCsvSettings } from "./helpers/csv-settings";
import { processSheet } from "./excel.sheet-reader";

export async function readImport(
  importable: object,
  source: string | Buffer,
  type: ExcelType,
  options: ExcelModuleOptions,
): Promise<ImportResult> {
  const workbook = new Workbook();

  if (Buffer.isBuffer(source)) {
    if (type === ExcelType.CSV) {
      const csvOpts = resolveCsvSettings(importable, options);
      await workbook.csv.read(Readable.from(source), {
        parserOptions: {
          delimiter: csvOpts.delimiter,
          quote: csvOpts.quoteChar,
        },
      });
    } else {
      await workbook.xlsx.load(source as any);
    }
  } else {
    if (type === ExcelType.CSV) {
      const csvOpts = resolveCsvSettings(importable, options);
      await workbook.csv.readFile(source, {
        parserOptions: {
          delimiter: csvOpts.delimiter,
          quote: csvOpts.quoteChar,
        },
      });
    } else {
      await workbook.xlsx.readFile(source);
    }
  }

  const worksheet = workbook.worksheets[0];
  if (!worksheet) {
    throw new Error("No worksheet found in the imported file.");
  }

  return processSheet(worksheet, importable);
}
