import { Workbook } from "exceljs";
import type { ExcelModuleOptions } from "./interfaces";
import { ExcelType } from "./excel.constants";
import type {
  WithMultipleSheets,
  WithProperties,
  WithEvents,
  WithCsvSettings,
  CsvSettings,
} from "./concerns";
import { ExcelExportEvent } from "./concerns";
import { populateSheet } from "./excel.sheet";

/* ------------------------------------------------------------------ */
/*  Type guards                                                        */
/* ------------------------------------------------------------------ */

function isWithMultipleSheets(obj: any): obj is WithMultipleSheets {
  return typeof obj.sheets === "function";
}

function isWithProperties(obj: any): obj is WithProperties {
  return typeof obj.properties === "function";
}

function isWithEvents(obj: any): obj is WithEvents {
  return typeof obj.registerEvents === "function";
}

function isWithCsvSettings(obj: any): obj is WithCsvSettings {
  return typeof obj.csvSettings === "function";
}

/* ------------------------------------------------------------------ */
/*  Event helper                                                       */
/* ------------------------------------------------------------------ */

function fireEvent(
  exportable: object,
  event: ExcelExportEvent,
  payload: any,
): void {
  if (!isWithEvents(exportable)) return;
  const handlers = exportable.registerEvents();
  const handler = handlers[event];
  if (handler) handler(payload);
}

/* ------------------------------------------------------------------ */
/*  Writer                                                             */
/* ------------------------------------------------------------------ */

export async function writeExport(
  exportable: object,
  type: ExcelType,
  options: ExcelModuleOptions,
): Promise<Buffer> {
  const workbook = new Workbook();

  // --- document properties ------------------------------------------
  if (isWithProperties(exportable)) {
    const props = exportable.properties();
    if (props.creator) workbook.creator = props.creator;
    if (props.lastModifiedBy) workbook.lastModifiedBy = props.lastModifiedBy;
    if (props.title) workbook.title = props.title;
    if (props.subject) workbook.subject = props.subject;
    if (props.description) workbook.description = props.description;
    if (props.keywords) workbook.keywords = props.keywords;
    if (props.category) workbook.category = props.category;
    if (props.company) workbook.company = props.company;
    if (props.manager) workbook.manager = props.manager;
  }

  // --- fire beforeExport --------------------------------------------
  fireEvent(exportable, ExcelExportEvent.BEFORE_EXPORT, {
    exportable,
    workbook,
  });

  // --- sheets -------------------------------------------------------
  if (isWithMultipleSheets(exportable)) {
    const sheets = exportable.sheets();
    if (sheets.length === 0) {
      throw new Error("WithMultipleSheets.sheets() returned an empty array.");
    }
    for (const sheetExport of sheets) {
      const worksheet = workbook.addWorksheet();
      fireEvent(exportable, ExcelExportEvent.BEFORE_SHEET, {
        exportable: sheetExport,
        worksheet,
      });
      await populateSheet(worksheet, sheetExport);
      fireEvent(exportable, ExcelExportEvent.AFTER_SHEET, {
        exportable: sheetExport,
        worksheet,
      });
    }
  } else {
    const worksheet = workbook.addWorksheet();
    fireEvent(exportable, ExcelExportEvent.BEFORE_SHEET, {
      exportable,
      worksheet,
    });
    await populateSheet(worksheet, exportable);
    fireEvent(exportable, ExcelExportEvent.AFTER_SHEET, {
      exportable,
      worksheet,
    });
  }

  // --- fire beforeWriting -------------------------------------------
  fireEvent(exportable, ExcelExportEvent.BEFORE_WRITING, {
    exportable,
    workbook,
  });

  // --- serialize ----------------------------------------------------
  if (type === ExcelType.CSV) {
    const csvOpts = resolveCsvSettings(exportable, options);
    const arrayBuffer = await workbook.csv.writeBuffer({
      formatterOptions: {
        delimiter: csvOpts.delimiter,
        quote: csvOpts.quoteChar,
        rowDelimiter: csvOpts.lineEnding,
      },
    });
    let buffer = Buffer.from(arrayBuffer);
    if (csvOpts.useBom) {
      const bom = Buffer.from([0xef, 0xbb, 0xbf]);
      buffer = Buffer.concat([bom, buffer]);
    }
    return buffer;
  }

  // Default: XLSX
  const arrayBuffer = await workbook.xlsx.writeBuffer();
  return Buffer.from(arrayBuffer);
}

/* ------------------------------------------------------------------ */
/*  CSV settings resolution                                            */
/* ------------------------------------------------------------------ */

function resolveCsvSettings(
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
