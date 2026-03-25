import type { CellValue, Worksheet } from "exceljs";
import type {
  SkipsEmptyRows,
  SkipsOnError,
  ToArray,
  ToCollection,
  WithBatchInserts,
  WithColumnMapping,
  WithHeadingRow,
  WithImportMapping,
  WithLimit,
  WithStartRow,
  WithValidation,
} from "./concerns";
import { columnLetterToNumber } from "./helpers";
import { validateRow } from "./helpers/validate-row";
import type { ImportResult, ImportValidationError } from "./interfaces";

/* ------------------------------------------------------------------ */
/*  Type guards                                                        */
/* ------------------------------------------------------------------ */

function isToArray(obj: any): obj is ToArray {
  return typeof obj.handleArray === "function";
}

function isToCollection(obj: any): obj is ToCollection {
  return typeof obj.handleCollection === "function";
}

function isWithHeadingRow(obj: any): obj is WithHeadingRow {
  return obj.hasHeadingRow === true;
}

function isWithImportMapping(obj: any): obj is WithImportMapping {
  return typeof obj.mapRow === "function";
}

function isWithColumnMapping(obj: any): obj is WithColumnMapping {
  return typeof obj.columnMapping === "function";
}

function isWithValidation(obj: any): obj is WithValidation {
  return typeof obj.rules === "function";
}

function isWithBatchInserts(obj: any): obj is WithBatchInserts {
  return (
    typeof obj.batchSize === "function" && typeof obj.handleBatch === "function"
  );
}

function isWithLimit(obj: any): obj is WithLimit {
  return typeof obj.limit === "function";
}

function isWithStartRow(obj: any): obj is WithStartRow {
  return typeof obj.startRow === "function";
}

function isSkipsOnError(obj: any): obj is SkipsOnError {
  return obj.skipsOnError === true;
}

function isSkipsEmptyRows(obj: any): obj is SkipsEmptyRows {
  return obj.skipsEmptyRows === true;
}

/* ------------------------------------------------------------------ */
/*  Helpers                                                            */
/* ------------------------------------------------------------------ */

function extractCellValue(value: CellValue): any {
  if (value === null || value === undefined) return null;
  if (typeof value === "object" && "result" in value) {
    return (value as any).result;
  }
  if (typeof value === "object" && "richText" in value) {
    return (value as any).richText.map((rt: any) => rt.text).join("");
  }
  return value;
}

function isEmptyRow(row: any[]): boolean {
  return row.every((v) => v === null || v === undefined || v === "");
}

/* ------------------------------------------------------------------ */
/*  Sheet reader                                                       */
/* ------------------------------------------------------------------ */

export async function processSheet(
  worksheet: Worksheet,
  importable: object,
): Promise<ImportResult> {
  // --- extract all raw rows from worksheet --------------------------
  const rawRows: { rowNumber: number; values: any[] }[] = [];

  worksheet.eachRow((row, rowNumber) => {
    const values: any[] = [];
    row.eachCell({ includeEmpty: true }, (cell, colNumber) => {
      // Expand values array to accommodate the column
      while (values.length < colNumber) values.push(null);
      values[colNumber - 1] = extractCellValue(cell.value);
    });
    rawRows.push({ rowNumber, values });
  });

  // --- determine heading row ----------------------------------------
  let headings: string[] | null = null;
  let headingRowNumber = 0;

  if (isWithHeadingRow(importable)) {
    headingRowNumber =
      typeof importable.headingRow === "function" ? importable.headingRow() : 1;

    const headingEntry = rawRows.find((r) => r.rowNumber === headingRowNumber);
    if (headingEntry) {
      headings = headingEntry.values.map((v) =>
        v !== null && v !== undefined ? String(v) : "",
      );
    }
  }

  // --- determine data start row -------------------------------------
  let dataStartRow: number;
  if (isWithStartRow(importable)) {
    dataStartRow = importable.startRow();
  } else if (headingRowNumber > 0) {
    dataStartRow = headingRowNumber + 1;
  } else {
    dataStartRow = 1;
  }

  // --- filter to data rows ------------------------------------------
  let dataRows = rawRows.filter((r) => r.rowNumber >= dataStartRow);

  // --- skip empty rows ----------------------------------------------
  if (isSkipsEmptyRows(importable)) {
    dataRows = dataRows.filter((r) => !isEmptyRow(r.values));
  }

  // --- apply limit --------------------------------------------------
  if (isWithLimit(importable)) {
    dataRows = dataRows.slice(0, importable.limit());
  }

  // --- convert to objects if headings or column mapping present ------
  const useColumnMapping = isWithColumnMapping(importable);
  const useObjects = headings !== null || useColumnMapping;

  let processedRows: any[];
  let objectKeyOrder: string[] | null = null;

  if (useObjects) {
    let columnMap: Record<string, number> | null = null;

    if (useColumnMapping) {
      const raw = (importable as WithColumnMapping).columnMapping();
      columnMap = {};
      const entries: [string, number][] = [];
      for (const [fieldName, colRef] of Object.entries(raw)) {
        const idx =
          typeof colRef === "number"
            ? colRef - 1
            : columnLetterToNumber(colRef) - 1;
        columnMap[fieldName] = idx;
        entries.push([fieldName, idx]);
      }
      // Deterministic key order: sorted by column index
      objectKeyOrder = entries
        .sort((a, b) => a[1] - b[1])
        .map(([k]) => k);
    }

    if (!objectKeyOrder && headings) {
      objectKeyOrder = headings.map((h, i) => h || `__col${i}`);
    }

    processedRows = dataRows.map((r) => {
      const obj: Record<string, any> = {};
      if (columnMap) {
        for (const [field, idx] of Object.entries(columnMap)) {
          obj[field] = idx < r.values.length ? r.values[idx] : null;
        }
      } else if (headings) {
        for (let i = 0; i < headings.length; i++) {
          const key = headings[i] || `__col${i}`;
          obj[key] = i < r.values.length ? r.values[i] : null;
        }
      }
      return { rowNumber: r.rowNumber, data: obj };
    });
  } else {
    processedRows = dataRows.map((r) => ({
      rowNumber: r.rowNumber,
      data: r.values,
    }));
  }

  // --- apply import mapping -----------------------------------------
  if (isWithImportMapping(importable)) {
    processedRows = processedRows.map((r) => ({
      rowNumber: r.rowNumber,
      data: (importable as WithImportMapping).mapRow(r.data),
    }));
  }

  // --- validation ---------------------------------------------------
  const validationErrors: ImportValidationError[] = [];
  let skipped = 0;
  const validRows: any[] = [];

  if (isWithValidation(importable)) {
    const rules = importable.rules();
    const skipOnError = isSkipsOnError(importable);

    for (const row of processedRows) {
      const error = await validateRow(row.data, rules, row.rowNumber);
      if (error) {
        validationErrors.push(error);
        if (skipOnError) {
          skipped++;
          continue;
        }
      }
      validRows.push(row.data);
    }

    if (!skipOnError && validationErrors.length > 0) {
      const err = new Error(
        `Import validation failed with ${validationErrors.length} error(s).`,
      );
      (err as any).validationErrors = validationErrors;
      throw err;
    }
  } else {
    for (const row of processedRows) {
      validRows.push(row.data);
    }
  }

  // --- deliver to importable ----------------------------------------
  if (isWithBatchInserts(importable)) {
    const size = importable.batchSize();
    if (!Number.isInteger(size) || size < 1) {
      throw new Error(
        `WithBatchInserts.batchSize() must return a positive integer, got ${size}.`,
      );
    }
    for (let i = 0; i < validRows.length; i += size) {
      const batch = validRows.slice(i, i + size);
      await importable.handleBatch(batch);
    }
  }

  if (isToCollection(importable)) {
    await importable.handleCollection(validRows);
  }

  if (isToArray(importable)) {
    // If rows are objects, convert back to arrays for ToArray using
    // deterministic key order so column ordering matches the spreadsheet.
    const arrayRows =
      useObjects && objectKeyOrder
        ? validRows.map((obj) => objectKeyOrder!.map((k) => obj[k]))
        : validRows;
    await importable.handleArray(arrayRows);
  }

  return {
    rows: validRows,
    errors: validationErrors,
    skipped,
  };
}
