import "reflect-metadata";
import {
  EXPORTABLE_META,
  EXPORT_COLUMNS_META,
  EXPORT_IGNORE_META,
} from "./constants";
import type { ExportableOptions, ExportColumnOptions } from "./interfaces";
import { numberToColumnLetter } from "../helpers";

/* ------------------------------------------------------------------ */
/*  Internal types                                                     */
/* ------------------------------------------------------------------ */

interface ResolvedColumn {
  propertyKey: string;
  order: number;
  header: string;
  options: ExportColumnOptions;
}

/* ------------------------------------------------------------------ */
/*  Helpers                                                            */
/* ------------------------------------------------------------------ */

function toTitleCase(str: string): string {
  return str
    .replace(/([a-z])([A-Z])/g, "$1 $2")
    .replace(/[_-]/g, " ")
    .replace(/\b\w/g, (c) => c.toUpperCase());
}

function collectColumns(entityClass: Function): ResolvedColumn[] {
  const chain: Function[] = [];
  let current: Function | null = entityClass;
  while (
    current &&
    current !== Function.prototype &&
    current !== Object
  ) {
    chain.unshift(current);
    current = Object.getPrototypeOf(current);
  }

  const merged = new Map<string, ExportColumnOptions>();

  for (const ctor of chain) {
    const cols: Map<string, ExportColumnOptions> | undefined =
      Reflect.getOwnMetadata(EXPORT_COLUMNS_META, ctor);
    if (cols) {
      for (const [key, opts] of cols) {
        merged.set(key, opts);
      }
    }

    const ign: Set<string> | undefined =
      Reflect.getOwnMetadata(EXPORT_IGNORE_META, ctor);
    if (ign) {
      for (const key of ign) {
        merged.delete(key);
      }
    }
  }

  let insertionIndex = 0;
  const resolved: ResolvedColumn[] = [];
  for (const [propertyKey, options] of merged) {
    resolved.push({
      propertyKey,
      order: options.order ?? 1_000_000 + insertionIndex,
      header: options.header ?? toTitleCase(propertyKey),
      options,
    });
    insertionIndex++;
  }

  resolved.sort((a, b) => a.order - b.order);
  return resolved;
}

function buildColumnWidths(
  opts: ExportableOptions,
  columns: ResolvedColumn[],
): Record<string, number> | null {
  const result: Record<string, number> = {};

  if (opts.columnWidths) {
    Object.assign(result, opts.columnWidths);
  }

  columns.forEach((col, idx) => {
    if (col.options.width !== undefined) {
      result[numberToColumnLetter(idx + 1)] = col.options.width;
    }
  });

  return Object.keys(result).length > 0 ? result : null;
}

function buildColumnFormats(
  columns: ResolvedColumn[],
): Record<string, string> {
  const result: Record<string, string> = {};
  columns.forEach((col, idx) => {
    if (col.options.format) {
      result[numberToColumnLetter(idx + 1)] = col.options.format;
    }
  });
  return result;
}

/* ------------------------------------------------------------------ */
/*  Public API                                                         */
/* ------------------------------------------------------------------ */

/**
 * Read decorator metadata from `entityClass` and return a plain object
 * that implements the appropriate export concerns.
 *
 * The returned object can be passed directly to any `ExcelService`
 * export method (`download`, `raw`, etc.).
 */
export function buildExportFromEntity<T>(
  entityClass: new (...args: any[]) => T,
  data: T[],
): object {
  const exportableOpts: ExportableOptions | undefined =
    Reflect.getOwnMetadata(EXPORTABLE_META, entityClass);

  if (!exportableOpts) {
    throw new Error(
      `Class "${entityClass.name}" is not decorated with @Exportable().`,
    );
  }

  const columns = collectColumns(entityClass);

  if (columns.length === 0) {
    throw new Error(
      `Class "${entityClass.name}" has no @ExportColumn() properties.`,
    );
  }

  // --- build the export object --------------------------------------
  const exportObj: Record<string, any> = {};

  // FromCollection
  exportObj.collection = () => data;

  // WithHeadings
  exportObj.headings = () => columns.map((col) => col.header);

  // WithMapping
  exportObj.map = (row: any) =>
    columns.map((col) => {
      const raw = row[col.propertyKey];
      return col.options.map ? col.options.map(raw, row) : raw;
    });

  // WithTitle
  if (exportableOpts.title) {
    const title = exportableOpts.title;
    exportObj.title = () => title;
  }

  // WithColumnWidths
  const widths = buildColumnWidths(exportableOpts, columns);
  if (widths) {
    exportObj.columnWidths = () => widths;
  }

  // WithColumnFormatting
  const formats = buildColumnFormats(columns);
  if (Object.keys(formats).length > 0) {
    exportObj.columnFormats = () => formats;
  }

  // WithAutoFilter
  if (exportableOpts.autoFilter) {
    const af = exportableOpts.autoFilter;
    exportObj.autoFilter = () => af;
  }

  // ShouldAutoSize
  if (exportableOpts.autoSize) {
    exportObj.shouldAutoSize = true;
  }

  // WithFrozenRows
  if (exportableOpts.frozenRows && exportableOpts.frozenRows > 0) {
    const fr = exportableOpts.frozenRows;
    exportObj.frozenRows = () => fr;
  }

  // WithFrozenColumns
  if (exportableOpts.frozenColumns && exportableOpts.frozenColumns > 0) {
    const fc = exportableOpts.frozenColumns;
    exportObj.frozenColumns = () => fc;
  }

  return exportObj;
}
