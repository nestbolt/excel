# Changelog

All notable changes to `@nestbolt/excel` will be documented in this file.

## v0.2.0 — Import XLSX/CSV with Validation

### Features

- **ExcelService** — New import methods: `import()`, `importFromBuffer()`, `toArray()`, `toCollection()`
- **ImportResult** — Structured return type with `rows`, `errors`, and `skipped` count
- **ToArray** — Receive imported data as a two-dimensional array
- **ToCollection** — Receive imported data as an array of objects
- **WithHeadingRow** — Use a spreadsheet row as column headings to derive object keys
- **WithImportMapping** — Transform each row after reading
- **WithColumnMapping** — Map column letters or indices to named fields
- **WithValidation** — Validate rows using custom rules or class-validator DTOs
- **WithBatchInserts** — Insert imported rows in configurable batch sizes
- **WithStartRow** — Skip rows before a given row number
- **WithLimit** — Limit the number of data rows read
- **SkipsEmptyRows** — Ignore blank rows during import
- **SkipsOnError** — Skip invalid rows instead of throwing

### Export Enhancements

- **WithAutoFilter** — Add auto-filter dropdowns to heading rows; supports `'auto'` detection or explicit range
- **WithFrozenRows** — Freeze a specified number of rows so they remain visible when scrolling
- **WithFrozenColumns** — Freeze a specified number of columns so they remain visible when scrolling
- **FromTemplate** — Load an existing `.xlsx` template and replace `{{placeholder}}` patterns with bound values
- **WithTemplateData** — Insert repeating row data into a template starting at a specified cell

### Internal

- Extracted shared `resolveCsvSettings` helper for both import and export paths
- Template exports now fire full event lifecycle (`BEFORE_EXPORT`, `BEFORE_SHEET`, `AFTER_SHEET`, `BEFORE_WRITING`)
- Auto-filter `'auto'` mode now correctly targets the last heading row for multi-row headings
- `numberToColumnLetter()` validates input and throws for non-positive or non-integer values
- `batchSize()` validated to prevent infinite loops from zero or negative values
- Cached dynamic imports of `class-validator`/`class-transformer` for performance

## v0.1.0 — Export to XLSX/CSV

### Features

- **ExcelModule** — NestJS DynamicModule with `forRoot()` / `forRootAsync()` configuration
- **ExcelService** — Core service with `download()`, `downloadAsStream()`, `store()`, and `raw()` methods
- **FromCollection** — Provide export data as an array of objects or arrays
- **FromArray** — Provide export data as a two-dimensional array
- **WithHeadings** — Add single or multiple heading rows
- **WithMapping** — Transform each row before writing
- **WithTitle** — Set worksheet tab name
- **WithMultipleSheets** — Export multiple sheets in a single workbook
- **WithColumnWidths** — Set explicit column widths
- **WithColumnFormatting** — Apply number formats to columns
- **WithStyles** — Apply font, alignment, fill, and border styles to rows, columns, or cells
- **ShouldAutoSize** — Auto-size columns based on content
- **WithProperties** — Set workbook document properties (creator, title, etc.)
- **WithCustomStartCell** — Start writing data at a specific cell
- **WithCsvSettings** — Override CSV delimiter, quote char, BOM, encoding
- **WithEvents** — Register lifecycle event listeners (beforeExport, beforeSheet, afterSheet, beforeWriting)
