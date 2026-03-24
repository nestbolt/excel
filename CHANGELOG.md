# Changelog

All notable changes to `@nestbolt/excel` will be documented in this file.

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
