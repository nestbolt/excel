<p align="center">
    <h1 align="center">@nestbolt/excel</h1>
    <p align="center">Supercharged Excel and CSV exports and imports for NestJS applications. Effortlessly create, download, and import spreadsheets with powerful features and seamless integration.</p>
</p>

<p align="center">
    <a href="https://www.npmjs.com/package/@nestbolt/excel"><img src="https://img.shields.io/npm/v/@nestbolt/excel.svg?style=flat-square" alt="npm version"></a>
    <a href="https://www.npmjs.com/package/@nestbolt/excel"><img src="https://img.shields.io/npm/dt/@nestbolt/excel.svg?style=flat-square" alt="npm downloads"></a>
    <a href="https://github.com/nestbolt/excel/actions"><img src="https://img.shields.io/github/actions/workflow/status/nestbolt/excel/tests.yml?branch=main&style=flat-square&label=tests" alt="tests"></a>
    <a href="https://opensource.org/licenses/MIT"><img src="https://img.shields.io/badge/license-MIT-brightgreen.svg?style=flat-square" alt="license"></a>
</p>

<hr>

This package provides a **clean, decorator-based export API** for [NestJS](https://nestjs.com) that makes generating XLSX and CSV files effortless.

Once installed, using it is as simple as:

```typescript
@Exportable({ title: "Users" })
class UserEntity {
  @ExportColumn({ order: 1, header: "ID" })
  id!: number;

  @ExportColumn({ order: 2 })
  firstName!: string;

  @ExportColumn({ order: 3 })
  email!: string;

  @ExportIgnore()
  password!: string;
}

// In your controller
return this.excelService.downloadFromEntityAsStream(UserEntity, users, "users.xlsx");
```

## Table of Contents

- [Installation](#installation)
- [Quick Start](#quick-start)
- [Module Configuration](#module-configuration)
  - [Static Configuration (forRoot)](#static-configuration-forroot)
  - [Async Configuration (forRootAsync)](#async-configuration-forrootasync)
- [Exports — Decorator API](#exports--decorator-api)
  - [@Exportable](#exportable)
  - [@ExportColumn](#exportcolumn)
  - [@ExportIgnore](#exportignore)
  - [Decorator Options](#decorator-options)
  - [Inheritance](#inheritance)
- [Exports — Concern-based API](#exports--concern-based-api)
  - [FromCollection](#fromcollection)
  - [FromArray](#fromarray)
  - [WithHeadings](#withheadings)
  - [WithMapping](#withmapping)
  - [WithTitle](#withtitle)
  - [WithMultipleSheets](#withmultiplesheets)
  - [WithColumnWidths](#withcolumnwidths)
  - [WithColumnFormatting](#withcolumnformatting)
  - [WithStyles](#withstyles)
  - [ShouldAutoSize](#shouldautosize)
  - [WithProperties](#withproperties)
  - [WithCustomStartCell](#withcustomstartcell)
  - [WithCsvSettings](#withcsvsettings)
  - [WithEvents](#withevents)
  - [WithAutoFilter](#withautofilter)
  - [WithFrozenRows / WithFrozenColumns](#withfrozenrows--withfrozencolumns)
  - [FromTemplate](#fromtemplate)
  - [WithTemplateData](#withtemplatedata)
- [Imports](#imports)
  - [ToArray](#toarray)
  - [ToCollection](#tocollection)
  - [WithHeadingRow](#withheadingrow)
  - [WithImportMapping](#withimportmapping)
  - [WithColumnMapping](#withcolumnmapping)
  - [WithValidation](#withvalidation)
  - [WithBatchInserts](#withbatchinserts)
  - [WithStartRow](#withstartrow)
  - [WithLimit](#withlimit)
  - [SkipsEmptyRows](#skipsemptyrows)
  - [SkipsOnError](#skipsonerror)
- [Using the Service Directly](#using-the-service-directly)
- [Configuration Options](#configuration-options)
- [Testing](#testing)
- [Changelog](#changelog)
- [Contributing](#contributing)
- [Security](#security)
- [Credits](#credits)
- [License](#license)

## Installation

Install the package via npm:

```bash
npm install @nestbolt/excel
```

Or via yarn:

```bash
yarn add @nestbolt/excel
```

Or via pnpm:

```bash
pnpm add @nestbolt/excel
```

### Peer Dependencies

This package requires the following peer dependencies, which you likely already have in a NestJS project:

```
@nestjs/common   ^10.0.0 || ^11.0.0
@nestjs/core     ^10.0.0 || ^11.0.0
reflect-metadata ^0.1.13 || ^0.2.0
```

## Quick Start

### 1. Register the module

```typescript
import { ExcelModule } from "@nestbolt/excel";

@Module({
  imports: [ExcelModule.forRoot()],
})
export class AppModule {}
```

### 2. Decorate your entity or DTO

```typescript
import { Exportable, ExportColumn, ExportIgnore } from "@nestbolt/excel";

@Exportable({ title: "Users" })
export class UserEntity {
  @ExportColumn({ order: 1, header: "ID" })
  id!: number;

  @ExportColumn({ order: 2 })
  firstName!: string;

  @ExportColumn({ order: 3 })
  email!: string;

  @ExportIgnore()
  password!: string;
}
```

### 3. Use it in your controller

```typescript
import { Controller, Get } from "@nestjs/common";
import { ExcelService } from "@nestbolt/excel";
import { UserEntity } from "./user.entity";

@Controller("users")
export class UsersController {
  constructor(private readonly excelService: ExcelService) {}

  @Get("export")
  async export() {
    const users: UserEntity[] = [
      { id: 1, firstName: "Alice", email: "alice@example.com", password: "s" },
      { id: 2, firstName: "Bob", email: "bob@example.com", password: "s" },
    ];
    return this.excelService.downloadFromEntityAsStream(
      UserEntity,
      users,
      "users.xlsx",
    );
  }
}
```

The `password` field is automatically excluded from the export thanks to `@ExportIgnore()`.

## Module Configuration

### Static Configuration (forRoot)

```typescript
ExcelModule.forRoot({
  defaultType: "xlsx",
  csv: {
    delimiter: ",",
    useBom: false,
  },
});
```

### Async Configuration (forRootAsync)

```typescript
ExcelModule.forRootAsync({
  imports: [ConfigModule],
  inject: [ConfigService],
  useFactory: (config: ConfigService) => ({
    defaultType: config.get("EXCEL_DEFAULT_TYPE", "xlsx"),
  }),
});
```

The module is registered as **global** — import it once in your root module.

## Exports — Decorator API

The recommended way to define exports. Decorate your existing entities or DTOs — no separate export class needed.

### @Exportable

Mark a class as exportable. Accepts optional configuration:

```typescript
@Exportable({
  title: "Users",           // worksheet tab name
  autoFilter: "auto",       // add auto-filter to headings
  autoSize: true,           // auto-size columns to fit content
  frozenRows: 1,            // freeze heading row
  frozenColumns: 1,         // freeze first column
  columnWidths: { A: 10 },  // explicit column widths
})
class UserEntity { /* ... */ }
```

### @ExportColumn

Mark a property for export. Without options, the column header is derived from the property name (camelCase → Title Case).

```typescript
@Exportable()
class ProductEntity {
  @ExportColumn({ order: 1, header: "SKU", width: 15 })
  sku!: string;

  @ExportColumn({ order: 2, format: "#,##0.00" })
  price!: number;

  @ExportColumn({
    order: 3,
    header: "In Stock",
    map: (val) => (val ? "Yes" : "No"),
  })
  inStock!: boolean;
}
```

**Options:**

| Option   | Type                         | Description                                        |
| -------- | ---------------------------- | -------------------------------------------------- |
| `order`  | `number`                     | Column position (lower = further left)             |
| `header` | `string`                     | Column heading text                                |
| `format` | `string`                     | Excel number format (e.g. `'#,##0.00'`)            |
| `map`    | `(value, row) => any`        | Transform the value before writing                 |
| `width`  | `number`                     | Column width in character units                    |

### @ExportIgnore

Exclude a property from the export.

```typescript
@Exportable()
class UserEntity {
  @ExportColumn() name!: string;
  @ExportColumn() email!: string;
  @ExportIgnore() password!: string;  // excluded
}
```

### Decorator Options

All `@Exportable()` options map to the same concern-based features:

| Option          | Equivalent Concern   |
| --------------- | -------------------- |
| `title`         | `WithTitle`          |
| `columnWidths`  | `WithColumnWidths`   |
| `autoFilter`    | `WithAutoFilter`     |
| `autoSize`      | `ShouldAutoSize`     |
| `frozenRows`    | `WithFrozenRows`     |
| `frozenColumns` | `WithFrozenColumns`  |

### Inheritance

Decorators support class inheritance. Child classes inherit parent columns and can override or ignore them:

```typescript
@Exportable({ title: "Base" })
class BaseEntity {
  @ExportColumn({ order: 1 }) id!: number;
  @ExportColumn({ order: 2 }) name!: string;
}

@Exportable({ title: "Employees" })
class EmployeeEntity extends BaseEntity {
  @ExportColumn({ order: 3 }) department!: string;
  @ExportIgnore() name!: string;  // remove name from export
}
// Columns: ID, Department
```

### Service Methods (Decorator API)

| Method                                                    | Returns               | Description                          |
| --------------------------------------------------------- | --------------------- | ------------------------------------ |
| `downloadFromEntity(entityClass, data, filename, type?)`  | `ExcelDownloadResult` | Buffer + filename + content type     |
| `downloadFromEntityAsStream(entityClass, data, filename, type?)` | `StreamableFile` | NestJS StreamableFile for controllers |
| `storeFromEntity(entityClass, data, filePath, type?)`     | `void`                | Write to a local file                |
| `rawFromEntity(entityClass, data, type)`                  | `Buffer`              | Raw file buffer                      |

---

## Exports — Concern-based API

For advanced use cases (multiple sheets, templates, events, custom start cells, CSV settings), use the concern-based pattern. Implement one or more interfaces to opt in to features.

### FromCollection

Provide data as an array of objects or arrays. Supports async.

```typescript
class UsersExport implements FromCollection {
  async collection() {
    return await this.usersService.findAll();
  }
}
```

### FromArray

Provide data as a two-dimensional array.

```typescript
class ReportExport implements FromArray {
  array() {
    return [
      [1, "Alice", 100],
      [2, "Bob", 200],
    ];
  }
}
```

### WithHeadings

Add a heading row (or multiple rows).

```typescript
class UsersExport implements FromCollection, WithHeadings {
  collection() {
    return this.users;
  }

  headings() {
    return ["ID", "Name", "Email"];
  }
}
```

### WithMapping

Transform each row before writing.

```typescript
class UsersExport implements FromCollection, WithMapping {
  collection() {
    return this.users;
  }

  map(user: User) {
    return [user.id, `${user.firstName} ${user.lastName}`, user.email];
  }
}
```

### WithTitle

Set the worksheet tab name.

```typescript
class UsersExport implements FromCollection, WithTitle {
  collection() {
    return this.users;
  }
  title() {
    return "Active Users";
  }
}
```

### WithMultipleSheets

Export multiple sheets in one workbook.

```typescript
class MonthlyReport implements WithMultipleSheets {
  sheets() {
    return [new JanuarySheet(), new FebruarySheet(), new MarchSheet()];
  }
}
```

### WithColumnWidths

Set explicit column widths (in character units).

```typescript
columnWidths() {
  return { A: 10, B: 30, C: 20 };
}
```

### WithColumnFormatting

Apply Excel number formats.

```typescript
columnFormats() {
  return { A: '#,##0.00', B: 'yyyy-mm-dd' };
}
```

### WithStyles

Apply styles to rows, columns, or individual cells.

```typescript
styles() {
  return {
    1:    { font: { bold: true, size: 14 } },         // row 1
    'A':  { alignment: { horizontal: 'center' } },    // column A
    'B2': { fill: { fgColor: 'FFD700' } },            // cell B2
  };
}
```

### ShouldAutoSize

Auto-size all columns to fit their content.

```typescript
class UsersExport implements FromCollection, ShouldAutoSize {
  readonly shouldAutoSize = true as const;
  collection() {
    return this.users;
  }
}
```

### WithProperties

Set workbook document properties.

```typescript
properties() {
  return { creator: 'MyApp', title: 'User Report' };
}
```

### WithCustomStartCell

Start writing at a specific cell instead of A1.

```typescript
startCell() { return 'C3'; }
```

### WithCsvSettings

Override CSV options per export.

```typescript
csvSettings() {
  return { delimiter: ';', useBom: true };
}
```

### WithEvents

Hook into the export lifecycle.

```typescript
import { ExcelExportEvent } from '@nestbolt/excel';

registerEvents() {
  return {
    [ExcelExportEvent.BEFORE_EXPORT]: ({ workbook }) => { /* ... */ },
    [ExcelExportEvent.AFTER_SHEET]:   ({ worksheet }) => { /* ... */ },
  };
}
```

### WithAutoFilter

Add an auto-filter dropdown to your heading row. Use `'auto'` to automatically detect the range from your headings, or specify an explicit range.

```typescript
class UsersExport implements FromCollection, WithHeadings, WithAutoFilter {
  collection() {
    return this.users;
  }
  headings() {
    return ["ID", "Name", "Email"];
  }
  autoFilter() {
    return "auto"; // automatically covers A1:C1
  }
}
```

Or with an explicit range:

```typescript
autoFilter() {
  return "A1:D10";
}
```

### WithFrozenRows / WithFrozenColumns

Freeze rows or columns so they stay visible when scrolling.

```typescript
class UsersExport implements FromCollection, WithHeadings, WithFrozenRows {
  collection() {
    return this.users;
  }
  headings() {
    return ["ID", "Name", "Email"];
  }
  frozenRows() {
    return 1; // freeze the first row (headings)
  }
}
```

You can freeze columns too, or combine both:

```typescript
class ReportExport
  implements FromCollection, WithFrozenRows, WithFrozenColumns
{
  collection() {
    return this.data;
  }
  frozenRows() {
    return 2;
  }
  frozenColumns() {
    return 1; // freeze column A
  }
}
```

### FromTemplate

Fill an existing `.xlsx` template with data. Define placeholder bindings that replace `{{placeholder}}` patterns in the template.

```typescript
class InvoiceExport implements FromTemplate {
  templatePath() {
    return "/path/to/invoice-template.xlsx";
  }

  bindings() {
    return {
      "{{company}}": "Acme Corp",
      "{{date}}": "2026-01-15",
      "{{total}}": 1500,
    };
  }
}
```

When a cell contains exactly one placeholder and nothing else, the binding value is written with its original type (number, date, etc.). When a placeholder is embedded in a longer string, the result is a string concatenation.

### WithTemplateData

Extend `FromTemplate` with repeating row data — ideal for line items in invoices, reports, etc.

```typescript
class InvoiceExport implements FromTemplate, WithTemplateData {
  templatePath() {
    return "/path/to/invoice-template.xlsx";
  }

  bindings() {
    return {
      "{{company}}": "Acme Corp",
      "{{date}}": "2026-01-15",
      "{{total}}": 4200,
    };
  }

  dataStartCell() {
    return "A6"; // row data starts at A6
  }

  async templateData() {
    return [
      ["Widget", 10, 42],
      ["Gadget", 5, 840],
    ];
  }
}
```

The `dataStartCell()` specifies where the first row of data is written. Each subsequent row is placed on the next row below.

## Imports

Import classes use the same **concern-based** pattern as exports. Implement one or more interfaces to configure how data is read, transformed, and validated.

### Quick Import Example

```typescript
class UsersImport implements ToCollection, WithHeadingRow, WithValidation, SkipsOnError {
  readonly hasHeadingRow = true as const;
  readonly skipsOnError = true as const;

  handleCollection(rows: Record<string, any>[]) {
    // Process imported rows
  }

  rules() {
    return {
      name: [{ validate: (v) => v?.length > 0, message: "Name is required" }],
      email: [{ validate: (v) => /^.+@.+\..+$/.test(v), message: "Invalid email" }],
    };
  }
}

// In your controller
const result = await this.excelService.import(new UsersImport(), "users.xlsx");
// result.rows, result.errors, result.skipped
```

### ToArray

Receive imported data as a two-dimensional array.

```typescript
class DataImport implements ToArray {
  handleArray(rows: any[][]) {
    console.log(rows); // [[1, "Alice"], [2, "Bob"]]
  }
}
```

### ToCollection

Receive imported data as an array of objects. Requires `WithHeadingRow` or `WithColumnMapping` to derive object keys.

```typescript
class UsersImport implements ToCollection, WithHeadingRow {
  readonly hasHeadingRow = true as const;

  handleCollection(rows: Record<string, any>[]) {
    console.log(rows); // [{ ID: 1, Name: "Alice" }, ...]
  }
}
```

### WithHeadingRow

Use a row in the spreadsheet as column headings. Defaults to row 1.

```typescript
class ImportWithCustomHeading implements WithHeadingRow {
  readonly hasHeadingRow = true as const;

  headingRow() {
    return 2; // row 2 contains the headers
  }
}
```

### WithImportMapping

Transform each row after reading.

```typescript
class MappedImport implements WithHeadingRow, WithImportMapping {
  readonly hasHeadingRow = true as const;

  mapRow(row: Record<string, any>) {
    return {
      fullName: row.first_name + " " + row.last_name,
      email: row.email.toLowerCase(),
    };
  }
}
```

### WithColumnMapping

Map column letters or 1-based indices to named fields, useful for files without headers.

```typescript
class NoHeaderImport implements WithColumnMapping {
  columnMapping() {
    return { name: "A", email: "C", age: 2 };
  }
}
```

### WithValidation

Validate imported rows using custom rules or class-validator DTOs.

**Custom rules:**

```typescript
rules() {
  return {
    name: [
      { validate: (v) => v?.length > 0, message: "Name is required" },
    ],
    email: [
      { validate: (v) => /^.+@.+\..+$/.test(v), message: "Invalid email" },
    ],
  };
}
```

**class-validator DTO:**

```typescript
import { IsString, IsEmail, IsNotEmpty } from "class-validator";

class UserDto {
  @IsString() @IsNotEmpty() name!: string;
  @IsEmail() email!: string;
}

// In your import class
rules() {
  return { dto: UserDto };
}
```

> **Note:** DTO mode requires `class-validator` and `class-transformer` as peer dependencies:
> ```bash
> pnpm add class-validator class-transformer
> ```

The `ImportResult` returned from the service contains:

```typescript
interface ImportResult<T = any> {
  rows: T[];                        // valid rows
  errors: ImportValidationError[];  // per-row validation errors
  skipped: number;                  // count of skipped rows
}
```

### WithBatchInserts

Insert imported rows in configurable batch sizes.

```typescript
class BatchImport implements WithBatchInserts {
  batchSize() {
    return 100;
  }

  async handleBatch(batch: any[]) {
    await this.userRepo.save(batch);
  }
}
```

### WithStartRow

Skip rows before a given row number.

```typescript
startRow() {
  return 3; // start reading from row 3
}
```

### WithLimit

Limit the number of data rows read.

```typescript
limit() {
  return 1000; // only read first 1000 data rows
}
```

### SkipsEmptyRows

Ignore blank rows during import.

```typescript
class CleanImport implements SkipsEmptyRows {
  readonly skipsEmptyRows = true as const;
}
```

### SkipsOnError

Skip invalid rows instead of throwing. Without this concern, the first validation failure throws an error with all collected errors attached.

```typescript
class TolerantImport implements WithValidation, SkipsOnError {
  readonly skipsOnError = true as const;
  rules() { /* ... */ }
}
```

## Using the Service Directly

Inject `ExcelService` and call its methods:

### Export Methods (Concern-based)

| Method                                          | Returns               | Description                                                  |
| ----------------------------------------------- | --------------------- | ------------------------------------------------------------ |
| `download(exportable, filename, type?)`         | `ExcelDownloadResult` | Returns buffer + filename + content type                     |
| `downloadAsStream(exportable, filename, type?)` | `StreamableFile`      | Returns a NestJS StreamableFile for direct controller return |
| `store(exportable, filePath, type?)`            | `void`                | Writes the export to a local file                            |
| `raw(exportable, type)`                         | `Buffer`              | Returns the raw file buffer                                  |

### Export Methods (Decorator-based)

| Method                                                    | Returns               | Description                          |
| --------------------------------------------------------- | --------------------- | ------------------------------------ |
| `downloadFromEntity(entityClass, data, filename, type?)`  | `ExcelDownloadResult` | Buffer + filename + content type     |
| `downloadFromEntityAsStream(entityClass, data, filename, type?)` | `StreamableFile` | NestJS StreamableFile for controllers |
| `storeFromEntity(entityClass, data, filePath, type?)`     | `void`                | Write to a local file                |
| `rawFromEntity(entityClass, data, type)`                  | `Buffer`              | Raw file buffer                      |

### Import Methods

| Method                                              | Returns                      | Description                                  |
| --------------------------------------------------- | ---------------------------- | -------------------------------------------- |
| `import(importable, filePath, type?)`               | `ImportResult`               | Read and process a local file                |
| `importFromBuffer(importable, buffer, type?)`       | `ImportResult`               | Read and process a buffer                    |
| `toArray(filePath, type?)`                          | `any[][]`                    | Shorthand: returns raw 2D array              |
| `toCollection(filePath, type?)`                     | `Record<string, any>[]`      | Shorthand: returns objects using row 1 as headings |

## Configuration Options

| Option           | Type              | Default     | Description                                  |
| ---------------- | ----------------- | ----------- | -------------------------------------------- |
| `defaultType`    | `'xlsx' \| 'csv'` | `'xlsx'`    | Fallback type when extension is unrecognised |
| `tempDirectory`  | `string`          | OS temp dir | Directory for temporary files                |
| `csv.delimiter`  | `string`          | `','`       | CSV column delimiter                         |
| `csv.quoteChar`  | `string`          | `'"'`       | CSV quote character                          |
| `csv.lineEnding` | `string`          | `'\n'`      | CSV line ending                              |
| `csv.useBom`     | `boolean`         | `false`     | Prepend UTF-8 BOM                            |
| `csv.encoding`   | `BufferEncoding`  | `'utf-8'`   | Output encoding                              |

## Testing

```bash
npm test
```

Run tests in watch mode:

```bash
npm run test:watch
```

Generate coverage report:

```bash
npm run test:cov
```

## Changelog

Please see [CHANGELOG](CHANGELOG.md) for more information on what has changed recently.

## Contributing

Please see [CONTRIBUTING](CONTRIBUTING.md) for details.

## Security

If you discover any security-related issues, please report them via [GitHub Issues](https://github.com/nestbolt/excel/issues) with the **security** label instead of using the public issue tracker.

## Credits

- Built on top of [ExcelJS](https://github.com/exceljs/exceljs)

## License

The MIT License (MIT). Please see [License File](LICENSE.md) for more information.
