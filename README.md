<p align="center">
    <h1 align="center">@nestbolt/excel</h1>
    <p align="center">Supercharged Excel and CSV exports for NestJS applications. Effortlessly create and download spreadsheets with powerful features and seamless integration.</p>
</p>

<p align="center">
    <a href="https://www.npmjs.com/package/@nestbolt/excel"><img src="https://img.shields.io/npm/v/@nestbolt/excel.svg?style=flat-square" alt="npm version"></a>
    <a href="https://www.npmjs.com/package/@nestbolt/excel"><img src="https://img.shields.io/npm/dt/@nestbolt/excel.svg?style=flat-square" alt="npm downloads"></a>
    <a href="https://github.com/nestbolt/excel/actions"><img src="https://img.shields.io/github/actions/workflow/status/nestbolt/excel/tests.yml?branch=main&style=flat-square&label=tests" alt="tests"></a>
    <a href="https://opensource.org/licenses/MIT"><img src="https://img.shields.io/badge/license-MIT-brightgreen.svg?style=flat-square" alt="license"></a>
</p>

<hr>

This package provides a **clean, concern-based export API** for [NestJS](https://nestjs.com) that makes generating XLSX and CSV files effortless.

Once installed, using it is as simple as:

```typescript
class UsersExport implements FromCollection, WithHeadings {
  collection() {
    return this.users;
  }
  headings() {
    return ["ID", "Name", "Email"];
  }
}

// In your controller
return this.excelService.downloadAsStream(new UsersExport(), "users.xlsx");
```

## Table of Contents

- [Installation](#installation)
- [Quick Start](#quick-start)
- [Module Configuration](#module-configuration)
  - [Static Configuration (forRoot)](#static-configuration-forroot)
  - [Async Configuration (forRootAsync)](#async-configuration-forrootasync)
- [Exports](#exports)
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

### 2. Create an export class

```typescript
import { FromCollection, WithHeadings } from "@nestbolt/excel";

export class UsersExport implements FromCollection, WithHeadings {
  constructor(private readonly users: any[]) {}

  collection() {
    return this.users;
  }

  headings() {
    return ["ID", "Name", "Email"];
  }
}
```

### 3. Use it in your controller

```typescript
import { Controller, Get } from "@nestjs/common";
import { ExcelService } from "@nestbolt/excel";
import { UsersExport } from "./users.export";

@Controller("users")
export class UsersController {
  constructor(private readonly excelService: ExcelService) {}

  @Get("export")
  async export() {
    const users = [
      { id: 1, name: "Alice", email: "alice@example.com" },
      { id: 2, name: "Bob", email: "bob@example.com" },
    ];
    return this.excelService.downloadAsStream(
      new UsersExport(users),
      "users.xlsx",
    );
  }
}
```

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

## Exports

Export classes use a **concern-based** pattern. Implement one or more interfaces to opt in to features.

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

## Using the Service Directly

Inject `ExcelService` and call its methods:

| Method                                          | Returns               | Description                                                  |
| ----------------------------------------------- | --------------------- | ------------------------------------------------------------ |
| `download(exportable, filename, type?)`         | `ExcelDownloadResult` | Returns buffer + filename + content type                     |
| `downloadAsStream(exportable, filename, type?)` | `StreamableFile`      | Returns a NestJS StreamableFile for direct controller return |
| `store(exportable, filePath, type?)`            | `void`                | Writes the export to a local file                            |
| `raw(exportable, type)`                         | `Buffer`              | Returns the raw file buffer                                  |

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
