import { describe, it, expect, beforeEach, afterEach } from "vitest";
import { Test, TestingModule } from "@nestjs/testing";
import { Workbook } from "exceljs";
import * as fs from "fs";
import * as os from "os";
import * as path from "path";
import { ExcelService } from "../src/excel.service";
import { ExcelModule } from "../src/excel.module";
import { EXCEL_OPTIONS, ExcelType } from "../src/excel.constants";
import { DiskManager } from "../src/storage/disk-manager";
import {
  detectType,
  parseCellRef,
  columnLetterToNumber,
  numberToColumnLetter,
} from "../src/helpers";
import type {
  FromCollection,
  FromArray,
  WithHeadings,
  WithMapping,
  WithTitle,
  WithMultipleSheets,
  WithColumnWidths,
  WithColumnFormatting,
  WithStyles,
  WithProperties,
  ShouldAutoSize,
  WithCsvSettings,
  WithEvents,
  WithCustomStartCell,
  WithAutoFilter,
  WithFrozenRows,
  WithFrozenColumns,
  FromTemplate,
  WithTemplateData,
  BeforeExportEventPayload,
  AfterSheetEventPayload,
} from "../src/concerns";
import { ExcelExportEvent } from "../src/concerns";
import { createTestTemplate } from "./fixtures/create-template";

/* ------------------------------------------------------------------ */
/*  Helpers                                                            */
/* ------------------------------------------------------------------ */

async function createService(options = {}): Promise<ExcelService> {
  const module: TestingModule = await Test.createTestingModule({
    providers: [
      { provide: EXCEL_OPTIONS, useValue: options },
      DiskManager,
      ExcelService,
    ],
  }).compile();

  return module.get<ExcelService>(ExcelService);
}

async function readXlsx(buffer: Buffer) {
  const wb = new Workbook();
  await wb.xlsx.load(buffer);
  return wb;
}

function sheetToArray(wb: Workbook, sheetIndex = 0): any[][] {
  const ws = wb.worksheets[sheetIndex];
  const rows: any[][] = [];
  ws.eachRow((row) => {
    const vals: any[] = [];
    row.eachCell({ includeEmpty: true }, (cell, colNumber) => {
      while (vals.length < colNumber - 1) vals.push(undefined);
      vals.push(cell.value);
    });
    rows.push(vals);
  });
  return rows;
}

function parseCsv(buffer: Buffer): string[][] {
  const text = buffer.toString("utf-8").replace(/^\uFEFF/, "");
  return text
    .trim()
    .split("\n")
    .map((line) =>
      line.split(",").map((cell) => cell.replace(/^"|"$/g, "").trim()),
    );
}

/* ================================================================== */
/*  Test suites                                                        */
/* ================================================================== */

describe("ExcelService", () => {
  let service: ExcelService;

  beforeEach(async () => {
    service = await createService();
  });

  /* ---------------------------------------------------------------- */
  /*  Basic XLSX export                                                */
  /* ---------------------------------------------------------------- */

  describe("XLSX export", () => {
    it("should export a simple collection to xlsx", async () => {
      class SimpleExport implements FromCollection {
        collection() {
          return [
            { id: 1, name: "Alice" },
            { id: 2, name: "Bob" },
          ];
        }
      }

      const buffer = await service.raw(new SimpleExport(), ExcelType.XLSX);
      expect(buffer).toBeInstanceOf(Buffer);
      expect(buffer.length).toBeGreaterThan(0);

      const wb = await readXlsx(buffer);
      expect(wb.worksheets).toHaveLength(1);

      const rows = sheetToArray(wb);
      expect(rows).toHaveLength(2);
      expect(rows[0]).toEqual([1, "Alice"]);
      expect(rows[1]).toEqual([2, "Bob"]);
    });

    it("should export a FromArray source", async () => {
      class ArrayExport implements FromArray {
        array() {
          return [
            [1, "Alice", "alice@test.com"],
            [2, "Bob", "bob@test.com"],
          ];
        }
      }

      const buffer = await service.raw(new ArrayExport(), ExcelType.XLSX);
      const wb = await readXlsx(buffer);
      const rows = sheetToArray(wb);
      expect(rows).toEqual([
        [1, "Alice", "alice@test.com"],
        [2, "Bob", "bob@test.com"],
      ]);
    });

    it("should support async collection()", async () => {
      class AsyncExport implements FromCollection {
        async collection() {
          return [{ val: "async-data" }];
        }
      }

      const buffer = await service.raw(new AsyncExport(), ExcelType.XLSX);
      const wb = await readXlsx(buffer);
      const rows = sheetToArray(wb);
      expect(rows[0]).toEqual(["async-data"]);
    });
  });

  /* ---------------------------------------------------------------- */
  /*  Headings                                                         */
  /* ---------------------------------------------------------------- */

  describe("WithHeadings", () => {
    it("should prepend a heading row", async () => {
      class HeadingsExport implements FromCollection, WithHeadings {
        collection() {
          return [
            { id: 1, name: "Alice" },
            { id: 2, name: "Bob" },
          ];
        }
        headings() {
          return ["ID", "Name"];
        }
      }

      const buffer = await service.raw(new HeadingsExport(), ExcelType.XLSX);
      const wb = await readXlsx(buffer);
      const rows = sheetToArray(wb);
      expect(rows[0]).toEqual(["ID", "Name"]);
      expect(rows[1]).toEqual([1, "Alice"]);
    });

    it("should support multiple heading rows", async () => {
      class MultiHeadingsExport implements FromCollection, WithHeadings {
        collection() {
          return [[10]];
        }
        headings() {
          return [
            ["Group A"],
            ["ID"],
          ] as string[][];
        }
      }

      const buffer = await service.raw(
        new MultiHeadingsExport(),
        ExcelType.XLSX,
      );
      const wb = await readXlsx(buffer);
      const rows = sheetToArray(wb);
      expect(rows[0]).toEqual(["Group A"]);
      expect(rows[1]).toEqual(["ID"]);
      expect(rows[2]).toEqual([10]);
    });
  });

  /* ---------------------------------------------------------------- */
  /*  Mapping                                                          */
  /* ---------------------------------------------------------------- */

  describe("WithMapping", () => {
    it("should transform rows with map()", async () => {
      class MappedExport implements FromCollection, WithMapping {
        collection() {
          return [
            { id: 1, first: "Alice", last: "Smith" },
            { id: 2, first: "Bob", last: "Jones" },
          ];
        }
        map(row: any) {
          return [row.id, `${row.first} ${row.last}`];
        }
      }

      const buffer = await service.raw(new MappedExport(), ExcelType.XLSX);
      const wb = await readXlsx(buffer);
      const rows = sheetToArray(wb);
      expect(rows[0]).toEqual([1, "Alice Smith"]);
      expect(rows[1]).toEqual([2, "Bob Jones"]);
    });
  });

  /* ---------------------------------------------------------------- */
  /*  Sheet title                                                      */
  /* ---------------------------------------------------------------- */

  describe("WithTitle", () => {
    it("should set the worksheet name", async () => {
      class TitledExport implements FromCollection, WithTitle {
        collection() {
          return [[1]];
        }
        title() {
          return "Users";
        }
      }

      const buffer = await service.raw(new TitledExport(), ExcelType.XLSX);
      const wb = await readXlsx(buffer);
      expect(wb.worksheets[0].name).toBe("Users");
    });
  });

  /* ---------------------------------------------------------------- */
  /*  Multiple sheets                                                  */
  /* ---------------------------------------------------------------- */

  describe("WithMultipleSheets", () => {
    it("should export multiple sheets", async () => {
      class Sheet1 implements FromCollection, WithTitle {
        collection() {
          return [["Sheet1Data"]];
        }
        title() {
          return "First";
        }
      }

      class Sheet2 implements FromCollection, WithTitle {
        collection() {
          return [["Sheet2Data"]];
        }
        title() {
          return "Second";
        }
      }

      class MultiSheetExport implements WithMultipleSheets {
        sheets() {
          return [new Sheet1(), new Sheet2()];
        }
      }

      const buffer = await service.raw(
        new MultiSheetExport(),
        ExcelType.XLSX,
      );
      const wb = await readXlsx(buffer);
      expect(wb.worksheets).toHaveLength(2);
      expect(wb.worksheets[0].name).toBe("First");
      expect(wb.worksheets[1].name).toBe("Second");

      expect(sheetToArray(wb, 0)[0]).toEqual(["Sheet1Data"]);
      expect(sheetToArray(wb, 1)[0]).toEqual(["Sheet2Data"]);
    });

    it("should throw when sheets() returns empty array", async () => {
      class EmptySheets implements WithMultipleSheets {
        sheets() {
          return [];
        }
      }

      await expect(
        service.raw(new EmptySheets(), ExcelType.XLSX),
      ).rejects.toThrow("empty array");
    });
  });

  /* ---------------------------------------------------------------- */
  /*  Column widths                                                    */
  /* ---------------------------------------------------------------- */

  describe("WithColumnWidths", () => {
    it("should set column widths", async () => {
      class WidthExport implements FromCollection, WithColumnWidths {
        collection() {
          return [["val"]];
        }
        columnWidths() {
          return { A: 25, B: 40 };
        }
      }

      const buffer = await service.raw(new WidthExport(), ExcelType.XLSX);
      const wb = await readXlsx(buffer);
      expect(wb.worksheets[0].getColumn(1).width).toBe(25);
      expect(wb.worksheets[0].getColumn(2).width).toBe(40);
    });
  });

  /* ---------------------------------------------------------------- */
  /*  Column formatting                                                */
  /* ---------------------------------------------------------------- */

  describe("WithColumnFormatting", () => {
    it("should apply number formats to columns", async () => {
      class FormatExport implements FromCollection, WithColumnFormatting {
        collection() {
          return [[1234.5, new Date(2024, 0, 15)]];
        }
        columnFormats() {
          return { A: "#,##0.00", B: "yyyy-mm-dd" };
        }
      }

      const buffer = await service.raw(new FormatExport(), ExcelType.XLSX);
      const wb = await readXlsx(buffer);
      const ws = wb.worksheets[0];
      expect(ws.getColumn(1).numFmt).toBe("#,##0.00");
      expect(ws.getColumn(2).numFmt).toBe("yyyy-mm-dd");
    });
  });

  /* ---------------------------------------------------------------- */
  /*  Styles                                                           */
  /* ---------------------------------------------------------------- */

  describe("WithStyles", () => {
    it("should apply styles to a specific cell", async () => {
      class StyledExport implements FromCollection, WithStyles {
        collection() {
          return [["Hello"]];
        }
        styles() {
          return {
            A1: { font: { bold: true, size: 14 } },
          };
        }
      }

      const buffer = await service.raw(new StyledExport(), ExcelType.XLSX);
      const wb = await readXlsx(buffer);
      const cell = wb.worksheets[0].getCell("A1");
      expect(cell.font.bold).toBe(true);
      expect(cell.font.size).toBe(14);
    });

    it("should apply styles to a row", async () => {
      class RowStyledExport
        implements FromCollection, WithHeadings, WithStyles
      {
        collection() {
          return [["data"]];
        }
        headings() {
          return ["Header"];
        }
        styles() {
          return {
            1: { font: { bold: true } },
          };
        }
      }

      const buffer = await service.raw(new RowStyledExport(), ExcelType.XLSX);
      const wb = await readXlsx(buffer);
      const cell = wb.worksheets[0].getCell("A1");
      expect(cell.font.bold).toBe(true);
    });
  });

  /* ---------------------------------------------------------------- */
  /*  ShouldAutoSize                                                   */
  /* ---------------------------------------------------------------- */

  describe("ShouldAutoSize", () => {
    it("should auto-size columns based on content", async () => {
      class AutoSizeExport implements FromCollection, ShouldAutoSize {
        readonly shouldAutoSize = true as const;
        collection() {
          return [["Short", "This is a much longer text value"]];
        }
      }

      const buffer = await service.raw(new AutoSizeExport(), ExcelType.XLSX);
      const wb = await readXlsx(buffer);
      const ws = wb.worksheets[0];
      const colAWidth = ws.getColumn(1).width ?? 0;
      const colBWidth = ws.getColumn(2).width ?? 0;

      expect(colBWidth).toBeGreaterThan(colAWidth);
    });
  });

  /* ---------------------------------------------------------------- */
  /*  Properties                                                       */
  /* ---------------------------------------------------------------- */

  describe("WithProperties", () => {
    it("should set document properties", async () => {
      class PropsExport implements FromCollection, WithProperties {
        collection() {
          return [["data"]];
        }
        properties() {
          return {
            creator: "TestApp",
            title: "Test Report",
            subject: "Testing",
          };
        }
      }

      const buffer = await service.raw(new PropsExport(), ExcelType.XLSX);
      const wb = await readXlsx(buffer);
      expect(wb.creator).toBe("TestApp");
      expect(wb.title).toBe("Test Report");
      expect(wb.subject).toBe("Testing");
    });
  });

  /* ---------------------------------------------------------------- */
  /*  Custom start cell                                                */
  /* ---------------------------------------------------------------- */

  describe("WithCustomStartCell", () => {
    it("should start data at the specified cell", async () => {
      class OffsetExport
        implements FromCollection, WithHeadings, WithCustomStartCell
      {
        collection() {
          return [["val1"]];
        }
        headings() {
          return ["Header"];
        }
        startCell() {
          return "C3";
        }
      }

      const buffer = await service.raw(new OffsetExport(), ExcelType.XLSX);
      const wb = await readXlsx(buffer);
      const ws = wb.worksheets[0];

      // Heading at C3, data at C4
      expect(ws.getCell("C3").value).toBe("Header");
      expect(ws.getCell("C4").value).toBe("val1");
      // A1 should be empty
      expect(ws.getCell("A1").value).toBeNull();
    });
  });

  /* ---------------------------------------------------------------- */
  /*  Events                                                           */
  /* ---------------------------------------------------------------- */

  describe("WithEvents", () => {
    it("should fire lifecycle events", async () => {
      const fired: string[] = [];

      class EventExport implements FromCollection, WithEvents {
        collection() {
          return [["data"]];
        }
        registerEvents() {
          return {
            [ExcelExportEvent.BEFORE_EXPORT]: (
              _e: BeforeExportEventPayload,
            ) => {
              fired.push("beforeExport");
            },
            [ExcelExportEvent.BEFORE_SHEET]: () => {
              fired.push("beforeSheet");
            },
            [ExcelExportEvent.AFTER_SHEET]: (
              _e: AfterSheetEventPayload,
            ) => {
              fired.push("afterSheet");
            },
            [ExcelExportEvent.BEFORE_WRITING]: () => {
              fired.push("beforeWriting");
            },
          };
        }
      }

      await service.raw(new EventExport(), ExcelType.XLSX);
      expect(fired).toEqual([
        "beforeExport",
        "beforeSheet",
        "afterSheet",
        "beforeWriting",
      ]);
    });
  });

  /* ---------------------------------------------------------------- */
  /*  CSV export                                                       */
  /* ---------------------------------------------------------------- */

  describe("CSV export", () => {
    it("should export as CSV", async () => {
      class CsvExport implements FromCollection, WithHeadings {
        collection() {
          return [
            { id: 1, name: "Alice" },
            { id: 2, name: "Bob" },
          ];
        }
        headings() {
          return ["ID", "Name"];
        }
      }

      const buffer = await service.raw(new CsvExport(), ExcelType.CSV);
      const rows = parseCsv(buffer);
      expect(rows[0]).toEqual(["ID", "Name"]);
      expect(rows[1]).toEqual(["1", "Alice"]);
      expect(rows[2]).toEqual(["2", "Bob"]);
    });

    it("should apply custom CSV settings", async () => {
      class CustomCsvExport implements FromCollection, WithCsvSettings {
        collection() {
          return [[1, "Alice"]];
        }
        csvSettings() {
          return { delimiter: ";", useBom: true };
        }
      }

      const buffer = await service.raw(new CustomCsvExport(), ExcelType.CSV);
      // BOM check
      expect(buffer[0]).toBe(0xef);
      expect(buffer[1]).toBe(0xbb);
      expect(buffer[2]).toBe(0xbf);

      const text = buffer.toString("utf-8").replace(/^\uFEFF/, "");
      expect(text).toContain(";");
    });
  });

  /* ---------------------------------------------------------------- */
  /*  download() and store()                                           */
  /* ---------------------------------------------------------------- */

  describe("download()", () => {
    it("should return buffer, filename, and contentType", async () => {
      class SimpleExport implements FromCollection {
        collection() {
          return [[1]];
        }
      }

      const result = await service.download(
        new SimpleExport(),
        "test.xlsx",
      );
      expect(result.buffer).toBeInstanceOf(Buffer);
      expect(result.filename).toBe("test.xlsx");
      expect(result.contentType).toBe(
        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
      );
    });

    it("should detect CSV type from filename", async () => {
      class SimpleExport implements FromCollection {
        collection() {
          return [[1]];
        }
      }

      const result = await service.download(new SimpleExport(), "test.csv");
      expect(result.contentType).toBe("text/csv");
    });
  });

  describe("downloadAsStream()", () => {
    it("should return a StreamableFile", async () => {
      class SimpleExport implements FromCollection {
        collection() {
          return [[1]];
        }
      }

      const stream = await service.downloadAsStream(
        new SimpleExport(),
        "test.xlsx",
      );
      // StreamableFile has getStream() method
      expect(stream).toBeDefined();
      expect(typeof stream.getStream).toBe("function");
    });
  });

  describe("store()", () => {
    it("should write file to disk", async () => {
      const fs = await import("fs");
      const os = await import("os");
      const path = await import("path");

      class SimpleExport implements FromCollection {
        collection() {
          return [["stored"]];
        }
      }

      const tmpDir = fs.mkdtempSync(path.join(os.tmpdir(), "excel-test-"));
      const filePath = path.join(tmpDir, "output.xlsx");

      try {
        await service.store(new SimpleExport(), filePath);
        expect(fs.existsSync(filePath)).toBe(true);

        const buffer = fs.readFileSync(filePath);
        const wb = await readXlsx(buffer);
        expect(sheetToArray(wb)[0]).toEqual(["stored"]);
      } finally {
        fs.rmSync(tmpDir, { recursive: true, force: true });
      }
    });
  });

  /* ---------------------------------------------------------------- */
  /*  Error handling                                                   */
  /* ---------------------------------------------------------------- */

  describe("error handling", () => {
    it("should throw when no data source is implemented", async () => {
      class EmptyExport {}

      await expect(
        service.raw(new EmptyExport(), ExcelType.XLSX),
      ).rejects.toThrow("FromCollection or FromArray");
    });
  });

  /* ---------------------------------------------------------------- */
  /*  Combined concerns                                                */
  /* ---------------------------------------------------------------- */

  describe("combined concerns", () => {
    it("should work with headings + mapping + title + styles", async () => {
      class FullExport
        implements FromCollection, WithHeadings, WithMapping, WithTitle, WithStyles
      {
        collection() {
          return [
            { id: 1, first: "Alice", last: "Smith", salary: 50000 },
            { id: 2, first: "Bob", last: "Jones", salary: 60000 },
          ];
        }
        headings() {
          return ["ID", "Full Name", "Salary"];
        }
        map(row: any) {
          return [row.id, `${row.first} ${row.last}`, row.salary];
        }
        title() {
          return "Employees";
        }
        styles() {
          return {
            1: { font: { bold: true } },
          };
        }
      }

      const buffer = await service.raw(new FullExport(), ExcelType.XLSX);
      const wb = await readXlsx(buffer);
      const ws = wb.worksheets[0];

      expect(ws.name).toBe("Employees");

      const rows = sheetToArray(wb);
      expect(rows[0]).toEqual(["ID", "Full Name", "Salary"]);
      expect(rows[1]).toEqual([1, "Alice Smith", 50000]);
      expect(rows[2]).toEqual([2, "Bob Jones", 60000]);

      expect(ws.getCell("A1").font.bold).toBe(true);
    });
  });

  /* ---------------------------------------------------------------- */
  /*  ExcelModule forRoot / forRootAsync                               */
  /* ---------------------------------------------------------------- */

  describe("ExcelModule", () => {
    it("forRoot() should return a valid DynamicModule", () => {
      const mod = ExcelModule.forRoot({ defaultType: "csv" });
      expect(mod.module).toBe(ExcelModule);
      expect(mod.global).toBe(true);
      expect(mod.providers).toBeDefined();
      expect(mod.exports).toBeDefined();
    });

    it("forRoot() should work with no arguments", () => {
      const mod = ExcelModule.forRoot();
      expect(mod.module).toBe(ExcelModule);
    });

    it("forRootAsync() should return a valid DynamicModule", () => {
      const mod = ExcelModule.forRootAsync({
        useFactory: () => ({ defaultType: "xlsx" }),
      });
      expect(mod.module).toBe(ExcelModule);
      expect(mod.global).toBe(true);
    });

    it("forRootAsync() should accept imports and inject", () => {
      const mod = ExcelModule.forRootAsync({
        imports: [],
        inject: ["CONFIG"],
        useFactory: () => ({}),
      });
      expect(mod.imports).toEqual([]);
    });
  });

  /* ---------------------------------------------------------------- */
  /*  Helpers: detectType, parseCellRef, columnLetterToNumber          */
  /* ---------------------------------------------------------------- */

  describe("helpers", () => {
    it("detectType should return XLSX for .xlsx", () => {
      expect(detectType("report.xlsx")).toBe(ExcelType.XLSX);
    });

    it("detectType should return CSV for .csv", () => {
      expect(detectType("report.csv")).toBe(ExcelType.CSV);
    });

    it("detectType should return fallback for unknown extension", () => {
      expect(detectType("report.unknown")).toBe(ExcelType.XLSX);
      expect(detectType("report.unknown", ExcelType.CSV)).toBe(ExcelType.CSV);
    });

    it("detectType should return fallback for no extension", () => {
      expect(detectType("report")).toBe(ExcelType.XLSX);
    });

    it("parseCellRef should parse valid references", () => {
      expect(parseCellRef("A1")).toEqual({ col: 1, row: 1 });
      expect(parseCellRef("C3")).toEqual({ col: 3, row: 3 });
      expect(parseCellRef("AA10")).toEqual({ col: 27, row: 10 });
    });

    it("parseCellRef should throw for invalid references", () => {
      expect(() => parseCellRef("")).toThrow("Invalid cell reference");
      expect(() => parseCellRef("123")).toThrow("Invalid cell reference");
      expect(() => parseCellRef("!@#")).toThrow("Invalid cell reference");
    });

    it("columnLetterToNumber should convert letters", () => {
      expect(columnLetterToNumber("A")).toBe(1);
      expect(columnLetterToNumber("Z")).toBe(26);
      expect(columnLetterToNumber("AA")).toBe(27);
      expect(columnLetterToNumber("AZ")).toBe(52);
    });
  });

  /* ---------------------------------------------------------------- */
  /*  WithStyles — full coverage (borders, fill, alignment, column)    */
  /* ---------------------------------------------------------------- */

  describe("WithStyles — full coverage", () => {
    it("should apply border styles", async () => {
      class BorderExport implements FromCollection, WithStyles {
        collection() {
          return [["data"]];
        }
        styles() {
          return {
            A1: {
              border: {
                top: { style: "thin" as const, color: "000000" },
                bottom: { style: "medium" as const, color: "FF0000" },
                left: { style: "dashed" as const },
                right: { style: "double" as const },
              },
            },
          };
        }
      }

      const buffer = await service.raw(new BorderExport(), ExcelType.XLSX);
      const wb = await readXlsx(buffer);
      const cell = wb.worksheets[0].getCell("A1");
      expect(cell.border.top?.style).toBe("thin");
      expect(cell.border.bottom?.style).toBe("medium");
      expect(cell.border.left?.style).toBe("dashed");
      expect(cell.border.right?.style).toBe("double");
    });

    it("should apply fill styles", async () => {
      class FillExport implements FromCollection, WithStyles {
        collection() {
          return [["data"]];
        }
        styles() {
          return {
            A1: {
              fill: { fgColor: "FFFF00" },
            },
          };
        }
      }

      const buffer = await service.raw(new FillExport(), ExcelType.XLSX);
      const wb = await readXlsx(buffer);
      const cell = wb.worksheets[0].getCell("A1");
      expect((cell.fill as any)?.fgColor?.argb).toBe("FFFFFF00");
    });

    it("should apply alignment styles", async () => {
      class AlignExport implements FromCollection, WithStyles {
        collection() {
          return [["data"]];
        }
        styles() {
          return {
            A1: {
              alignment: { horizontal: "center" as const, wrapText: true },
            },
          };
        }
      }

      const buffer = await service.raw(new AlignExport(), ExcelType.XLSX);
      const wb = await readXlsx(buffer);
      const cell = wb.worksheets[0].getCell("A1");
      expect(cell.alignment.horizontal).toBe("center");
      expect(cell.alignment.wrapText).toBe(true);
    });

    it("should apply numFmt via styles", async () => {
      class NumFmtExport implements FromCollection, WithStyles {
        collection() {
          return [[1234.5]];
        }
        styles() {
          return {
            A1: { numFmt: "#,##0.00" },
          };
        }
      }

      const buffer = await service.raw(new NumFmtExport(), ExcelType.XLSX);
      const wb = await readXlsx(buffer);
      const cell = wb.worksheets[0].getCell("A1");
      expect(cell.numFmt).toBe("#,##0.00");
    });

    it("should apply styles to a column letter", async () => {
      class ColStyleExport implements FromCollection, WithStyles {
        collection() {
          return [
            ["row1-a", "row1-b"],
            ["row2-a", "row2-b"],
          ];
        }
        styles() {
          return {
            A: { font: { italic: true } },
          };
        }
      }

      const buffer = await service.raw(new ColStyleExport(), ExcelType.XLSX);
      const wb = await readXlsx(buffer);
      const ws = wb.worksheets[0];
      expect(ws.getCell("A1").font.italic).toBe(true);
      expect(ws.getCell("A2").font.italic).toBe(true);
    });

    it("should apply font color and name", async () => {
      class FontDetailExport implements FromCollection, WithStyles {
        collection() {
          return [["data"]];
        }
        styles() {
          return {
            A1: {
              font: {
                name: "Arial",
                color: "#FF0000",
                underline: true,
                strike: true,
              },
            },
          };
        }
      }

      const buffer = await service.raw(
        new FontDetailExport(),
        ExcelType.XLSX,
      );
      const wb = await readXlsx(buffer);
      const cell = wb.worksheets[0].getCell("A1");
      expect(cell.font.name).toBe("Arial");
      expect(cell.font.color?.argb).toBe("FFFF0000");
      expect(cell.font.underline).toBe(true);
      expect(cell.font.strike).toBe(true);
    });

    it("should apply fill with bgColor", async () => {
      class BgFillExport implements FromCollection, WithStyles {
        collection() {
          return [["data"]];
        }
        styles() {
          return {
            A1: {
              fill: {
                type: "pattern" as const,
                pattern: "solid" as const,
                fgColor: "00FF00",
                bgColor: "0000FF",
              },
            },
          };
        }
      }

      const buffer = await service.raw(new BgFillExport(), ExcelType.XLSX);
      const wb = await readXlsx(buffer);
      const cell = wb.worksheets[0].getCell("A1");
      expect((cell.fill as any)?.fgColor?.argb).toBe("FF00FF00");
      expect((cell.fill as any)?.bgColor?.argb).toBe("FF0000FF");
    });
  });

  /* ---------------------------------------------------------------- */
  /*  WithProperties — all fields                                      */
  /* ---------------------------------------------------------------- */

  describe("WithProperties — all fields", () => {
    it("should set all document properties", async () => {
      class AllPropsExport implements FromCollection, WithProperties {
        collection() {
          return [["data"]];
        }
        properties() {
          return {
            creator: "App",
            lastModifiedBy: "Admin",
            title: "Report",
            subject: "Data",
            description: "A test report",
            keywords: "test,report",
            category: "Reports",
            company: "TestCo",
            manager: "Boss",
          };
        }
      }

      const buffer = await service.raw(new AllPropsExport(), ExcelType.XLSX);
      const wb = await readXlsx(buffer);
      expect(wb.creator).toBe("App");
      expect(wb.lastModifiedBy).toBe("Admin");
      expect(wb.title).toBe("Report");
      expect(wb.subject).toBe("Data");
      expect(wb.description).toBe("A test report");
      expect(wb.keywords).toBe("test,report");
      expect(wb.category).toBe("Reports");
      expect(wb.company).toBe("TestCo");
      expect(wb.manager).toBe("Boss");
    });
  });

  /* ---------------------------------------------------------------- */
  /*  store() — creates parent directories                             */
  /* ---------------------------------------------------------------- */

  describe("store() — nested directory creation", () => {
    it("should create parent directories if they do not exist", async () => {
      const fs = await import("fs");
      const os = await import("os");
      const path = await import("path");

      class SimpleExport implements FromCollection {
        collection() {
          return [["nested"]];
        }
      }

      const tmpDir = fs.mkdtempSync(path.join(os.tmpdir(), "excel-nest-"));
      const filePath = path.join(tmpDir, "deep", "nested", "output.xlsx");

      try {
        await service.store(new SimpleExport(), filePath);
        expect(fs.existsSync(filePath)).toBe(true);
      } finally {
        fs.rmSync(tmpDir, { recursive: true, force: true });
      }
    });
  });

  /* ---------------------------------------------------------------- */
  /*  ShouldAutoSize — edge cases                                      */
  /* ---------------------------------------------------------------- */

  describe("ShouldAutoSize — edge cases", () => {
    it("should handle null and short values in auto-size", async () => {
      class NullAutoSize implements FromCollection, ShouldAutoSize {
        readonly shouldAutoSize = true as const;
        collection() {
          return [[null, undefined, "", "ab"]];
        }
      }

      const buffer = await service.raw(new NullAutoSize(), ExcelType.XLSX);
      const wb = await readXlsx(buffer);
      const ws = wb.worksheets[0];
      // All columns should have a width (min 10 + 2 = 12)
      expect(ws.getColumn(1).width).toBeGreaterThanOrEqual(12);
    });
  });

  /* ---------------------------------------------------------------- */
  /*  WithStyles — partial borders and no-match key                    */
  /* ---------------------------------------------------------------- */

  describe("WithStyles — partial borders", () => {
    it("should handle border with only top and bottom", async () => {
      class PartialBorderExport implements FromCollection, WithStyles {
        collection() {
          return [["data"]];
        }
        styles() {
          return {
            A1: {
              border: {
                top: { style: "thin" as const },
                bottom: { style: "thin" as const },
              },
            },
          };
        }
      }

      const buffer = await service.raw(
        new PartialBorderExport(),
        ExcelType.XLSX,
      );
      const wb = await readXlsx(buffer);
      const cell = wb.worksheets[0].getCell("A1");
      expect(cell.border.top?.style).toBe("thin");
      expect(cell.border.bottom?.style).toBe("thin");
      expect(cell.border.left).toBeUndefined();
      expect(cell.border.right).toBeUndefined();
    });

    it("should handle border with only left and right", async () => {
      class LRBorderExport implements FromCollection, WithStyles {
        collection() {
          return [["data"]];
        }
        styles() {
          return {
            A1: {
              border: {
                left: { style: "medium" as const },
                right: { style: "medium" as const },
              },
            },
          };
        }
      }

      const buffer = await service.raw(new LRBorderExport(), ExcelType.XLSX);
      const wb = await readXlsx(buffer);
      const cell = wb.worksheets[0].getCell("A1");
      expect(cell.border.left?.style).toBe("medium");
      expect(cell.border.right?.style).toBe("medium");
    });
  });

  /* ---------------------------------------------------------------- */
  /*  WithEvents — partial handlers                                    */
  /* ---------------------------------------------------------------- */

  describe("WithEvents — partial handlers", () => {
    it("should handle only some events registered", async () => {
      const fired: string[] = [];

      class PartialEventExport implements FromCollection, WithEvents {
        collection() {
          return [["data"]];
        }
        registerEvents() {
          return {
            [ExcelExportEvent.AFTER_SHEET]: () => {
              fired.push("afterSheet");
            },
          };
        }
      }

      await service.raw(new PartialEventExport(), ExcelType.XLSX);
      // Only afterSheet should fire; others should not throw
      expect(fired).toEqual(["afterSheet"]);
    });
  });

  /* ---------------------------------------------------------------- */
  /*  WithProperties — partial properties                              */
  /* ---------------------------------------------------------------- */

  describe("WithProperties — partial/empty properties", () => {
    it("should handle empty properties object", async () => {
      class EmptyPropsExport implements FromCollection, WithProperties {
        collection() {
          return [["data"]];
        }
        properties() {
          return {};
        }
      }

      const buffer = await service.raw(
        new EmptyPropsExport(),
        ExcelType.XLSX,
      );
      const wb = await readXlsx(buffer);
      // Should not throw and workbook should be valid
      expect(wb.worksheets).toHaveLength(1);
    });

    it("should set only specified properties", async () => {
      class SomePropsExport implements FromCollection, WithProperties {
        collection() {
          return [["data"]];
        }
        properties() {
          return {
            description: "Only description",
            keywords: "keyword1",
            category: "Cat",
            company: "Co",
            manager: "Mgr",
          };
        }
      }

      const buffer = await service.raw(new SomePropsExport(), ExcelType.XLSX);
      const wb = await readXlsx(buffer);
      expect(wb.description).toBe("Only description");
      expect(wb.keywords).toBe("keyword1");
      expect(wb.category).toBe("Cat");
      expect(wb.company).toBe("Co");
      expect(wb.manager).toBe("Mgr");
    });
  });

  /* ---------------------------------------------------------------- */
  /*  Style edge cases — toArgb, empty fill, empty border, bad keys    */
  /* ---------------------------------------------------------------- */

  describe("WithStyles — edge cases", () => {
    it("should handle 8-char ARGB color (no FF prefix needed)", async () => {
      class ArgbExport implements FromCollection, WithStyles {
        collection() {
          return [["data"]];
        }
        styles() {
          return {
            A1: { font: { color: "FF112233" } },
          };
        }
      }

      const buffer = await service.raw(new ArgbExport(), ExcelType.XLSX);
      const wb = await readXlsx(buffer);
      const cell = wb.worksheets[0].getCell("A1");
      expect(cell.font.color?.argb).toBe("FF112233");
    });

    it("should handle fill without fgColor or bgColor", async () => {
      class EmptyFillExport implements FromCollection, WithStyles {
        collection() {
          return [["data"]];
        }
        styles() {
          return {
            A1: {
              fill: { type: "pattern" as const, pattern: "none" as const },
            },
          };
        }
      }

      const buffer = await service.raw(new EmptyFillExport(), ExcelType.XLSX);
      const wb = await readXlsx(buffer);
      expect(wb.worksheets).toHaveLength(1);
    });

    it("should handle border side with no style or color", async () => {
      class EmptyBorderExport implements FromCollection, WithStyles {
        collection() {
          return [["data"]];
        }
        styles() {
          return {
            A1: {
              border: {
                top: {},
              },
            },
          };
        }
      }

      const buffer = await service.raw(
        new EmptyBorderExport(),
        ExcelType.XLSX,
      );
      const wb = await readXlsx(buffer);
      expect(wb.worksheets).toHaveLength(1);
    });

    it("should silently ignore unrecognised style keys", async () => {
      class BadKeyExport implements FromCollection, WithStyles {
        collection() {
          return [["data"]];
        }
        styles() {
          return {
            "!!!invalid!!!": { font: { bold: true } },
          };
        }
      }

      const buffer = await service.raw(new BadKeyExport(), ExcelType.XLSX);
      const wb = await readXlsx(buffer);
      expect(sheetToArray(wb)[0]).toEqual(["data"]);
    });
  });

  /* ---------------------------------------------------------------- */
  /*  FromCollection without mapping — array items                     */
  /* ---------------------------------------------------------------- */

  describe("FromCollection — array items without mapping", () => {
    it("should pass array items through directly", async () => {
      class ArrayItemsExport implements FromCollection {
        collection() {
          return [
            [1, "a"],
            [2, "b"],
          ];
        }
      }

      const buffer = await service.raw(
        new ArrayItemsExport(),
        ExcelType.XLSX,
      );
      const wb = await readXlsx(buffer);
      expect(sheetToArray(wb)).toEqual([
        [1, "a"],
        [2, "b"],
      ]);
    });
  });

  /* ---------------------------------------------------------------- */
  /*  Module-level CSV defaults                                        */
  /* ---------------------------------------------------------------- */

  describe("module-level CSV defaults", () => {
    it("should use global CSV settings from module options", async () => {
      const svc = await createService({
        csv: { delimiter: "|" },
      });

      class SimpleExport implements FromCollection {
        collection() {
          return [["a", "b"]];
        }
      }

      const buffer = await svc.raw(new SimpleExport(), ExcelType.CSV);
      const text = buffer.toString("utf-8");
      expect(text).toContain("|");
    });
  });

  /* ---------------------------------------------------------------- */
  /*  defaultType fallback                                             */
  /* ---------------------------------------------------------------- */

  describe("defaultType option", () => {
    it("should use defaultType when extension is unknown", async () => {
      const svc = await createService({ defaultType: "csv" });

      class SimpleExport implements FromCollection {
        collection() {
          return [["val"]];
        }
      }

      const result = await svc.download(new SimpleExport(), "report.unknown");
      expect(result.contentType).toBe("text/csv");
    });
  });

  /* ---------------------------------------------------------------- */
  /*  numberToColumnLetter helper                                      */
  /* ---------------------------------------------------------------- */

  describe("numberToColumnLetter helper", () => {
    it("should convert numbers to column letters", () => {
      expect(numberToColumnLetter(1)).toBe("A");
      expect(numberToColumnLetter(26)).toBe("Z");
      expect(numberToColumnLetter(27)).toBe("AA");
      expect(numberToColumnLetter(52)).toBe("AZ");
      expect(numberToColumnLetter(703)).toBe("AAA");
    });

    it("should throw for invalid input", () => {
      expect(() => numberToColumnLetter(0)).toThrow("Invalid column number");
      expect(() => numberToColumnLetter(-1)).toThrow("Invalid column number");
      expect(() => numberToColumnLetter(1.5)).toThrow("Invalid column number");
    });
  });

  /* ---------------------------------------------------------------- */
  /*  WithAutoFilter                                                   */
  /* ---------------------------------------------------------------- */

  describe("WithAutoFilter", () => {
    it("should set auto-filter with explicit range", async () => {
      class FilterExport
        implements FromCollection, WithHeadings, WithAutoFilter
      {
        collection() {
          return [
            [1, "Alice", "alice@test.com"],
            [2, "Bob", "bob@test.com"],
          ];
        }
        headings() {
          return ["ID", "Name", "Email"];
        }
        autoFilter() {
          return "A1:C1";
        }
      }

      const buffer = await service.raw(new FilterExport(), ExcelType.XLSX);
      const wb = await readXlsx(buffer);
      expect(wb.worksheets[0].autoFilter).toBe("A1:C1");
    });

    it("should auto-detect filter range from headings", async () => {
      class AutoFilterExport
        implements FromCollection, WithHeadings, WithAutoFilter
      {
        collection() {
          return [[1, "Alice", "alice@test.com", "active"]];
        }
        headings() {
          return ["ID", "Name", "Email", "Status"];
        }
        autoFilter() {
          return "auto";
        }
      }

      const buffer = await service.raw(new AutoFilterExport(), ExcelType.XLSX);
      const wb = await readXlsx(buffer);
      expect(wb.worksheets[0].autoFilter).toBe("A1:D1");
    });

    it("should handle auto-filter with custom start cell", async () => {
      class OffsetFilterExport
        implements
          FromCollection,
          WithHeadings,
          WithAutoFilter,
          WithCustomStartCell
      {
        collection() {
          return [[1, "Alice"]];
        }
        headings() {
          return ["ID", "Name"];
        }
        autoFilter() {
          return "auto";
        }
        startCell() {
          return "C3";
        }
      }

      const buffer = await service.raw(
        new OffsetFilterExport(),
        ExcelType.XLSX,
      );
      const wb = await readXlsx(buffer);
      expect(wb.worksheets[0].autoFilter).toBe("C3:D3");
    });

    it("should place auto-filter on last heading row when multi-row headings", async () => {
      class MultiHeadingFilter
        implements FromCollection, WithHeadings, WithAutoFilter
      {
        collection() {
          return [[1, "Alice", "alice@test.com"]];
        }
        headings() {
          return [
            ["Group A", "", ""],
            ["ID", "Name", "Email"],
          ];
        }
        autoFilter() {
          return "auto";
        }
      }

      const buffer = await service.raw(
        new MultiHeadingFilter(),
        ExcelType.XLSX,
      );
      const wb = await readXlsx(buffer);
      // Should be on row 2 (the last heading row), not row 1
      expect(wb.worksheets[0].autoFilter).toBe("A2:C2");
    });

    it("should not set auto-filter when auto mode and no headings", async () => {
      class NoHeadingsFilter implements FromCollection, WithAutoFilter {
        collection() {
          return [[1, 2]];
        }
        autoFilter() {
          return "auto";
        }
      }

      const buffer = await service.raw(
        new NoHeadingsFilter(),
        ExcelType.XLSX,
      );
      const wb = await readXlsx(buffer);
      // autoFilter should not be set (no headings to detect from)
      expect(wb.worksheets[0].autoFilter).toBeUndefined();
    });
  });

  /* ---------------------------------------------------------------- */
  /*  WithFrozenRows / WithFrozenColumns                               */
  /* ---------------------------------------------------------------- */

  describe("WithFrozenRows", () => {
    it("should freeze the heading row", async () => {
      class FrozenExport
        implements FromCollection, WithHeadings, WithFrozenRows
      {
        collection() {
          return [[1, "Alice"]];
        }
        headings() {
          return ["ID", "Name"];
        }
        frozenRows() {
          return 1;
        }
      }

      const buffer = await service.raw(new FrozenExport(), ExcelType.XLSX);
      const wb = await readXlsx(buffer);
      const views = wb.worksheets[0].views;
      expect(views).toHaveLength(1);
      expect(views[0].state).toBe("frozen");
      expect(views[0].ySplit).toBe(1);
      expect(views[0].xSplit).toBe(0);
    });
  });

  describe("WithFrozenColumns", () => {
    it("should freeze columns", async () => {
      class FrozenColExport implements FromCollection, WithFrozenColumns {
        collection() {
          return [[1, "Alice", "alice@test.com"]];
        }
        frozenColumns() {
          return 2;
        }
      }

      const buffer = await service.raw(new FrozenColExport(), ExcelType.XLSX);
      const wb = await readXlsx(buffer);
      const views = wb.worksheets[0].views;
      expect(views).toHaveLength(1);
      expect(views[0].state).toBe("frozen");
      expect(views[0].xSplit).toBe(2);
      expect(views[0].ySplit).toBe(0);
    });
  });

  describe("WithFrozenRows + WithFrozenColumns", () => {
    it("should freeze both rows and columns", async () => {
      class FrozenBothExport
        implements FromCollection, WithFrozenRows, WithFrozenColumns
      {
        collection() {
          return [[1, "Alice"]];
        }
        frozenRows() {
          return 2;
        }
        frozenColumns() {
          return 1;
        }
      }

      const buffer = await service.raw(new FrozenBothExport(), ExcelType.XLSX);
      const wb = await readXlsx(buffer);
      const views = wb.worksheets[0].views;
      expect(views[0].state).toBe("frozen");
      expect(views[0].ySplit).toBe(2);
      expect(views[0].xSplit).toBe(1);
    });
  });

  /* ---------------------------------------------------------------- */
  /*  FromTemplate                                                     */
  /* ---------------------------------------------------------------- */

  describe("FromTemplate", () => {
    let tmpDir: string;
    let templatePath: string;

    beforeEach(async () => {
      tmpDir = fs.mkdtempSync(path.join(os.tmpdir(), "excel-tpl-"));
      templatePath = path.join(tmpDir, "template.xlsx");
      await createTestTemplate(templatePath);
    });

    it("should replace placeholders in template", async () => {
      class InvoiceExport implements FromTemplate {
        templatePath() {
          return templatePath;
        }
        bindings() {
          return {
            "{{company}}": "Acme Inc",
            "{{date}}": "2026-03-25",
            "{{total}}": 1500,
          };
        }
      }

      const buffer = await service.raw(new InvoiceExport(), ExcelType.XLSX);
      const wb = await readXlsx(buffer);
      const ws = wb.worksheets[0];

      expect(ws.getCell("B1").value).toBe("Acme Inc");
      expect(ws.getCell("B2").value).toBe("2026-03-25");
      expect(ws.getCell("B3").value).toBe(1500);
    });

    it("should insert repeating row data with WithTemplateData", async () => {
      class InvoiceWithItems implements FromTemplate, WithTemplateData {
        templatePath() {
          return templatePath;
        }
        bindings() {
          return {
            "{{company}}": "TestCo",
            "{{date}}": "2026-01-01",
            "{{total}}": 300,
          };
        }
        dataStartCell() {
          return "A6";
        }
        templateData() {
          return [
            ["Widget", 2, 100],
            ["Gadget", 1, 200],
          ];
        }
      }

      const buffer = await service.raw(
        new InvoiceWithItems(),
        ExcelType.XLSX,
      );
      const wb = await readXlsx(buffer);
      const ws = wb.worksheets[0];

      expect(ws.getCell("B1").value).toBe("TestCo");
      expect(ws.getCell("A6").value).toBe("Widget");
      expect(ws.getCell("B6").value).toBe(2);
      expect(ws.getCell("C6").value).toBe(100);
      expect(ws.getCell("A7").value).toBe("Gadget");
    });

    it("should preserve template formatting", async () => {
      class SimpleTemplate implements FromTemplate {
        templatePath() {
          return templatePath;
        }
        bindings() {
          return { "{{company}}": "Test" };
        }
      }

      const buffer = await service.raw(new SimpleTemplate(), ExcelType.XLSX);
      const wb = await readXlsx(buffer);
      const ws = wb.worksheets[0];

      // Row 5 was styled bold in the template
      expect(ws.getRow(5).font?.bold).toBe(true);
      // Template labels should remain
      expect(ws.getCell("A1").value).toBe("Company:");
      expect(ws.getCell("A5").value).toBe("Item");
    });

    it("should throw when template file does not exist", async () => {
      class MissingTemplate implements FromTemplate {
        templatePath() {
          return "/nonexistent/template.xlsx";
        }
        bindings() {
          return {};
        }
      }

      await expect(
        service.raw(new MissingTemplate(), ExcelType.XLSX),
      ).rejects.toThrow("Template file not found");
    });

    it("should support async templateData()", async () => {
      class AsyncTemplateData implements FromTemplate, WithTemplateData {
        templatePath() {
          return templatePath;
        }
        bindings() {
          return { "{{company}}": "Async Co" };
        }
        dataStartCell() {
          return "A6";
        }
        async templateData() {
          return [["AsyncItem", 5, 50]];
        }
      }

      const buffer = await service.raw(
        new AsyncTemplateData(),
        ExcelType.XLSX,
      );
      const wb = await readXlsx(buffer);
      expect(wb.worksheets[0].getCell("A6").value).toBe("AsyncItem");
    });

    it("should combine FromTemplate with WithProperties", async () => {
      class TemplateWithProps implements FromTemplate, WithProperties {
        templatePath() {
          return templatePath;
        }
        bindings() {
          return { "{{company}}": "PropsCo" };
        }
        properties() {
          return {
            creator: "TemplateApp",
            title: "Template Report",
            subject: "Templating",
            lastModifiedBy: "Admin",
            description: "Test",
            keywords: "kw",
            category: "Cat",
            company: "Co",
            manager: "Mgr",
          };
        }
      }

      const buffer = await service.raw(
        new TemplateWithProps(),
        ExcelType.XLSX,
      );
      const wb = await readXlsx(buffer);
      expect(wb.creator).toBe("TemplateApp");
      expect(wb.title).toBe("Template Report");
      expect(wb.worksheets[0].getCell("B1").value).toBe("PropsCo");
    });

    it("should export template as CSV", async () => {
      class TemplateCsv implements FromTemplate {
        templatePath() {
          return templatePath;
        }
        bindings() {
          return {
            "{{company}}": "CsvCo",
            "{{date}}": "2026-01-01",
            "{{total}}": 999,
          };
        }
      }

      const buffer = await service.raw(new TemplateCsv(), ExcelType.CSV);
      const text = buffer.toString("utf-8");
      expect(text).toContain("CsvCo");
      expect(text).toContain("999");
    });

    it("should replace placeholders embedded in longer strings", async () => {
      // Create a template with embedded placeholder
      const embeddedPath = path.join(tmpDir, "embedded.xlsx");
      const ewb = new Workbook();
      const ews = ewb.addWorksheet("Sheet1");
      ews.getCell("A1").value = "Invoice for {{company}} - {{date}}";
      ews.getCell("A2").value = "simple";
      await ewb.xlsx.writeFile(embeddedPath);

      class EmbeddedExport implements FromTemplate {
        templatePath() {
          return embeddedPath;
        }
        bindings() {
          return {
            "{{company}}": "Acme",
            "{{date}}": "2026-03-25",
          };
        }
      }

      const buffer = await service.raw(new EmbeddedExport(), ExcelType.XLSX);
      const wb = await readXlsx(buffer);
      expect(wb.worksheets[0].getCell("A1").value).toBe(
        "Invoice for Acme - 2026-03-25",
      );
      // Non-placeholder cell should be unchanged
      expect(wb.worksheets[0].getCell("A2").value).toBe("simple");
    });

    it("should export template as CSV with BOM", async () => {
      class TemplateCsvBom implements FromTemplate, WithCsvSettings {
        templatePath() {
          return templatePath;
        }
        bindings() {
          return { "{{company}}": "BomCo" };
        }
        csvSettings() {
          return { useBom: true };
        }
      }

      const buffer = await service.raw(new TemplateCsvBom(), ExcelType.CSV);
      expect(buffer[0]).toBe(0xef);
      expect(buffer[1]).toBe(0xbb);
      expect(buffer[2]).toBe(0xbf);
      const text = buffer.toString("utf-8").replace(/^\uFEFF/, "");
      expect(text).toContain("BomCo");
    });

    it("should handle template with partial binding replacement", async () => {
      class PartialBindings implements FromTemplate {
        templatePath() {
          return templatePath;
        }
        bindings() {
          return { "{{company}}": "Only Company" };
          // {{date}} and {{total}} are NOT replaced
        }
      }

      const buffer = await service.raw(
        new PartialBindings(),
        ExcelType.XLSX,
      );
      const wb = await readXlsx(buffer);
      const ws = wb.worksheets[0];
      expect(ws.getCell("B1").value).toBe("Only Company");
      // Unreplaced placeholders remain as-is
      expect(ws.getCell("B2").value).toBe("{{date}}");
      expect(ws.getCell("B3").value).toBe("{{total}}");
    });

    it("should fire full event lifecycle for template exports", async () => {
      const events: string[] = [];

      class TemplateWithEvents implements FromTemplate, WithEvents {
        templatePath() {
          return templatePath;
        }
        bindings() {
          return { "{{company}}": "EventCo" };
        }
        registerEvents() {
          return {
            [ExcelExportEvent.BEFORE_EXPORT]: () => {
              events.push("beforeExport");
            },
            [ExcelExportEvent.BEFORE_SHEET]: () => {
              events.push("beforeSheet");
            },
            [ExcelExportEvent.AFTER_SHEET]: () => {
              events.push("afterSheet");
            },
            [ExcelExportEvent.BEFORE_WRITING]: () => {
              events.push("beforeWriting");
            },
          };
        }
      }

      await service.raw(new TemplateWithEvents(), ExcelType.XLSX);
      expect(events).toEqual([
        "beforeExport",
        "beforeSheet",
        "afterSheet",
        "beforeWriting",
      ]);
    });

    it("should handle template with partial properties", async () => {
      class PartialPropsTemplate implements FromTemplate, WithProperties {
        templatePath() {
          return templatePath;
        }
        bindings() {
          return { "{{company}}": "TestCo" };
        }
        properties() {
          return { creator: "OnlyCreator" };
        }
      }

      const buffer = await service.raw(
        new PartialPropsTemplate(),
        ExcelType.XLSX,
      );
      const wb = await readXlsx(buffer);
      expect(wb.creator).toBe("OnlyCreator");
      // Other props should remain default/empty
      expect(wb.title).toBeFalsy();
    });

    it("should skip non-string cells in template placeholder replacement", async () => {
      const numericPath = path.join(tmpDir, "numeric-tpl.xlsx");
      const nwb = new Workbook();
      const nws = nwb.addWorksheet("Sheet1");
      nws.getCell("A1").value = "{{name}}";
      nws.getCell("B1").value = 12345; // numeric cell — not a string
      nws.getCell("C1").value = true; // boolean cell
      await nwb.xlsx.writeFile(numericPath);

      class NumericTemplate implements FromTemplate {
        templatePath() {
          return numericPath;
        }
        bindings() {
          return { "{{name}}": "Replaced" };
        }
      }

      const buffer = await service.raw(new NumericTemplate(), ExcelType.XLSX);
      const wb = await readXlsx(buffer);
      expect(wb.worksheets[0].getCell("A1").value).toBe("Replaced");
      expect(wb.worksheets[0].getCell("B1").value).toBe(12345);
      expect(wb.worksheets[0].getCell("C1").value).toBe(true);
    });

    // cleanup
    afterEach(() => {
      if (fs.existsSync(tmpDir)) {
        fs.rmSync(tmpDir, { recursive: true, force: true });
      }
    });
  });
});
