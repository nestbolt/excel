import { describe, it, expect, beforeEach } from "vitest";
import { Test, TestingModule } from "@nestjs/testing";
import { Workbook } from "exceljs";
import { ExcelService } from "../src/excel.service";
import { EXCEL_OPTIONS, ExcelType } from "../src/excel.constants";
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
  BeforeExportEventPayload,
  AfterSheetEventPayload,
} from "../src/concerns";
import { ExcelExportEvent } from "../src/concerns";

/* ------------------------------------------------------------------ */
/*  Helpers                                                            */
/* ------------------------------------------------------------------ */

async function createService(options = {}): Promise<ExcelService> {
  const module: TestingModule = await Test.createTestingModule({
    providers: [
      { provide: EXCEL_OPTIONS, useValue: options },
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
});
