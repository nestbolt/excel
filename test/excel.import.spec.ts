import { describe, it, expect, beforeEach, afterEach } from "vitest";
import { Test, TestingModule } from "@nestjs/testing";
import { Workbook } from "exceljs";
import * as fs from "fs";
import * as os from "os";
import * as path from "path";
import { IsEmail, IsNotEmpty, IsString } from "class-validator";
import { ExcelService } from "../src/excel.service";
import { EXCEL_OPTIONS, ExcelType } from "../src/excel.constants";
import { DiskManager } from "../src/storage/disk-manager";
import type {
  ToArray,
  ToCollection,
  WithHeadingRow,
  WithImportMapping,
  WithColumnMapping,
  WithValidation,
  WithBatchInserts,
  WithLimit,
  WithStartRow,
  SkipsOnError,
  SkipsEmptyRows,
  WithCsvSettings,
} from "../src/concerns";

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
  return module.get(ExcelService);
}

async function createXlsxFile(
  filePath: string,
  data: any[][],
  headings?: string[],
): Promise<void> {
  const wb = new Workbook();
  const ws = wb.addWorksheet("Sheet1");
  if (headings) {
    ws.addRow(headings);
  }
  for (const row of data) {
    ws.addRow(row);
  }
  await wb.xlsx.writeFile(filePath);
}

async function createXlsxBuffer(
  data: any[][],
  headings?: string[],
): Promise<Buffer> {
  const wb = new Workbook();
  const ws = wb.addWorksheet("Sheet1");
  if (headings) {
    ws.addRow(headings);
  }
  for (const row of data) {
    ws.addRow(row);
  }
  const arrayBuffer = await wb.xlsx.writeBuffer();
  return Buffer.from(arrayBuffer);
}

function createCsvFile(filePath: string, text: string): void {
  fs.writeFileSync(filePath, text, "utf-8");
}

/* ------------------------------------------------------------------ */
/*  Test suite                                                         */
/* ------------------------------------------------------------------ */

describe("ExcelService — Import", () => {
  let service: ExcelService;
  let tmpDir: string;

  beforeEach(async () => {
    service = await createService();
    tmpDir = fs.mkdtempSync(path.join(os.tmpdir(), "excel-import-"));
  });

  afterEach(() => {
    if (fs.existsSync(tmpDir)) {
      fs.rmSync(tmpDir, { recursive: true, force: true });
    }
  });

  /* ---------------------------------------------------------------- */
  /*  Basic XLSX import                                                */
  /* ---------------------------------------------------------------- */

  describe("basic XLSX import", () => {
    it("should import a simple XLSX file via toArray()", async () => {
      const filePath = path.join(tmpDir, "simple.xlsx");
      await createXlsxFile(filePath, [
        [1, "Alice", "alice@test.com"],
        [2, "Bob", "bob@test.com"],
      ]);

      const rows = await service.toArray(filePath);
      expect(rows).toHaveLength(2);
      expect(rows[0]).toEqual([1, "Alice", "alice@test.com"]);
      expect(rows[1]).toEqual([2, "Bob", "bob@test.com"]);
    });

    it("should import XLSX via toCollection() using row 1 as headings", async () => {
      const filePath = path.join(tmpDir, "headings.xlsx");
      await createXlsxFile(
        filePath,
        [
          [1, "Alice", "alice@test.com"],
          [2, "Bob", "bob@test.com"],
        ],
        ["ID", "Name", "Email"],
      );

      const rows = await service.toCollection(filePath);
      expect(rows).toHaveLength(2);
      expect(rows[0]).toEqual({ ID: 1, Name: "Alice", Email: "alice@test.com" });
      expect(rows[1]).toEqual({ ID: 2, Name: "Bob", Email: "bob@test.com" });
    });

    it("should import from buffer", async () => {
      const buffer = await createXlsxBuffer([
        [10, "Charlie"],
        [20, "Diana"],
      ]);

      const result = await service.importFromBuffer({}, buffer);
      expect(result.rows).toHaveLength(2);
      expect(result.rows[0]).toEqual([10, "Charlie"]);
    });

    it("should return ImportResult with errors and skipped", async () => {
      const filePath = path.join(tmpDir, "result.xlsx");
      await createXlsxFile(filePath, [[1, "Test"]]);

      const result = await service.import({}, filePath);
      expect(result.rows).toHaveLength(1);
      expect(result.errors).toEqual([]);
      expect(result.skipped).toBe(0);
    });
  });

  /* ---------------------------------------------------------------- */
  /*  Basic CSV import                                                 */
  /* ---------------------------------------------------------------- */

  describe("basic CSV import", () => {
    it("should import a CSV file", async () => {
      const filePath = path.join(tmpDir, "data.csv");
      createCsvFile(filePath, "1,Alice,alice@test.com\n2,Bob,bob@test.com\n");

      const rows = await service.toArray(filePath);
      expect(rows).toHaveLength(2);
      expect(rows[0][1]).toBe("Alice");
    });

    it("should import CSV from buffer", async () => {
      const buffer = Buffer.from("a,b\n1,2\n3,4\n", "utf-8");
      const result = await service.importFromBuffer({}, buffer, ExcelType.CSV);
      expect(result.rows).toHaveLength(3); // includes heading row as data
    });

    it("should respect WithCsvSettings on import", async () => {
      const filePath = path.join(tmpDir, "semicolons.csv");
      createCsvFile(filePath, "1;Alice\n2;Bob\n");

      class SemicolonImport implements WithCsvSettings {
        csvSettings() {
          return { delimiter: ";" };
        }
      }

      const result = await service.import(
        new SemicolonImport(),
        filePath,
        ExcelType.CSV,
      );
      expect(result.rows[0][0]).toBe(1);
      expect(result.rows[0][1]).toBe("Alice");
    });
  });

  /* ---------------------------------------------------------------- */
  /*  ToArray concern                                                  */
  /* ---------------------------------------------------------------- */

  describe("ToArray", () => {
    it("should call handleArray with data", async () => {
      let captured: any[][] = [];

      class ArrayImport implements ToArray {
        handleArray(rows: any[][]) {
          captured = rows;
        }
      }

      const filePath = path.join(tmpDir, "arr.xlsx");
      await createXlsxFile(filePath, [
        [1, "A"],
        [2, "B"],
      ]);

      await service.import(new ArrayImport(), filePath);
      expect(captured).toHaveLength(2);
      expect(captured[0]).toEqual([1, "A"]);
    });
  });

  /* ---------------------------------------------------------------- */
  /*  ToCollection concern                                             */
  /* ---------------------------------------------------------------- */

  describe("ToCollection", () => {
    it("should call handleCollection with objects", async () => {
      let captured: any[] = [];

      class CollectionImport implements ToCollection, WithHeadingRow {
        readonly hasHeadingRow = true as const;
        handleCollection(rows: Record<string, any>[]) {
          captured = rows;
        }
      }

      const filePath = path.join(tmpDir, "coll.xlsx");
      await createXlsxFile(filePath, [[1, "Alice"]], ["ID", "Name"]);

      await service.import(new CollectionImport(), filePath);
      expect(captured).toHaveLength(1);
      expect(captured[0]).toEqual({ ID: 1, Name: "Alice" });
    });
  });

  /* ---------------------------------------------------------------- */
  /*  WithHeadingRow                                                   */
  /* ---------------------------------------------------------------- */

  describe("WithHeadingRow", () => {
    it("should use row 1 as default heading row", async () => {
      class DefaultHeading implements WithHeadingRow {
        readonly hasHeadingRow = true as const;
      }

      const filePath = path.join(tmpDir, "heading.xlsx");
      await createXlsxFile(filePath, [[10, "Val"]], ["Col1", "Col2"]);

      const result = await service.import(new DefaultHeading(), filePath);
      expect(result.rows[0]).toEqual({ Col1: 10, Col2: "Val" });
    });

    it("should support custom heading row number", async () => {
      class CustomHeading implements WithHeadingRow {
        readonly hasHeadingRow = true as const;
        headingRow() {
          return 2;
        }
      }

      // Row 1: title, Row 2: headings, Row 3+: data
      const filePath = path.join(tmpDir, "custom-heading.xlsx");
      const wb = new Workbook();
      const ws = wb.addWorksheet("Sheet1");
      ws.addRow(["Report Title", ""]);
      ws.addRow(["Name", "Score"]);
      ws.addRow(["Alice", 95]);
      ws.addRow(["Bob", 88]);
      await wb.xlsx.writeFile(filePath);

      const result = await service.import(new CustomHeading(), filePath);
      expect(result.rows).toHaveLength(2);
      expect(result.rows[0]).toEqual({ Name: "Alice", Score: 95 });
    });
  });

  /* ---------------------------------------------------------------- */
  /*  WithImportMapping                                                */
  /* ---------------------------------------------------------------- */

  describe("WithImportMapping", () => {
    it("should transform each row", async () => {
      class MappedImport implements WithHeadingRow, WithImportMapping {
        readonly hasHeadingRow = true as const;
        mapRow(row: Record<string, any>) {
          return {
            fullName: String(row.Name).toUpperCase(),
            score: Number(row.Score) * 2,
          };
        }
      }

      const filePath = path.join(tmpDir, "mapped.xlsx");
      await createXlsxFile(filePath, [["Alice", 50]], ["Name", "Score"]);

      const result = await service.import(new MappedImport(), filePath);
      expect(result.rows[0]).toEqual({ fullName: "ALICE", score: 100 });
    });
  });

  /* ---------------------------------------------------------------- */
  /*  WithStartRow                                                     */
  /* ---------------------------------------------------------------- */

  describe("WithStartRow", () => {
    it("should skip rows before startRow", async () => {
      class SkipFirst implements WithStartRow {
        startRow() {
          return 3;
        }
      }

      const filePath = path.join(tmpDir, "startrow.xlsx");
      const wb = new Workbook();
      const ws = wb.addWorksheet("Sheet1");
      ws.addRow(["Title"]);
      ws.addRow(["Subtitle"]);
      ws.addRow([1, "Data1"]);
      ws.addRow([2, "Data2"]);
      await wb.xlsx.writeFile(filePath);

      const result = await service.import(new SkipFirst(), filePath);
      expect(result.rows).toHaveLength(2);
      expect(result.rows[0]).toEqual([1, "Data1"]);
    });

    it("should combine with WithHeadingRow", async () => {
      class HeadingAtRow2 implements WithHeadingRow, WithStartRow {
        readonly hasHeadingRow = true as const;
        headingRow() {
          return 2;
        }
        startRow() {
          return 3;
        }
      }

      const filePath = path.join(tmpDir, "headstart.xlsx");
      const wb = new Workbook();
      const ws = wb.addWorksheet("Sheet1");
      ws.addRow(["Junk"]);
      ws.addRow(["Name", "Age"]);
      ws.addRow(["Alice", 30]);
      await wb.xlsx.writeFile(filePath);

      const result = await service.import(new HeadingAtRow2(), filePath);
      expect(result.rows[0]).toEqual({ Name: "Alice", Age: 30 });
    });
  });

  /* ---------------------------------------------------------------- */
  /*  WithLimit                                                        */
  /* ---------------------------------------------------------------- */

  describe("WithLimit", () => {
    it("should limit number of rows read", async () => {
      class LimitedImport implements WithLimit {
        limit() {
          return 2;
        }
      }

      const filePath = path.join(tmpDir, "limit.xlsx");
      await createXlsxFile(filePath, [
        [1, "A"],
        [2, "B"],
        [3, "C"],
        [4, "D"],
      ]);

      const result = await service.import(new LimitedImport(), filePath);
      expect(result.rows).toHaveLength(2);
    });
  });

  /* ---------------------------------------------------------------- */
  /*  SkipsEmptyRows                                                   */
  /* ---------------------------------------------------------------- */

  describe("SkipsEmptyRows", () => {
    it("should filter out blank rows", async () => {
      class SkipEmpty implements SkipsEmptyRows {
        readonly skipsEmptyRows = true as const;
      }

      const filePath = path.join(tmpDir, "empty.xlsx");
      const wb = new Workbook();
      const ws = wb.addWorksheet("Sheet1");
      ws.addRow([1, "Alice"]);
      ws.addRow([null, null]);
      ws.addRow([2, "Bob"]);
      await wb.xlsx.writeFile(filePath);

      const result = await service.import(new SkipEmpty(), filePath);
      expect(result.rows).toHaveLength(2);
      expect(result.rows[0]).toEqual([1, "Alice"]);
      expect(result.rows[1]).toEqual([2, "Bob"]);
    });

    it("should keep rows with at least one non-empty cell", async () => {
      class SkipEmpty implements SkipsEmptyRows {
        readonly skipsEmptyRows = true as const;
      }

      const filePath = path.join(tmpDir, "partial.xlsx");
      const wb = new Workbook();
      const ws = wb.addWorksheet("Sheet1");
      ws.addRow([null, "Partial"]);
      ws.addRow([null, null]);
      await wb.xlsx.writeFile(filePath);

      const result = await service.import(new SkipEmpty(), filePath);
      expect(result.rows).toHaveLength(1);
    });
  });

  /* ---------------------------------------------------------------- */
  /*  WithColumnMapping                                                */
  /* ---------------------------------------------------------------- */

  describe("WithColumnMapping", () => {
    it("should map column letters to field names", async () => {
      class ColMapped implements WithColumnMapping {
        columnMapping() {
          return { name: "A", email: "C" };
        }
      }

      const filePath = path.join(tmpDir, "colmap.xlsx");
      await createXlsxFile(filePath, [
        ["Alice", 25, "alice@test.com"],
      ]);

      const result = await service.import(new ColMapped(), filePath);
      expect(result.rows[0]).toEqual({
        name: "Alice",
        email: "alice@test.com",
      });
    });

    it("should map 1-based column indices to field names", async () => {
      class IdxMapped implements WithColumnMapping {
        columnMapping() {
          return { id: 1, score: 3 };
        }
      }

      const filePath = path.join(tmpDir, "idxmap.xlsx");
      await createXlsxFile(filePath, [[100, "ignored", 95]]);

      const result = await service.import(new IdxMapped(), filePath);
      expect(result.rows[0]).toEqual({ id: 100, score: 95 });
    });
  });

  /* ---------------------------------------------------------------- */
  /*  WithValidation — custom rules                                    */
  /* ---------------------------------------------------------------- */

  describe("WithValidation — custom rules", () => {
    it("should pass valid rows through", async () => {
      class ValidImport implements WithHeadingRow, WithValidation {
        readonly hasHeadingRow = true as const;
        rules() {
          return {
            name: [
              {
                validate: (v: any) => typeof v === "string" && v.length > 0,
                message: "Name is required",
              },
            ],
          };
        }
      }

      const filePath = path.join(tmpDir, "valid.xlsx");
      await createXlsxFile(filePath, [["Alice"]], ["name"]);

      const result = await service.import(new ValidImport(), filePath);
      expect(result.rows).toHaveLength(1);
      expect(result.errors).toHaveLength(0);
    });

    it("should throw for invalid rows without SkipsOnError", async () => {
      class StrictImport implements WithHeadingRow, WithValidation {
        readonly hasHeadingRow = true as const;
        rules() {
          return {
            email: [
              {
                validate: (v: any) => /^.+@.+\..+$/.test(String(v)),
                message: "Invalid email",
              },
            ],
          };
        }
      }

      const filePath = path.join(tmpDir, "invalid.xlsx");
      await createXlsxFile(filePath, [["not-an-email"]], ["email"]);

      await expect(
        service.import(new StrictImport(), filePath),
      ).rejects.toThrow("Import validation failed");
    });

    it("should collect multiple field errors per row", async () => {
      class MultiRule
        implements WithHeadingRow, WithValidation, SkipsOnError
      {
        readonly hasHeadingRow = true as const;
        readonly skipsOnError = true as const;
        rules() {
          return {
            name: [
              {
                validate: (v: any) => typeof v === "string" && v.length > 0,
                message: "Name required",
              },
            ],
            age: [
              {
                validate: (v: any) => typeof v === "number" && v > 0,
                message: "Age must be positive",
              },
            ],
          };
        }
      }

      const filePath = path.join(tmpDir, "multi-err.xlsx");
      await createXlsxFile(filePath, [[null, -5]], ["name", "age"]);

      const result = await service.import(new MultiRule(), filePath);
      expect(result.errors).toHaveLength(1);
      expect(result.errors[0].errors).toHaveLength(2);
      expect(result.skipped).toBe(1);
      expect(result.rows).toHaveLength(0);
    });

    it("should include row number in validation errors", async () => {
      class RowNumCheck
        implements WithHeadingRow, WithValidation, SkipsOnError
      {
        readonly hasHeadingRow = true as const;
        readonly skipsOnError = true as const;
        rules() {
          return {
            val: [
              {
                validate: (v: any) => v !== null && v !== "",
                message: "Required",
              },
            ],
          };
        }
      }

      const filePath = path.join(tmpDir, "rownum.xlsx");
      await createXlsxFile(
        filePath,
        [
          ["good"],
          [""],
          ["also good"],
        ],
        ["val"],
      );

      const result = await service.import(new RowNumCheck(), filePath);
      expect(result.errors).toHaveLength(1);
      expect(result.errors[0].row).toBe(3); // heading=row1, data starts row2, bad row is row3
    });
  });

  /* ---------------------------------------------------------------- */
  /*  WithValidation — DTO mode                                        */
  /* ---------------------------------------------------------------- */

  describe("WithValidation — DTO mode", () => {
    class UserDto {
      @IsString()
      @IsNotEmpty()
      name!: string;

      @IsEmail()
      email!: string;
    }

    it("should validate valid DTO rows", async () => {
      class DtoImport implements WithHeadingRow, WithValidation {
        readonly hasHeadingRow = true as const;
        rules() {
          return { dto: UserDto };
        }
      }

      const filePath = path.join(tmpDir, "dto-valid.xlsx");
      await createXlsxFile(
        filePath,
        [["Alice", "alice@example.com"]],
        ["name", "email"],
      );

      const result = await service.import(new DtoImport(), filePath);
      expect(result.rows).toHaveLength(1);
      expect(result.errors).toHaveLength(0);
    });

    it("should collect errors for invalid DTO rows", async () => {
      class DtoSkipImport
        implements WithHeadingRow, WithValidation, SkipsOnError
      {
        readonly hasHeadingRow = true as const;
        readonly skipsOnError = true as const;
        rules() {
          return { dto: UserDto };
        }
      }

      const filePath = path.join(tmpDir, "dto-invalid.xlsx");
      await createXlsxFile(
        filePath,
        [["", "not-email"]],
        ["name", "email"],
      );

      const result = await service.import(new DtoSkipImport(), filePath);
      expect(result.rows).toHaveLength(0);
      expect(result.errors).toHaveLength(1);
      expect(result.skipped).toBe(1);

      const fieldNames = result.errors[0].errors.map((e) => e.field);
      expect(fieldNames).toContain("name");
      expect(fieldNames).toContain("email");
    });
  });

  /* ---------------------------------------------------------------- */
  /*  SkipsOnError                                                     */
  /* ---------------------------------------------------------------- */

  describe("SkipsOnError", () => {
    it("should skip invalid rows and continue", async () => {
      class SkippingImport
        implements WithHeadingRow, WithValidation, SkipsOnError
      {
        readonly hasHeadingRow = true as const;
        readonly skipsOnError = true as const;
        rules() {
          return {
            score: [
              {
                validate: (v: any) => typeof v === "number" && v >= 0,
                message: "Score must be non-negative",
              },
            ],
          };
        }
      }

      const filePath = path.join(tmpDir, "skip.xlsx");
      await createXlsxFile(
        filePath,
        [
          ["Alice", 100],
          ["Bob", -5],
          ["Charlie", 85],
        ],
        ["name", "score"],
      );

      const result = await service.import(new SkippingImport(), filePath);
      expect(result.rows).toHaveLength(2);
      expect(result.skipped).toBe(1);
      expect(result.errors).toHaveLength(1);
    });
  });

  /* ---------------------------------------------------------------- */
  /*  WithBatchInserts                                                 */
  /* ---------------------------------------------------------------- */

  describe("WithBatchInserts", () => {
    it("should deliver rows in batches", async () => {
      const batches: any[][] = [];

      class BatchImport implements WithBatchInserts {
        batchSize() {
          return 2;
        }
        handleBatch(batch: any[]) {
          batches.push([...batch]);
        }
      }

      const filePath = path.join(tmpDir, "batch.xlsx");
      await createXlsxFile(filePath, [
        [1, "A"],
        [2, "B"],
        [3, "C"],
        [4, "D"],
        [5, "E"],
      ]);

      await service.import(new BatchImport(), filePath);
      expect(batches).toHaveLength(3); // 2, 2, 1
      expect(batches[0]).toHaveLength(2);
      expect(batches[1]).toHaveLength(2);
      expect(batches[2]).toHaveLength(1);
    });

    it("should support async handleBatch", async () => {
      const results: number[] = [];

      class AsyncBatch implements WithBatchInserts {
        batchSize() {
          return 3;
        }
        async handleBatch(batch: any[]) {
          results.push(batch.length);
        }
      }

      const filePath = path.join(tmpDir, "async-batch.xlsx");
      await createXlsxFile(filePath, [
        [1],
        [2],
        [3],
        [4],
      ]);

      await service.import(new AsyncBatch(), filePath);
      expect(results).toEqual([3, 1]);
    });

    it("should throw when batchSize is zero", async () => {
      class ZeroBatch implements WithBatchInserts {
        batchSize() {
          return 0;
        }
        handleBatch() {}
      }

      const filePath = path.join(tmpDir, "zero-batch.xlsx");
      await createXlsxFile(filePath, [[1]]);

      await expect(
        service.import(new ZeroBatch(), filePath),
      ).rejects.toThrow("batchSize() must return a positive integer");
    });

    it("should throw when batchSize is negative", async () => {
      class NegBatch implements WithBatchInserts {
        batchSize() {
          return -5;
        }
        handleBatch() {}
      }

      const filePath = path.join(tmpDir, "neg-batch.xlsx");
      await createXlsxFile(filePath, [[1]]);

      await expect(
        service.import(new NegBatch(), filePath),
      ).rejects.toThrow("batchSize() must return a positive integer");
    });
  });

  /* ---------------------------------------------------------------- */
  /*  Combined concerns                                                */
  /* ---------------------------------------------------------------- */

  describe("combined concerns", () => {
    it("should run full pipeline: heading + mapping + validation + skip", async () => {
      class FullPipeline
        implements
          WithHeadingRow,
          WithImportMapping,
          WithValidation,
          SkipsOnError
      {
        readonly hasHeadingRow = true as const;
        readonly skipsOnError = true as const;

        mapRow(row: Record<string, any>) {
          return {
            name: String(row.name).trim(),
            score: Number(row.score),
          };
        }

        rules() {
          return {
            name: [
              {
                validate: (v: any) => v.length > 0,
                message: "Name required",
              },
            ],
            score: [
              {
                validate: (v: any) => !isNaN(v) && v >= 0,
                message: "Score must be non-negative number",
              },
            ],
          };
        }
      }

      const filePath = path.join(tmpDir, "full.xlsx");
      await createXlsxFile(
        filePath,
        [
          ["  Alice  ", 95],
          ["", 50],
          ["Charlie", -10],
          ["Diana", 88],
        ],
        ["name", "score"],
      );

      const result = await service.import(new FullPipeline(), filePath);
      expect(result.rows).toHaveLength(2);
      expect(result.rows[0]).toEqual({ name: "Alice", score: 95 });
      expect(result.rows[1]).toEqual({ name: "Diana", score: 88 });
      expect(result.skipped).toBe(2);
      expect(result.errors).toHaveLength(2);
    });

    it("should combine StartRow + Limit + SkipsEmptyRows", async () => {
      class CombinedLimits implements WithStartRow, WithLimit, SkipsEmptyRows {
        readonly skipsEmptyRows = true as const;
        startRow() {
          return 2;
        }
        limit() {
          return 3;
        }
      }

      const filePath = path.join(tmpDir, "combined-limits.xlsx");
      const wb = new Workbook();
      const ws = wb.addWorksheet("Sheet1");
      ws.addRow(["Header"]); // row 1, skipped by startRow
      ws.addRow([1, "A"]); // row 2
      ws.addRow([null, null]); // row 3, empty → skipped
      ws.addRow([2, "B"]); // row 4
      ws.addRow([3, "C"]); // row 5
      ws.addRow([4, "D"]); // row 6
      await wb.xlsx.writeFile(filePath);

      const result = await service.import(new CombinedLimits(), filePath);
      // After skipping row 1 (startRow=2), filtering empties, then limit 3
      expect(result.rows).toHaveLength(3);
      expect(result.rows[0]).toEqual([1, "A"]);
    });

    it("should import and re-export round-trip", async () => {
      // Export
      const exportFilePath = path.join(tmpDir, "roundtrip.xlsx");
      await createXlsxFile(
        exportFilePath,
        [
          [1, "Alice"],
          [2, "Bob"],
        ],
        ["ID", "Name"],
      );

      // Import
      const data = await service.toCollection(exportFilePath);
      expect(data).toHaveLength(2);
      expect(data[0]).toEqual({ ID: 1, Name: "Alice" });

      // Re-export
      class ReExport {
        collection() {
          return data.map((r) => [r.ID, r.Name]);
        }
      }

      const buffer = await service.raw(new ReExport(), ExcelType.XLSX);
      expect(buffer.length).toBeGreaterThan(0);
    });
  });

  /* ---------------------------------------------------------------- */
  /*  Edge cases                                                       */
  /* ---------------------------------------------------------------- */

  describe("edge cases", () => {
    it("should return empty result for empty worksheet", async () => {
      const filePath = path.join(tmpDir, "empty-ws.xlsx");
      const wb = new Workbook();
      wb.addWorksheet("Empty");
      await wb.xlsx.writeFile(filePath);

      const result = await service.import({}, filePath);
      expect(result.rows).toHaveLength(0);
      expect(result.errors).toEqual([]);
      expect(result.skipped).toBe(0);
    });

    it("should handle heading row with no data rows", async () => {
      class EmptyData implements WithHeadingRow {
        readonly hasHeadingRow = true as const;
      }

      const filePath = path.join(tmpDir, "heading-only.xlsx");
      await createXlsxFile(filePath, [], ["Col1", "Col2"]);

      const result = await service.import(new EmptyData(), filePath);
      expect(result.rows).toHaveLength(0);
    });

    it("should handle rich text cells", async () => {
      const filePath = path.join(tmpDir, "richtext.xlsx");
      const wb = new Workbook();
      const ws = wb.addWorksheet("Sheet1");
      ws.getCell("A1").value = {
        richText: [
          { text: "Hello " },
          { font: { bold: true }, text: "World" },
        ],
      } as any;
      await wb.xlsx.writeFile(filePath);

      const result = await service.import({}, filePath);
      expect(result.rows[0][0]).toBe("Hello World");
    });

    it("should convert objects to arrays for ToArray with headings", async () => {
      let captured: any[][] = [];

      class ArrayWithHeadings implements ToArray, WithHeadingRow {
        readonly hasHeadingRow = true as const;
        handleArray(rows: any[][]) {
          captured = rows;
        }
      }

      const filePath = path.join(tmpDir, "arr-head.xlsx");
      await createXlsxFile(filePath, [[1, "Alice"]], ["ID", "Name"]);

      await service.import(new ArrayWithHeadings(), filePath);
      expect(captured).toHaveLength(1);
      expect(captured[0]).toEqual([1, "Alice"]);
    });

    it("should handle formula cells by extracting result", async () => {
      const filePath = path.join(tmpDir, "formula.xlsx");
      const wb = new Workbook();
      const ws = wb.addWorksheet("Sheet1");
      ws.getCell("A1").value = 10;
      ws.getCell("B1").value = { formula: "A1*2", result: 20 } as any;
      await wb.xlsx.writeFile(filePath);

      const result = await service.import({}, filePath);
      expect(result.rows[0][0]).toBe(10);
      expect(result.rows[0][1]).toBe(20);
    });

    it("should throw when no worksheet exists in buffer", async () => {
      // Create a workbook with no worksheets — ExcelJS always adds one on
      // xlsx.writeBuffer, so we manipulate the workbook after creation.
      const wb = new Workbook();
      wb.addWorksheet("TempSheet");
      const buf = await wb.xlsx.writeBuffer();
      const wb2 = new Workbook();
      await wb2.xlsx.load(Buffer.from(buf));
      // Remove all worksheets
      while (wb2.worksheets.length > 0) {
        wb2.removeWorksheet(wb2.worksheets[0].id);
      }
      const emptyBuf = Buffer.from(await wb2.xlsx.writeBuffer());

      await expect(
        service.importFromBuffer({}, emptyBuf),
      ).rejects.toThrow("No worksheet found");
    });

    it("should use __col fallback for empty heading names", async () => {
      class EmptyHeading implements WithHeadingRow {
        readonly hasHeadingRow = true as const;
      }

      const filePath = path.join(tmpDir, "empty-heading.xlsx");
      const wb = new Workbook();
      const ws = wb.addWorksheet("Sheet1");
      ws.addRow(["Name", null, "Email"]); // middle heading is null → empty
      ws.addRow(["Alice", 25, "alice@test.com"]);
      await wb.xlsx.writeFile(filePath);

      const result = await service.import(new EmptyHeading(), filePath);
      expect(result.rows[0]).toHaveProperty("Name", "Alice");
      expect(result.rows[0]).toHaveProperty("__col1"); // fallback key
      expect(result.rows[0]).toHaveProperty("Email", "alice@test.com");
    });

    it("should return null for column mapping index beyond row length", async () => {
      class WideMapping implements WithColumnMapping {
        columnMapping() {
          return { name: "A", missing: "Z" }; // column Z won't exist in data
        }
      }

      const filePath = path.join(tmpDir, "wide-map.xlsx");
      await createXlsxFile(filePath, [["Alice"]]);

      const result = await service.import(new WideMapping(), filePath);
      expect(result.rows[0].name).toBe("Alice");
      expect(result.rows[0].missing).toBeNull();
    });

    it("should return null for heading beyond row value length", async () => {
      class WideHeading implements WithHeadingRow {
        readonly hasHeadingRow = true as const;
      }

      const filePath = path.join(tmpDir, "wide-heading.xlsx");
      const wb = new Workbook();
      const ws = wb.addWorksheet("Sheet1");
      ws.addRow(["Col1", "Col2", "Col3"]); // 3 headings
      ws.addRow(["A"]); // only 1 value — Col2 and Col3 should be null
      await wb.xlsx.writeFile(filePath);

      const result = await service.import(new WideHeading(), filePath);
      expect(result.rows[0]).toEqual({ Col1: "A", Col2: null, Col3: null });
    });

    it("should handle heading row beyond worksheet range", async () => {
      class FarHeading implements WithHeadingRow {
        readonly hasHeadingRow = true as const;
        headingRow() {
          return 999; // way beyond actual data
        }
      }

      const filePath = path.join(tmpDir, "far-heading.xlsx");
      await createXlsxFile(filePath, [["data"]]);

      // Should still work — no headings detected, data returned as arrays
      const result = await service.import(new FarHeading(), filePath);
      // Row 1 is before startRow (999+1=1000), so no data
      expect(result.rows).toHaveLength(0);
    });

    it("should attach validationErrors to thrown error", async () => {
      class StrictValidation implements WithHeadingRow, WithValidation {
        readonly hasHeadingRow = true as const;
        rules() {
          return {
            val: [{ validate: () => false, message: "Always fails" }],
          };
        }
      }

      const filePath = path.join(tmpDir, "throw-err.xlsx");
      await createXlsxFile(filePath, [["test"]], ["val"]);

      try {
        await service.import(new StrictValidation(), filePath);
        expect.fail("Should have thrown");
      } catch (err: any) {
        expect(err.validationErrors).toBeDefined();
        expect(err.validationErrors).toHaveLength(1);
        expect(err.validationErrors[0].row).toBe(2);
      }
    });
  });
});
