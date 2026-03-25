import { describe, it, expect } from "vitest";
import "reflect-metadata";
import { Workbook } from "exceljs";
import {
  Exportable,
  ExportColumn,
  ExportIgnore,
  buildExportFromEntity,
} from "../src/decorators";
import { writeExport } from "../src/excel.writer";
import { ExcelType } from "../src/excel.constants";

/* ------------------------------------------------------------------ */
/*  Test entities                                                      */
/* ------------------------------------------------------------------ */

@Exportable({ title: "Users" })
class UserEntity {
  @ExportColumn({ order: 1, header: "ID" })
  id!: number;

  @ExportColumn({ order: 2 })
  firstName!: string;

  @ExportColumn({ order: 3, header: "E-Mail", format: "@" })
  email!: string;

  @ExportIgnore()
  password!: string;
}

@Exportable()
class MinimalEntity {
  @ExportColumn()
  name!: string;
}

@Exportable({
  columnWidths: { A: 5 },
  autoFilter: "auto",
  autoSize: true,
  frozenRows: 1,
  frozenColumns: 2,
})
class FullOptionsEntity {
  @ExportColumn({ order: 1, width: 20 })
  col1!: string;

  @ExportColumn({ order: 2 })
  col2!: string;
}

@Exportable()
class MappedEntity {
  @ExportColumn({
    header: "Full Name",
    map: (val: any, row: any) => `${row.first} ${row.last}`,
  })
  first!: string;

  @ExportColumn()
  last!: string;
}

/* Inheritance test */
@Exportable({ title: "Base" })
class BaseEntity {
  @ExportColumn({ order: 1, header: "ID" })
  id!: number;

  @ExportColumn({ order: 2 })
  name!: string;
}

@Exportable({ title: "Child" })
class ChildEntity extends BaseEntity {
  @ExportColumn({ order: 3, header: "Extra" })
  extra!: string;

  @ExportIgnore()
  name!: string; // override: ignore name in child
}

/* ------------------------------------------------------------------ */
/*  Helper to read XLSX buffer                                         */
/* ------------------------------------------------------------------ */

async function readBuffer(buf: Buffer): Promise<Workbook> {
  const wb = new Workbook();
  await wb.xlsx.load(buf as any);
  return wb;
}

/* ------------------------------------------------------------------ */
/*  Tests                                                              */
/* ------------------------------------------------------------------ */

describe("Export Decorators", () => {
  describe("buildExportFromEntity", () => {
    it("should throw if class is not decorated with @Exportable()", () => {
      class Plain {
        value!: string;
      }
      expect(() => buildExportFromEntity(Plain, [])).toThrow(
        'Class "Plain" is not decorated with @Exportable().',
      );
    });

    it("should throw if class has no @ExportColumn() properties", () => {
      @Exportable()
      class EmptyEntity {}
      expect(() => buildExportFromEntity(EmptyEntity, [])).toThrow(
        'Class "EmptyEntity" has no @ExportColumn() properties.',
      );
    });

    it("should build export object with headings from decorated class", () => {
      const exportObj = buildExportFromEntity(UserEntity, []) as any;
      expect(exportObj.headings()).toEqual(["ID", "First Name", "E-Mail"]);
    });

    it("should exclude @ExportIgnore() properties", () => {
      const exportObj = buildExportFromEntity(UserEntity, []) as any;
      const headings = exportObj.headings();
      expect(headings).not.toContain("Password");
      expect(headings).toHaveLength(3);
    });

    it("should map data rows correctly", () => {
      const data = [
        { id: 1, firstName: "Alice", email: "a@b.com", password: "secret" },
      ];
      const exportObj = buildExportFromEntity(UserEntity, data) as any;
      const row = exportObj.map(data[0]);
      expect(row).toEqual([1, "Alice", "a@b.com"]);
    });

    it("should return collection from data", () => {
      const data = [{ id: 1, firstName: "Alice", email: "a@b.com", password: "x" }];
      const exportObj = buildExportFromEntity(UserEntity, data) as any;
      expect(exportObj.collection()).toBe(data);
    });

    it("should set title when provided", () => {
      const exportObj = buildExportFromEntity(UserEntity, []) as any;
      expect(exportObj.title()).toBe("Users");
    });

    it("should not set title when not provided", () => {
      const exportObj = buildExportFromEntity(MinimalEntity, []) as any;
      expect(exportObj.title).toBeUndefined();
    });

    it("should auto-title-case property names without explicit header", () => {
      const exportObj = buildExportFromEntity(MinimalEntity, []) as any;
      expect(exportObj.headings()).toEqual(["Name"]);
    });

    it("should apply map function from ExportColumnOptions", () => {
      const data = [{ first: "John", last: "Doe" }];
      const exportObj = buildExportFromEntity(MappedEntity, data) as any;
      const row = exportObj.map(data[0]);
      expect(row[0]).toBe("John Doe");
    });
  });

  describe("class options", () => {
    it("should set columnWidths from class + per-column options", () => {
      const exportObj = buildExportFromEntity(FullOptionsEntity, []) as any;
      const widths = exportObj.columnWidths();
      expect(widths.A).toBe(20); // per-column overrides class-level
      expect(widths).toBeDefined();
    });

    it("should set autoFilter", () => {
      const exportObj = buildExportFromEntity(FullOptionsEntity, []) as any;
      expect(exportObj.autoFilter()).toBe("auto");
    });

    it("should set shouldAutoSize", () => {
      const exportObj = buildExportFromEntity(FullOptionsEntity, []) as any;
      expect(exportObj.shouldAutoSize).toBe(true);
    });

    it("should set frozenRows", () => {
      const exportObj = buildExportFromEntity(FullOptionsEntity, []) as any;
      expect(exportObj.frozenRows()).toBe(1);
    });

    it("should set frozenColumns", () => {
      const exportObj = buildExportFromEntity(FullOptionsEntity, []) as any;
      expect(exportObj.frozenColumns()).toBe(2);
    });

    it("should not set frozen/auto-filter when not provided", () => {
      const exportObj = buildExportFromEntity(MinimalEntity, []) as any;
      expect(exportObj.frozenRows).toBeUndefined();
      expect(exportObj.frozenColumns).toBeUndefined();
      expect(exportObj.autoFilter).toBeUndefined();
      expect(exportObj.shouldAutoSize).toBeUndefined();
    });

    it("should set columnFormats from per-column format", () => {
      const exportObj = buildExportFromEntity(UserEntity, []) as any;
      const formats = exportObj.columnFormats();
      expect(formats.C).toBe("@"); // email column (3rd → C)
    });

    it("should not set columnFormats when none specified", () => {
      const exportObj = buildExportFromEntity(MinimalEntity, []) as any;
      expect(exportObj.columnFormats).toBeUndefined();
    });
  });

  describe("inheritance", () => {
    it("should inherit parent columns and add child columns", () => {
      const exportObj = buildExportFromEntity(ChildEntity, []) as any;
      const headings = exportObj.headings();
      expect(headings).toEqual(["ID", "Extra"]);
    });

    it("should use child @Exportable title", () => {
      const exportObj = buildExportFromEntity(ChildEntity, []) as any;
      expect(exportObj.title()).toBe("Child");
    });

    it("should map inherited + child data correctly", () => {
      const data = [{ id: 1, name: "ignored", extra: "value" }];
      const exportObj = buildExportFromEntity(ChildEntity, data) as any;
      const row = exportObj.map(data[0]);
      expect(row).toEqual([1, "value"]);
    });
  });

  describe("end-to-end with writeExport", () => {
    it("should produce a valid XLSX file", async () => {
      const data = [
        { id: 1, firstName: "Alice", email: "alice@test.com", password: "s" },
        { id: 2, firstName: "Bob", email: "bob@test.com", password: "s" },
      ];
      const exportObj = buildExportFromEntity(UserEntity, data);
      const buffer = await writeExport(exportObj, ExcelType.XLSX, {});

      const wb = await readBuffer(buffer);
      const ws = wb.getWorksheet("Users")!;
      expect(ws).toBeDefined();

      // Headings row
      const headings = [
        ws.getCell("A1").value,
        ws.getCell("B1").value,
        ws.getCell("C1").value,
      ];
      expect(headings).toEqual(["ID", "First Name", "E-Mail"]);

      // Data rows
      expect(ws.getCell("A2").value).toBe(1);
      expect(ws.getCell("B2").value).toBe("Alice");
      expect(ws.getCell("C2").value).toBe("alice@test.com");
      expect(ws.getCell("A3").value).toBe(2);
    });

    it("should produce a valid CSV file", async () => {
      const data = [{ name: "Test" }];
      const exportObj = buildExportFromEntity(MinimalEntity, data);
      const buffer = await writeExport(exportObj, ExcelType.CSV, {});
      const csv = buffer.toString("utf-8");
      expect(csv).toContain("Name");
      expect(csv).toContain("Test");
    });

    it("should not include ignored columns in output", async () => {
      const data = [
        { id: 1, firstName: "Alice", email: "a@b.com", password: "secret123" },
      ];
      const exportObj = buildExportFromEntity(UserEntity, data);
      const buffer = await writeExport(exportObj, ExcelType.XLSX, {});

      const wb = await readBuffer(buffer);
      const ws = wb.getWorksheet("Users")!;
      // Should only have 3 columns (no password column)
      expect(ws.getCell("D1").value).toBeNull();

      // Ensure password value doesn't appear anywhere
      const csv = (
        await writeExport(exportObj, ExcelType.CSV, {})
      ).toString("utf-8");
      expect(csv).not.toContain("secret123");
    });
  });
});
