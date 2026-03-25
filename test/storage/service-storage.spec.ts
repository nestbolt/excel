import { describe, it, expect, beforeEach, afterEach } from "vitest";
import { Test, TestingModule } from "@nestjs/testing";
import { Workbook } from "exceljs";
import * as fs from "fs";
import * as os from "os";
import * as path from "path";
import { ExcelService } from "../../src/excel.service";
import { ExcelModule } from "../../src/excel.module";
import { ExcelType } from "../../src/excel.constants";
import { DiskManager } from "../../src/storage/disk-manager";
import type { FromCollection, WithHeadings } from "../../src/concerns";

/* ------------------------------------------------------------------ */
/*  Test export class                                                  */
/* ------------------------------------------------------------------ */

class SimpleExport implements FromCollection, WithHeadings {
  collection() {
    return [
      [1, "Alice"],
      [2, "Bob"],
    ];
  }
  headings() {
    return ["ID", "Name"];
  }
}

/* ------------------------------------------------------------------ */
/*  Tests                                                              */
/* ------------------------------------------------------------------ */

describe("ExcelService — storage integration", () => {
  let service: ExcelService;
  let diskManager: DiskManager;
  let mod: TestingModule;
  let tmpDir: string;

  beforeEach(async () => {
    tmpDir = fs.mkdtempSync(path.join(os.tmpdir(), "storage-test-"));

    mod = await Test.createTestingModule({
      imports: [
        ExcelModule.forRoot({
          disks: {
            local: { driver: "local", root: tmpDir },
          },
          defaultDisk: "local",
        }),
      ],
    }).compile();

    service = mod.get(ExcelService);
    diskManager = mod.get(DiskManager);
  });

  afterEach(async () => {
    await mod.close();
    fs.rmSync(tmpDir, { recursive: true, force: true });
  });

  describe("store() with disk", () => {
    it("should store to the default disk", async () => {
      await service.store(new SimpleExport(), "output.xlsx");
      expect(fs.existsSync(path.join(tmpDir, "output.xlsx"))).toBe(true);
    });

    it("should store to an explicitly named disk", async () => {
      await service.store(
        new SimpleExport(),
        "report.xlsx",
        undefined,
        "local",
      );
      const filePath = path.join(tmpDir, "report.xlsx");
      expect(fs.existsSync(filePath)).toBe(true);

      const wb = new Workbook();
      await wb.xlsx.readFile(filePath);
      expect(wb.getWorksheet(1)!.getCell("A1").value).toBe("ID");
    });

    it("should create nested directories via driver", async () => {
      await service.store(
        new SimpleExport(),
        "nested/deep/report.xlsx",
        undefined,
        "local",
      );
      expect(
        fs.existsSync(path.join(tmpDir, "nested/deep/report.xlsx")),
      ).toBe(true);
    });
  });

  describe("import() with disk", () => {
    it("should import from the named disk", async () => {
      // First store a file
      await service.store(new SimpleExport(), "data.xlsx", undefined, "local");

      // Then import it back using the disk
      const result = await service.import({}, "data.xlsx", undefined, "local");
      expect(result.rows.length).toBeGreaterThan(0);
    });
  });

  describe("toArray() with disk", () => {
    it("should read file from disk and return 2D array", async () => {
      await service.store(new SimpleExport(), "arr.xlsx", undefined, "local");
      const rows = await service.toArray("arr.xlsx", undefined, "local");
      expect(rows.length).toBeGreaterThan(0);
    });
  });

  describe("toCollection() with disk", () => {
    it("should read file from disk and return objects", async () => {
      await service.store(
        new SimpleExport(),
        "coll.xlsx",
        undefined,
        "local",
      );
      const rows = await service.toCollection(
        "coll.xlsx",
        undefined,
        "local",
      );
      expect(rows.length).toBeGreaterThan(0);
      expect(rows[0]).toHaveProperty("ID");
    });
  });

  describe("DiskManager direct usage", () => {
    it("should be injectable and usable directly", async () => {
      const driver = diskManager.disk("local");
      await driver.put("direct.txt", Buffer.from("hello"));
      expect(await driver.exists("direct.txt")).toBe(true);
      const content = await driver.get("direct.txt");
      expect(content.toString()).toBe("hello");
      await driver.delete("direct.txt");
      expect(await driver.exists("direct.txt")).toBe(false);
    });
  });
});
