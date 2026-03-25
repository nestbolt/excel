import { Inject, Injectable, Logger, StreamableFile } from "@nestjs/common";
import * as path from "path";
import { EXCEL_OPTIONS, ExcelType, CONTENT_TYPES } from "./excel.constants";
import type {
  ExcelModuleOptions,
  ExcelDownloadResult,
  ImportResult,
} from "./interfaces";
import { detectType } from "./helpers";
import { writeExport } from "./excel.writer";
import { readImport } from "./excel.reader";
import { buildExportFromEntity } from "./decorators";
import { DiskManager } from "./storage/disk-manager";

@Injectable()
export class ExcelService {
  private readonly logger = new Logger(ExcelService.name);

  constructor(
    @Inject(EXCEL_OPTIONS) private readonly options: ExcelModuleOptions,
    private readonly diskManager: DiskManager,
  ) {}

  /**
   * Generate the export and return a result object suitable for building
   * an HTTP download response.
   *
   * ```ts
   * const result = await this.excelService.download(new UsersExport(), 'users.xlsx');
   * return new StreamableFile(result.buffer, {
   *   type: result.contentType,
   *   disposition: `attachment; filename="${result.filename}"`,
   * });
   * ```
   */
  async download(
    exportable: object,
    filename: string,
    writerType?: ExcelType,
  ): Promise<ExcelDownloadResult> {
    const type = writerType ?? this.resolveType(filename);
    const buffer = await writeExport(exportable, type, this.options);

    return {
      buffer,
      filename,
      contentType: CONTENT_TYPES[type],
    };
  }

  /**
   * Generate the export and return it as a NestJS `StreamableFile`,
   * ready to be returned directly from a controller method.
   *
   * ```ts
   * @Get('export')
   * export() {
   *   return this.excelService.downloadAsStream(new UsersExport(), 'users.xlsx');
   * }
   * ```
   */
  async downloadAsStream(
    exportable: object,
    filename: string,
    writerType?: ExcelType,
  ): Promise<StreamableFile> {
    const result = await this.download(exportable, filename, writerType);
    return new StreamableFile(result.buffer, {
      type: result.contentType,
      disposition: `attachment; filename="${result.filename}"`,
    });
  }

  /**
   * Generate the export and write it to storage.
   * When `disk` is specified, uses the named storage driver.
   * Otherwise uses the default disk (implicit local driver).
   */
  async store(
    exportable: object,
    filePath: string,
    writerType?: ExcelType,
    disk?: string,
  ): Promise<void> {
    const type =
      writerType ?? this.resolveType(path.basename(filePath));
    const buffer = await writeExport(exportable, type, this.options);

    const driver = this.diskManager.disk(disk);
    await driver.put(filePath, buffer);

    this.logger.log(
      `Export stored at ${filePath} (disk: ${disk ?? "default"})`,
    );
  }

  /**
   * Generate the export and return the raw buffer.
   */
  async raw(exportable: object, writerType: ExcelType): Promise<Buffer> {
    return writeExport(exportable, writerType, this.options);
  }

  /* ---------------------------------------------------------------- */
  /*  Decorator-based export                                           */
  /* ---------------------------------------------------------------- */

  /**
   * Export a decorated entity class and return a download result.
   */
  async downloadFromEntity<T>(
    entityClass: new (...args: any[]) => T,
    data: T[],
    filename: string,
    writerType?: ExcelType,
  ): Promise<ExcelDownloadResult> {
    const exportable = buildExportFromEntity(entityClass, data);
    return this.download(exportable, filename, writerType);
  }

  /**
   * Export a decorated entity class as a NestJS StreamableFile.
   */
  async downloadFromEntityAsStream<T>(
    entityClass: new (...args: any[]) => T,
    data: T[],
    filename: string,
    writerType?: ExcelType,
  ): Promise<StreamableFile> {
    const exportable = buildExportFromEntity(entityClass, data);
    return this.downloadAsStream(exportable, filename, writerType);
  }

  /**
   * Export a decorated entity class to storage.
   */
  async storeFromEntity<T>(
    entityClass: new (...args: any[]) => T,
    data: T[],
    filePath: string,
    writerType?: ExcelType,
    disk?: string,
  ): Promise<void> {
    const exportable = buildExportFromEntity(entityClass, data);
    return this.store(exportable, filePath, writerType, disk);
  }

  /**
   * Export a decorated entity class and return the raw buffer.
   */
  async rawFromEntity<T>(
    entityClass: new (...args: any[]) => T,
    data: T[],
    writerType: ExcelType,
  ): Promise<Buffer> {
    const exportable = buildExportFromEntity(entityClass, data);
    return this.raw(exportable, writerType);
  }

  /* ---------------------------------------------------------------- */
  /*  Import                                                           */
  /* ---------------------------------------------------------------- */

  /**
   * Read and process a file through the importable's concerns.
   * When `disk` is specified, reads from the named storage driver.
   */
  async import(
    importable: object,
    filePath: string,
    readerType?: ExcelType,
    disk?: string,
  ): Promise<ImportResult> {
    const type = readerType ?? this.resolveType(path.basename(filePath));
    if (disk) {
      const buffer = await this.diskManager.disk(disk).get(filePath);
      return readImport(importable, buffer, type, this.options);
    }
    return readImport(importable, filePath, type, this.options);
  }

  /**
   * Read and process a buffer through the importable's concerns.
   */
  async importFromBuffer(
    importable: object,
    buffer: Buffer,
    readerType?: ExcelType,
  ): Promise<ImportResult> {
    const type = readerType ?? ExcelType.XLSX;
    return readImport(importable, buffer, type, this.options);
  }

  /**
   * Shorthand: read a file and return the raw 2D array.
   * When `disk` is specified, reads from the named storage driver.
   */
  async toArray(
    filePath: string,
    readerType?: ExcelType,
    disk?: string,
  ): Promise<any[][]> {
    const type = readerType ?? this.resolveType(path.basename(filePath));
    if (disk) {
      const buffer = await this.diskManager.disk(disk).get(filePath);
      const result = await readImport({}, buffer, type, this.options);
      return result.rows;
    }
    const result = await readImport({}, filePath, type, this.options);
    return result.rows;
  }

  /**
   * Shorthand: read a file and return an array of objects using row 1
   * as headings.
   * When `disk` is specified, reads from the named storage driver.
   */
  async toCollection(
    filePath: string,
    readerType?: ExcelType,
    disk?: string,
  ): Promise<Record<string, any>[]> {
    const type = readerType ?? this.resolveType(path.basename(filePath));
    const importable = { hasHeadingRow: true as const };
    if (disk) {
      const buffer = await this.diskManager.disk(disk).get(filePath);
      const result = await readImport(importable, buffer, type, this.options);
      return result.rows;
    }
    const result = await readImport(importable, filePath, type, this.options);
    return result.rows;
  }

  /* ---------------------------------------------------------------- */
  /*  Internal                                                         */
  /* ---------------------------------------------------------------- */

  private resolveType(filename: string): ExcelType {
    const defaultType =
      this.options.defaultType === "csv" ? ExcelType.CSV : ExcelType.XLSX;
    return detectType(filename, defaultType);
  }
}
