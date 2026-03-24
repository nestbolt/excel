import { Inject, Injectable, Logger, StreamableFile } from "@nestjs/common";
import * as fs from "fs";
import * as path from "path";
import { EXCEL_OPTIONS, ExcelType, CONTENT_TYPES } from "./excel.constants";
import type { ExcelModuleOptions, ExcelDownloadResult } from "./interfaces";
import { detectType } from "./helpers";
import { writeExport } from "./excel.writer";

@Injectable()
export class ExcelService {
  private readonly logger = new Logger(ExcelService.name);

  constructor(
    @Inject(EXCEL_OPTIONS) private readonly options: ExcelModuleOptions,
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
   * Generate the export and write it to a local file path.
   */
  async store(
    exportable: object,
    filePath: string,
    writerType?: ExcelType,
  ): Promise<void> {
    const type =
      writerType ?? this.resolveType(path.basename(filePath));
    const buffer = await writeExport(exportable, type, this.options);

    const dir = path.dirname(filePath);
    if (!fs.existsSync(dir)) {
      fs.mkdirSync(dir, { recursive: true });
    }
    fs.writeFileSync(filePath, buffer);

    this.logger.log(`Export stored at ${filePath}`);
  }

  /**
   * Generate the export and return the raw buffer.
   */
  async raw(exportable: object, writerType: ExcelType): Promise<Buffer> {
    return writeExport(exportable, writerType, this.options);
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
