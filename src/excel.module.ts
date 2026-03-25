import { DynamicModule, Module } from "@nestjs/common";
import { EXCEL_OPTIONS } from "./excel.constants";
import { ExcelModuleOptions, ExcelAsyncOptions } from "./interfaces";
import { ExcelService } from "./excel.service";
import { DiskManager } from "./storage/disk-manager";

@Module({})
export class ExcelModule {
  static forRoot(options: ExcelModuleOptions = {}): DynamicModule {
    return {
      module: ExcelModule,
      global: true,
      providers: [
        { provide: EXCEL_OPTIONS, useValue: options },
        DiskManager,
        ExcelService,
      ],
      exports: [ExcelService, DiskManager],
    };
  }

  static forRootAsync(options: ExcelAsyncOptions): DynamicModule {
    return {
      module: ExcelModule,
      global: true,
      imports: options.imports ?? [],
      providers: [
        {
          provide: EXCEL_OPTIONS,
          useFactory: options.useFactory,
          inject: options.inject ?? [],
        },
        DiskManager,
        ExcelService,
      ],
      exports: [ExcelService, DiskManager],
    };
  }
}
