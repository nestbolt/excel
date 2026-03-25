import { Inject, Injectable } from "@nestjs/common";
import { EXCEL_OPTIONS } from "../excel.constants";
import type { ExcelModuleOptions } from "../interfaces";
import type { StorageDriver } from "./storage-driver.interface";
import type { DiskConfig } from "./storage.types";
import { LocalDriver } from "./drivers/local.driver";
import { S3Driver } from "./drivers/s3.driver";
import { GCSDriver } from "./drivers/gcs.driver";
import { AzureDriver } from "./drivers/azure.driver";

@Injectable()
export class DiskManager {
  private readonly drivers = new Map<string, StorageDriver>();

  constructor(
    @Inject(EXCEL_OPTIONS) private readonly options: ExcelModuleOptions,
  ) {}

  /**
   * Get the StorageDriver for the named disk.
   * Falls back to the default disk, then to an implicit LocalDriver.
   */
  disk(name?: string): StorageDriver {
    const diskName = name ?? this.options.defaultDisk ?? "local";

    const cached = this.drivers.get(diskName);
    if (cached) return cached;

    const config = this.options.disks?.[diskName];

    if (!config) {
      if (diskName === "local") {
        const driver = new LocalDriver();
        this.drivers.set(diskName, driver);
        return driver;
      }
      const available = Object.keys(this.options.disks ?? {}).join(", ");
      throw new Error(
        `Disk "${diskName}" is not configured. Available disks: ${available || "(none)"}`,
      );
    }

    const driver = this.createDriver(config);
    this.drivers.set(diskName, driver);
    return driver;
  }

  private createDriver(config: DiskConfig): StorageDriver {
    switch (config.driver) {
      case "local":
        return new LocalDriver(config);
      case "s3":
        return new S3Driver(config);
      case "gcs":
        return new GCSDriver(config);
      case "azure":
        return new AzureDriver(config);
      default:
        throw new Error(
          `Unknown storage driver: "${(config as any).driver}"`,
        );
    }
  }
}
