import { describe, it, expect } from "vitest";
import { Test } from "@nestjs/testing";
import { EXCEL_OPTIONS } from "../../src/excel.constants";
import { DiskManager } from "../../src/storage/disk-manager";
import { LocalDriver } from "../../src/storage/drivers/local.driver";
import { S3Driver } from "../../src/storage/drivers/s3.driver";
import { GCSDriver } from "../../src/storage/drivers/gcs.driver";
import { AzureDriver } from "../../src/storage/drivers/azure.driver";

async function createDiskManager(
  options: Record<string, any> = {},
): Promise<DiskManager> {
  const module = await Test.createTestingModule({
    providers: [
      { provide: EXCEL_OPTIONS, useValue: options },
      DiskManager,
    ],
  }).compile();
  return module.get(DiskManager);
}

describe("DiskManager", () => {
  it("should return an implicit LocalDriver when no disks configured", async () => {
    const dm = await createDiskManager();
    const driver = dm.disk();
    expect(driver).toBeInstanceOf(LocalDriver);
  });

  it("should return the same cached instance on repeated calls", async () => {
    const dm = await createDiskManager();
    const d1 = dm.disk();
    const d2 = dm.disk();
    expect(d1).toBe(d2);
  });

  it("should create a LocalDriver from explicit config", async () => {
    const dm = await createDiskManager({
      disks: { local: { driver: "local", root: "/tmp" } },
    });
    expect(dm.disk("local")).toBeInstanceOf(LocalDriver);
  });

  it("should use defaultDisk when no name is given", async () => {
    const dm = await createDiskManager({
      disks: {
        mylocal: { driver: "local", root: "/tmp" },
      },
      defaultDisk: "mylocal",
    });
    const driver = dm.disk();
    expect(driver).toBeInstanceOf(LocalDriver);
  });

  it("should throw for unconfigured disk name", async () => {
    const dm = await createDiskManager({
      disks: { local: { driver: "local" } },
    });
    expect(() => dm.disk("s3")).toThrow('Disk "s3" is not configured');
  });

  it("should throw for unknown driver type", async () => {
    const dm = await createDiskManager({
      disks: { bad: { driver: "ftp" } },
    });
    expect(() => dm.disk("bad")).toThrow('Unknown storage driver: "ftp"');
  });

  it("should list available disks in error message", async () => {
    const dm = await createDiskManager({
      disks: {
        local: { driver: "local" },
        backup: { driver: "local", root: "/backup" },
      },
    });
    expect(() => dm.disk("missing")).toThrow("local, backup");
  });

  it("should show (none) when no disks configured and name is not local", async () => {
    const dm = await createDiskManager();
    expect(() => dm.disk("s3")).toThrow("(none)");
  });

  it("should create an S3Driver from config", async () => {
    const dm = await createDiskManager({
      disks: {
        s3: { driver: "s3", bucket: "test", region: "us-east-1" },
      },
    });
    expect(dm.disk("s3")).toBeInstanceOf(S3Driver);
  });

  it("should create a GCSDriver from config", async () => {
    const dm = await createDiskManager({
      disks: {
        gcs: { driver: "gcs", bucket: "test" },
      },
    });
    expect(dm.disk("gcs")).toBeInstanceOf(GCSDriver);
  });

  it("should create an AzureDriver from config", async () => {
    const dm = await createDiskManager({
      disks: {
        azure: {
          driver: "azure",
          container: "test",
          connectionString:
            "DefaultEndpointsProtocol=https;AccountName=devstoreaccount1;AccountKey=Eby8vdM02xNOcqFlqUwJPLlmEtlCDXJ1OUzFT50uSRZ6IFsuFq2UVErCz4I6tq/K1SZFPTOtr/KBHBeksoGMGw==;BlobEndpoint=https://devstoreaccount1.blob.core.windows.net;",
        },
      },
    });
    expect(dm.disk("azure")).toBeInstanceOf(AzureDriver);
  });
});
