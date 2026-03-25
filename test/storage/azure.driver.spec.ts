import { describe, it, expect, vi, beforeEach } from "vitest";
import { AzureDriver } from "../../src/storage/drivers/azure.driver";

describe("AzureDriver", () => {
  const mockUpload = vi.fn().mockResolvedValue({});
  const mockDownload = vi.fn();
  const mockDeleteBlob = vi.fn().mockResolvedValue({});
  const mockBlobExists = vi.fn().mockResolvedValue(true);
  const mockGetBlockBlobClient = vi.fn().mockReturnValue({
    upload: mockUpload,
    download: mockDownload,
    delete: mockDeleteBlob,
    exists: mockBlobExists,
  });
  const mockContainerClient = {
    getBlockBlobClient: mockGetBlockBlobClient,
  };

  function createDriver(opts: Record<string, any> = {}) {
    return new AzureDriver({
      driver: "azure",
      container: "mycontainer",
      client: mockContainerClient,
      ...opts,
    });
  }

  beforeEach(() => {
    mockUpload.mockClear();
    mockDownload.mockClear();
    mockDeleteBlob.mockClear();
    mockBlobExists.mockClear();
    mockGetBlockBlobClient.mockClear();
  });

  it("should accept a pre-configured client", () => {
    const driver = createDriver();
    expect(driver).toBeDefined();
  });

  it("should construct with connectionString", () => {
    const driver = new AzureDriver({
      driver: "azure",
      container: "test",
      connectionString:
        "DefaultEndpointsProtocol=https;AccountName=devstoreaccount1;AccountKey=Eby8vdM02xNOcqFlqUwJPLlmEtlCDXJ1OUzFT50uSRZ6IFsuFq2UVErCz4I6tq/K1SZFPTOtr/KBHBeksoGMGw==;BlobEndpoint=https://devstoreaccount1.blob.core.windows.net;",
    });
    expect(driver).toBeDefined();
  });

  it("should construct with accountName and accountKey", () => {
    const driver = new AzureDriver({
      driver: "azure",
      container: "test",
      accountName: "devstoreaccount1",
      accountKey:
        "Eby8vdM02xNOcqFlqUwJPLlmEtlCDXJ1OUzFT50uSRZ6IFsuFq2UVErCz4I6tq/K1SZFPTOtr/KBHBeksoGMGw==",
    });
    expect(driver).toBeDefined();
  });

  it("should construct with DefaultAzureCredential fallback", () => {
    const driver = new AzureDriver({
      driver: "azure",
      container: "test",
      accountName: "devstoreaccount1",
    });
    expect(driver).toBeDefined();
  });

  it("should throw when DefaultAzureCredential branch has no accountName", () => {
    expect(
      () =>
        new AzureDriver({
          driver: "azure",
          container: "test",
        }),
    ).toThrow('AzureDriver requires "accountName"');
  });

  describe("put", () => {
    it("should upload buffer with correct length", async () => {
      const driver = createDriver();
      const buf = Buffer.from("content");
      await driver.put("file.xlsx", buf);

      expect(mockGetBlockBlobClient).toHaveBeenCalledWith("file.xlsx");
      expect(mockUpload).toHaveBeenCalledWith(buf, buf.length);
    });

    it("should prepend prefix to key", async () => {
      const driver = createDriver({ prefix: "exports" });
      await driver.put("file.xlsx", Buffer.from("data"));

      expect(mockGetBlockBlobClient).toHaveBeenCalledWith("exports/file.xlsx");
    });
  });

  describe("get", () => {
    it("should return buffer from download stream", async () => {
      const body = (async function* () {
        yield Buffer.from("chunk1");
        yield Buffer.from("chunk2");
      })();
      mockDownload.mockResolvedValue({ readableStreamBody: body });

      const driver = createDriver();
      const result = await driver.get("file.xlsx");
      expect(result.toString()).toBe("chunk1chunk2");
    });
  });

  describe("delete", () => {
    it("should call blob delete", async () => {
      const driver = createDriver();
      await driver.delete("file.xlsx");
      expect(mockDeleteBlob).toHaveBeenCalled();
    });

    it("should ignore 404 errors", async () => {
      mockDeleteBlob.mockRejectedValue({ statusCode: 404 });
      const driver = createDriver();
      await expect(driver.delete("missing.xlsx")).resolves.not.toThrow();
    });

    it("should rethrow non-404 errors", async () => {
      mockDeleteBlob.mockRejectedValue({ statusCode: 403 });
      const driver = createDriver();
      await expect(driver.delete("file.xlsx")).rejects.toEqual({
        statusCode: 403,
      });
    });
  });

  describe("exists", () => {
    it("should return true when blob exists", async () => {
      mockBlobExists.mockResolvedValue(true);
      const driver = createDriver();
      expect(await driver.exists("file.xlsx")).toBe(true);
    });

    it("should return false when blob does not exist", async () => {
      mockBlobExists.mockResolvedValue(false);
      const driver = createDriver();
      expect(await driver.exists("missing.xlsx")).toBe(false);
    });
  });
});
