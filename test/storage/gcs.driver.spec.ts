import { describe, it, expect, vi, beforeEach } from "vitest";
import { GCSDriver } from "../../src/storage/drivers/gcs.driver";

describe("GCSDriver", () => {
  const mockSave = vi.fn().mockResolvedValue(undefined);
  const mockDownload = vi.fn().mockResolvedValue([Buffer.from("data")]);
  const mockDelete = vi.fn().mockResolvedValue(undefined);
  const mockExists = vi.fn().mockResolvedValue([true]);
  const mockFile = vi.fn().mockReturnValue({
    save: mockSave,
    download: mockDownload,
    delete: mockDelete,
    exists: mockExists,
  });
  const mockClient = { bucket: vi.fn().mockReturnValue({ file: mockFile }) };

  function createDriver(opts: Record<string, any> = {}) {
    return new GCSDriver({
      driver: "gcs",
      bucket: "mybucket",
      client: mockClient,
      ...opts,
    });
  }

  beforeEach(() => {
    mockSave.mockClear();
    mockDownload.mockClear();
    mockDelete.mockClear();
    mockExists.mockClear();
    mockFile.mockClear();
  });

  it("should accept a pre-configured client", () => {
    const driver = createDriver();
    expect(driver).toBeDefined();
  });

  it("should construct with SDK when no client given", () => {
    const driver = new GCSDriver({
      driver: "gcs",
      bucket: "test-bucket",
    });
    expect(driver).toBeDefined();
  });

  it("should construct with keyFilename", () => {
    const driver = new GCSDriver({
      driver: "gcs",
      bucket: "test-bucket",
      keyFilename: "/tmp/fake-key.json",
    });
    expect(driver).toBeDefined();
  });

  it("should construct with inline credentials", () => {
    const driver = new GCSDriver({
      driver: "gcs",
      bucket: "test-bucket",
      credentials: {
        client_email: "test@test.iam.gserviceaccount.com",
        private_key: "fake-key",
      },
    });
    expect(driver).toBeDefined();
  });

  describe("put", () => {
    it("should call file().save() with buffer", async () => {
      const driver = createDriver();
      const buf = Buffer.from("content");
      await driver.put("file.xlsx", buf);

      expect(mockFile).toHaveBeenCalledWith("file.xlsx");
      expect(mockSave).toHaveBeenCalledWith(buf);
    });

    it("should prepend prefix to key", async () => {
      const driver = createDriver({ prefix: "reports" });
      await driver.put("file.xlsx", Buffer.from("data"));

      expect(mockFile).toHaveBeenCalledWith("reports/file.xlsx");
    });
  });

  describe("get", () => {
    it("should return buffer from download()", async () => {
      mockDownload.mockResolvedValue([Buffer.from("file-content")]);
      const driver = createDriver();
      const result = await driver.get("file.xlsx");
      expect(result.toString()).toBe("file-content");
    });
  });

  describe("delete", () => {
    it("should call file().delete()", async () => {
      const driver = createDriver();
      await driver.delete("file.xlsx");
      expect(mockDelete).toHaveBeenCalled();
    });

    it("should ignore 404 errors", async () => {
      mockDelete.mockRejectedValue({ code: 404 });
      const driver = createDriver();
      await expect(driver.delete("missing.xlsx")).resolves.not.toThrow();
    });

    it("should rethrow non-404 errors", async () => {
      mockDelete.mockRejectedValue({ code: 403 });
      const driver = createDriver();
      await expect(driver.delete("file.xlsx")).rejects.toEqual({ code: 403 });
    });
  });

  describe("exists", () => {
    it("should return true when file exists", async () => {
      mockExists.mockResolvedValue([true]);
      const driver = createDriver();
      expect(await driver.exists("file.xlsx")).toBe(true);
    });

    it("should return false when file does not exist", async () => {
      mockExists.mockResolvedValue([false]);
      const driver = createDriver();
      expect(await driver.exists("missing.xlsx")).toBe(false);
    });
  });
});
