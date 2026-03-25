import { describe, it, expect, vi, beforeEach } from "vitest";
import { S3Driver } from "../../src/storage/drivers/s3.driver";

describe("S3Driver", () => {
  const mockSend = vi.fn();
  const mockClient = { send: mockSend };

  function createDriver(opts: Record<string, any> = {}) {
    return new S3Driver({
      driver: "s3",
      bucket: "mybucket",
      client: mockClient,
      ...opts,
    });
  }

  beforeEach(() => {
    mockSend.mockReset();
  });

  it("should accept a pre-configured client", () => {
    const driver = createDriver();
    expect(driver).toBeDefined();
  });

  it("should construct with SDK when no client given", () => {
    const driver = new S3Driver({
      driver: "s3",
      bucket: "b",
      region: "us-east-1",
    });
    expect(driver).toBeDefined();
  });

  it("should construct with credentials", () => {
    const driver = new S3Driver({
      driver: "s3",
      bucket: "b",
      region: "us-east-1",
      credentials: {
        accessKeyId: "test",
        secretAccessKey: "test",
      },
    });
    expect(driver).toBeDefined();
  });

  it("should construct with endpoint for S3-compatible services", () => {
    const driver = new S3Driver({
      driver: "s3",
      bucket: "b",
      region: "us-east-1",
      endpoint: "http://localhost:9000",
    });
    expect(driver).toBeDefined();
  });

  it("should construct without optional region/credentials/endpoint", () => {
    const driver = new S3Driver({
      driver: "s3",
      bucket: "b",
      region: "us-east-1",
    });
    expect(driver).toBeDefined();
  });

  describe("put", () => {
    it("should send command with correct bucket and key", async () => {
      mockSend.mockResolvedValue({});
      const driver = createDriver();
      await driver.put("file.xlsx", Buffer.from("data"));

      expect(mockSend).toHaveBeenCalledOnce();
      const cmd = mockSend.mock.calls[0][0];
      expect(cmd.input.Bucket).toBe("mybucket");
      expect(cmd.input.Key).toBe("file.xlsx");
      expect(Buffer.isBuffer(cmd.input.Body)).toBe(true);
    });

    it("should prepend prefix to key", async () => {
      mockSend.mockResolvedValue({});
      const driver = createDriver({ prefix: "exports" });
      await driver.put("file.xlsx", Buffer.from("data"));

      const cmd = mockSend.mock.calls[0][0];
      expect(cmd.input.Key).toBe("exports/file.xlsx");
    });
  });

  describe("get", () => {
    it("should return a Buffer from the response body", async () => {
      const body = (async function* () {
        yield Buffer.from("chunk1");
        yield Buffer.from("chunk2");
      })();
      mockSend.mockResolvedValue({ Body: body });

      const driver = createDriver();
      const result = await driver.get("file.xlsx");
      expect(result.toString()).toBe("chunk1chunk2");
    });

    it("should throw when response body is empty", async () => {
      mockSend.mockResolvedValue({ Body: null });
      const driver = createDriver();
      await expect(driver.get("file.xlsx")).rejects.toThrow("empty body");
    });
  });

  describe("delete", () => {
    it("should send delete command", async () => {
      mockSend.mockResolvedValue({});
      const driver = createDriver();
      await driver.delete("file.xlsx");

      expect(mockSend).toHaveBeenCalledOnce();
      const cmd = mockSend.mock.calls[0][0];
      expect(cmd.input.Bucket).toBe("mybucket");
      expect(cmd.input.Key).toBe("file.xlsx");
    });
  });

  describe("exists", () => {
    it("should return true when HeadObject succeeds", async () => {
      mockSend.mockResolvedValue({});
      const driver = createDriver();
      expect(await driver.exists("file.xlsx")).toBe(true);
    });

    it("should return false on NotFound", async () => {
      mockSend.mockRejectedValue({ name: "NotFound" });
      const driver = createDriver();
      expect(await driver.exists("missing.xlsx")).toBe(false);
    });

    it("should return false on 404 status code", async () => {
      mockSend.mockRejectedValue({ $metadata: { httpStatusCode: 404 } });
      const driver = createDriver();
      expect(await driver.exists("missing.xlsx")).toBe(false);
    });

    it("should rethrow non-404 errors", async () => {
      mockSend.mockRejectedValue(new Error("AccessDenied"));
      const driver = createDriver();
      await expect(driver.exists("file.xlsx")).rejects.toThrow("AccessDenied");
    });
  });
});
