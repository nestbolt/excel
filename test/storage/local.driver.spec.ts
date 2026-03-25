import { describe, it, expect, beforeEach, afterEach } from "vitest";
import * as fs from "fs";
import * as os from "os";
import * as path from "path";
import { LocalDriver } from "../../src/storage/drivers/local.driver";

describe("LocalDriver", () => {
  let tmpDir: string;
  let driver: LocalDriver;

  beforeEach(() => {
    tmpDir = fs.mkdtempSync(path.join(os.tmpdir(), "local-driver-"));
    driver = new LocalDriver({ driver: "local", root: tmpDir });
  });

  afterEach(() => {
    fs.rmSync(tmpDir, { recursive: true, force: true });
  });

  describe("put", () => {
    it("should write a file", async () => {
      await driver.put("test.txt", Buffer.from("hello"));
      const content = fs.readFileSync(path.join(tmpDir, "test.txt"), "utf-8");
      expect(content).toBe("hello");
    });

    it("should create parent directories", async () => {
      await driver.put("nested/deep/test.txt", Buffer.from("data"));
      const content = fs.readFileSync(
        path.join(tmpDir, "nested/deep/test.txt"),
        "utf-8",
      );
      expect(content).toBe("data");
    });
  });

  describe("get", () => {
    it("should read a file as Buffer", async () => {
      fs.writeFileSync(path.join(tmpDir, "read.txt"), "content");
      const buf = await driver.get("read.txt");
      expect(buf).toBeInstanceOf(Buffer);
      expect(buf.toString()).toBe("content");
    });

    it("should throw for non-existent file", async () => {
      await expect(driver.get("missing.txt")).rejects.toThrow();
    });
  });

  describe("delete", () => {
    it("should remove a file", async () => {
      const filePath = path.join(tmpDir, "del.txt");
      fs.writeFileSync(filePath, "data");
      await driver.delete("del.txt");
      expect(fs.existsSync(filePath)).toBe(false);
    });

    it("should not throw if file does not exist", async () => {
      await expect(driver.delete("no-such-file.txt")).resolves.not.toThrow();
    });

    it("should rethrow non-ENOENT errors", async () => {
      // Deleting a directory triggers EPERM/EISDIR, not ENOENT
      fs.mkdirSync(path.join(tmpDir, "a-dir"));
      await expect(driver.delete("a-dir")).rejects.toThrow();
    });
  });

  describe("exists", () => {
    it("should return true for existing file", async () => {
      fs.writeFileSync(path.join(tmpDir, "exist.txt"), "data");
      expect(await driver.exists("exist.txt")).toBe(true);
    });

    it("should return false for missing file", async () => {
      expect(await driver.exists("nope.txt")).toBe(false);
    });
  });

  describe("default config", () => {
    it("should use cwd as root when no config provided", () => {
      const defaultDriver = new LocalDriver();
      expect(defaultDriver).toBeDefined();
    });
  });

  describe("path traversal protection", () => {
    it("should reject ../ paths that escape root", () => {
      expect(() => driver["resolve"]("../../etc/passwd")).toThrow(
        "resolves outside the root directory",
      );
    });

    it("should allow absolute paths as-is", async () => {
      const absPath = path.join(tmpDir, "abs-test.txt");
      await driver.put(absPath, Buffer.from("absolute"));
      expect(fs.readFileSync(absPath, "utf-8")).toBe("absolute");
    });

    it("should allow paths within root", async () => {
      await driver.put("safe/nested/file.txt", Buffer.from("ok"));
      expect(await driver.exists("safe/nested/file.txt")).toBe(true);
    });
  });
});
