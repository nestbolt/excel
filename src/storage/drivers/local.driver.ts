import * as fs from "fs/promises";
import * as path from "path";
import type { StorageDriver } from "../storage-driver.interface";
import type { LocalDiskConfig } from "../storage.types";

export class LocalDriver implements StorageDriver {
  private readonly root: string;

  constructor(config: LocalDiskConfig = { driver: "local" }) {
    this.root = path.resolve(config.root ?? process.cwd());
  }

  private resolve(filePath: string): string {
    if (path.isAbsolute(filePath)) return filePath;
    const resolved = path.resolve(this.root, filePath);
    if (!resolved.startsWith(this.root + path.sep) && resolved !== this.root) {
      throw new Error(
        `Path "${filePath}" resolves outside the root directory.`,
      );
    }
    return resolved;
  }

  async put(filePath: string, buffer: Buffer): Promise<void> {
    const resolved = this.resolve(filePath);
    await fs.mkdir(path.dirname(resolved), { recursive: true });
    await fs.writeFile(resolved, buffer);
  }

  async get(filePath: string): Promise<Buffer> {
    return fs.readFile(this.resolve(filePath));
  }

  async delete(filePath: string): Promise<void> {
    try {
      await fs.unlink(this.resolve(filePath));
    } catch (err: any) {
      if (err.code !== "ENOENT") throw err;
    }
  }

  async exists(filePath: string): Promise<boolean> {
    try {
      await fs.access(this.resolve(filePath));
      return true;
    } catch {
      return false;
    }
  }
}
