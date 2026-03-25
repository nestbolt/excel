import * as fs from "fs";
import * as path from "path";
import type { StorageDriver } from "../storage-driver.interface";
import type { LocalDiskConfig } from "../storage.types";

export class LocalDriver implements StorageDriver {
  private readonly root: string;

  constructor(config: LocalDiskConfig = { driver: "local" }) {
    this.root = config.root ?? process.cwd();
  }

  private resolve(filePath: string): string {
    return path.isAbsolute(filePath)
      ? filePath
      : path.resolve(this.root, filePath);
  }

  async put(filePath: string, buffer: Buffer): Promise<void> {
    const resolved = this.resolve(filePath);
    const dir = path.dirname(resolved);
    if (!fs.existsSync(dir)) {
      fs.mkdirSync(dir, { recursive: true });
    }
    fs.writeFileSync(resolved, buffer);
  }

  async get(filePath: string): Promise<Buffer> {
    return fs.readFileSync(this.resolve(filePath));
  }

  async delete(filePath: string): Promise<void> {
    const resolved = this.resolve(filePath);
    if (fs.existsSync(resolved)) {
      fs.unlinkSync(resolved);
    }
  }

  async exists(filePath: string): Promise<boolean> {
    return fs.existsSync(this.resolve(filePath));
  }
}
