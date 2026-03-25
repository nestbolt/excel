import type { StorageDriver } from "../storage-driver.interface";
import type { GCSDiskConfig } from "../storage.types";

export class GCSDriver implements StorageDriver {
  private readonly bucket: any;
  private readonly prefix: string;

  constructor(config: GCSDiskConfig) {
    this.prefix = config.prefix ?? "";

    if (config.client) {
      this.bucket = config.client.bucket(config.bucket);
    } else {
      const { Storage } = this.loadSdk();
      const opts: Record<string, any> = {};
      if (config.keyFilename) opts.keyFilename = config.keyFilename;
      if (config.credentials) opts.credentials = config.credentials;
      const storage = new Storage(opts);
      this.bucket = storage.bucket(config.bucket);
    }
  }

  private loadSdk(): any {
    try {
      return require("@google-cloud/storage");
    } catch /* v8 ignore next 4 */ {
      throw new Error(
        'GCSDriver requires "@google-cloud/storage". Install it: pnpm add @google-cloud/storage',
      );
    }
  }

  private key(path: string): string {
    return this.prefix ? `${this.prefix}/${path}` : path;
  }

  async put(path: string, buffer: Buffer): Promise<void> {
    await this.bucket.file(this.key(path)).save(buffer);
  }

  async get(path: string): Promise<Buffer> {
    const [contents] = await this.bucket.file(this.key(path)).download();
    return contents;
  }

  async delete(path: string): Promise<void> {
    try {
      await this.bucket.file(this.key(path)).delete();
    } catch (err: any) {
      if (err.code !== 404) throw err;
    }
  }

  async exists(path: string): Promise<boolean> {
    const [exists] = await this.bucket.file(this.key(path)).exists();
    return exists;
  }
}
