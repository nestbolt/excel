import type { StorageDriver } from "../storage-driver.interface";
import type { S3DiskConfig } from "../storage.types";

export class S3Driver implements StorageDriver {
  private readonly client: any;
  private readonly bucket: string;
  private readonly prefix: string;
  private sdk: any;

  constructor(config: S3DiskConfig) {
    this.bucket = config.bucket;
    this.prefix = config.prefix ?? "";
    this.sdk = this.loadSdk();

    if (config.client) {
      this.client = config.client;
    } else {
      const opts: Record<string, any> = {};
      if (config.region) opts.region = config.region;
      if (config.credentials) opts.credentials = config.credentials;
      if (config.endpoint) {
        opts.endpoint = config.endpoint;
        opts.forcePathStyle = true;
      }
      this.client = new this.sdk.S3Client(opts);
    }
  }

  private loadSdk(): any {
    try {
      return require("@aws-sdk/client-s3");
    } catch /* v8 ignore next 4 */ {
      throw new Error(
        'S3Driver requires "@aws-sdk/client-s3". Install it: pnpm add @aws-sdk/client-s3',
      );
    }
  }

  private key(path: string): string {
    return this.prefix ? `${this.prefix}/${path}` : path;
  }

  async put(path: string, buffer: Buffer): Promise<void> {
    await this.client.send(
      new this.sdk.PutObjectCommand({
        Bucket: this.bucket,
        Key: this.key(path),
        Body: buffer,
      }),
    );
  }

  async get(path: string): Promise<Buffer> {
    const response = await this.client.send(
      new this.sdk.GetObjectCommand({
        Bucket: this.bucket,
        Key: this.key(path),
      }),
    );
    const chunks: Buffer[] = [];
    for await (const chunk of response.Body) {
      chunks.push(Buffer.from(chunk));
    }
    return Buffer.concat(chunks);
  }

  async delete(path: string): Promise<void> {
    await this.client.send(
      new this.sdk.DeleteObjectCommand({
        Bucket: this.bucket,
        Key: this.key(path),
      }),
    );
  }

  async exists(path: string): Promise<boolean> {
    try {
      await this.client.send(
        new this.sdk.HeadObjectCommand({
          Bucket: this.bucket,
          Key: this.key(path),
        }),
      );
      return true;
    } catch (err: any) {
      if (err.name === "NotFound" || err.$metadata?.httpStatusCode === 404) {
        return false;
      }
      throw err;
    }
  }
}
