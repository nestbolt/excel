import type { StorageDriver } from "../storage-driver.interface";
import type { AzureDiskConfig } from "../storage.types";

export class AzureDriver implements StorageDriver {
  private readonly containerClient: any;
  private readonly prefix: string;

  constructor(config: AzureDiskConfig) {
    this.prefix = config.prefix ?? "";

    if (config.client) {
      this.containerClient = config.client;
    } else {
      const sdk = this.loadSdk();

      if (config.connectionString) {
        const blobService =
          sdk.BlobServiceClient.fromConnectionString(config.connectionString);
        this.containerClient = blobService.getContainerClient(config.container);
      } else if (config.accountName && config.accountKey) {
        const credential = new sdk.StorageSharedKeyCredential(
          config.accountName,
          config.accountKey,
        );
        const blobService = new sdk.BlobServiceClient(
          `https://${config.accountName}.blob.core.windows.net`,
          credential,
        );
        this.containerClient = blobService.getContainerClient(config.container);
      } else {
        if (!config.accountName) {
          throw new Error(
            'AzureDriver requires "accountName" when using DefaultAzureCredential.',
          );
        }
        const { DefaultAzureCredential } = this.loadIdentitySdk();
        const blobService = new sdk.BlobServiceClient(
          `https://${config.accountName}.blob.core.windows.net`,
          new DefaultAzureCredential(),
        );
        this.containerClient = blobService.getContainerClient(config.container);
      }
    }
  }

  private loadSdk(): any {
    try {
      return require("@azure/storage-blob");
    } catch /* v8 ignore next 4 */ {
      throw new Error(
        'AzureDriver requires "@azure/storage-blob". Install it: pnpm add @azure/storage-blob',
      );
    }
  }

  private loadIdentitySdk(): any {
    try {
      return require("@azure/identity");
    } catch /* v8 ignore next 4 */ {
      throw new Error(
        'AzureDriver default credential requires "@azure/identity". Install it: pnpm add @azure/identity',
      );
    }
  }

  private key(path: string): string {
    return this.prefix ? `${this.prefix}/${path}` : path;
  }

  async put(path: string, buffer: Buffer): Promise<void> {
    const blobClient = this.containerClient.getBlockBlobClient(this.key(path));
    await blobClient.upload(buffer, buffer.length);
  }

  async get(path: string): Promise<Buffer> {
    const blobClient = this.containerClient.getBlockBlobClient(this.key(path));
    const response = await blobClient.download(0);
    if (!response.readableStreamBody) {
      throw new Error(`Azure returned empty body for blob "${this.key(path)}".`);
    }
    const chunks: Buffer[] = [];
    for await (const chunk of response.readableStreamBody) {
      chunks.push(Buffer.from(chunk));
    }
    return Buffer.concat(chunks);
  }

  async delete(path: string): Promise<void> {
    const blobClient = this.containerClient.getBlockBlobClient(this.key(path));
    try {
      await blobClient.delete();
    } catch (err: any) {
      if (err.statusCode !== 404) throw err;
    }
  }

  async exists(path: string): Promise<boolean> {
    const blobClient = this.containerClient.getBlockBlobClient(this.key(path));
    return blobClient.exists();
  }
}
