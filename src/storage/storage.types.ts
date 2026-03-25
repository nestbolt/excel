export interface LocalDiskConfig {
  driver: "local";
  /** Base directory. Paths are resolved relative to this. */
  root?: string;
}

export interface S3DiskConfig {
  driver: "s3";
  bucket: string;
  region?: string;
  /** Prefix prepended to all keys. */
  prefix?: string;
  /** Inline credentials (overrides SDK default chain). */
  credentials?: {
    accessKeyId: string;
    secretAccessKey: string;
    sessionToken?: string;
  };
  /** Endpoint override for S3-compatible services (MinIO, R2, etc.). */
  endpoint?: string;
  /** Pre-configured S3Client instance. */
  client?: any;
}

export interface GCSDiskConfig {
  driver: "gcs";
  bucket: string;
  prefix?: string;
  /** Path to service-account JSON keyfile. */
  keyFilename?: string;
  /** Inline service-account credentials. */
  credentials?: {
    client_email: string;
    private_key: string;
    project_id?: string;
  };
  /** Pre-configured Storage instance. */
  client?: any;
}

export interface AzureDiskConfig {
  driver: "azure";
  container: string;
  prefix?: string;
  /** Connection string. */
  connectionString?: string;
  /** Account name + key auth. */
  accountName?: string;
  accountKey?: string;
  /** Pre-configured ContainerClient instance. */
  client?: any;
}

export type DiskConfig =
  | LocalDiskConfig
  | S3DiskConfig
  | GCSDiskConfig
  | AzureDiskConfig;
