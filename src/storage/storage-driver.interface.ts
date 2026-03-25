/**
 * Contract for all storage backends.
 *
 * The meaning of `path` depends on the driver:
 *   - LocalDriver  — filesystem path (absolute or relative to `root`)
 *   - S3Driver     — S3 object key (bucket configured at driver level)
 *   - GCSDriver    — GCS object name (bucket configured at driver level)
 *   - AzureDriver  — blob name (container configured at driver level)
 */
export interface StorageDriver {
  /** Write a buffer to the given path. */
  put(path: string, buffer: Buffer): Promise<void>;

  /** Read contents at the given path and return as Buffer. */
  get(path: string): Promise<Buffer>;

  /** Delete the object at the given path (no-op if missing). */
  delete(path: string): Promise<void>;

  /** Check whether an object exists at the given path. */
  exists(path: string): Promise<boolean>;
}
