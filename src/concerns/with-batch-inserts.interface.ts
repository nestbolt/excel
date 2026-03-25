/**
 * Insert imported rows in configurable batch sizes.
 */
export interface WithBatchInserts<T = any> {
  batchSize(): number;
  handleBatch(batch: T[]): void | Promise<void>;
}
