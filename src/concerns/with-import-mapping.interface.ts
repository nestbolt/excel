/**
 * Transform each row after reading during import.
 *
 * Named separately from the export `WithMapping` to allow a single class
 * to implement both import and export mapping.
 */
export interface WithImportMapping<
  TIn = Record<string, any>,
  TOut = Record<string, any>,
> {
  mapRow(row: TIn): TOut;
}
