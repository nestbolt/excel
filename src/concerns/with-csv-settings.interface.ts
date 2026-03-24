/**
 * Override CSV formatting options for this export.
 */
export interface WithCsvSettings {
  csvSettings(): CsvSettings;
}

export interface CsvSettings {
  /** Column delimiter (default `','`) */
  delimiter?: string;
  /** Quote character (default `'"'`) */
  quoteChar?: string;
  /** Line ending (default `'\n'`) */
  lineEnding?: string;
  /** Prepend a UTF-8 BOM (default `false`) */
  useBom?: boolean;
  /** Output encoding (default `'utf-8'`) */
  encoding?: BufferEncoding;
}
