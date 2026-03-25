export interface ExportableOptions {
  /** Sheet tab name. */
  title?: string;

  /** Column widths keyed by column letter. */
  columnWidths?: Record<string, number>;

  /** Auto-filter: `'auto'` to detect from headings, or an explicit range. */
  autoFilter?: string;

  /** Auto-size all columns to fit content. */
  autoSize?: boolean;

  /** Number of rows to freeze from the top. */
  frozenRows?: number;

  /** Number of columns to freeze from the left. */
  frozenColumns?: number;
}

export interface ExportColumnOptions {
  /** Column order (lower numbers appear first). */
  order?: number;

  /** Heading text. Defaults to title-cased property name. */
  header?: string;

  /** Excel number format (e.g. `'#,##0.00'`, `'yyyy-mm-dd'`). */
  format?: string;

  /** Transform function applied to the property value before writing. */
  map?: (value: any, row: any) => any;

  /** Explicit column width in character units. */
  width?: number;
}
