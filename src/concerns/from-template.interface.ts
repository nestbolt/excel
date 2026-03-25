/**
 * Load an existing `.xlsx` template and fill it with data.
 *
 * Placeholders in cells (e.g. `{{company}}`) are replaced with
 * values from `bindings()`. Use `templateData()` to fill repeating
 * rows starting at a given cell.
 */
export interface FromTemplate {
  /** Absolute or relative path to the `.xlsx` template file. */
  templatePath(): string;

  /** Map of placeholder keys to replacement values. */
  bindings(): Record<string, any>;
}

/**
 * Optionally provide repeating row data that is inserted into the
 * template starting at a specific cell.
 */
export interface WithTemplateData {
  /** Cell reference where repeating rows begin (e.g. `'A5'`). */
  dataStartCell(): string;

  /** The rows to insert. Each inner array is one row of cell values. */
  templateData(): any[][] | Promise<any[][]>;
}
