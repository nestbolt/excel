import { ExcelType, EXTENSION_MAP } from "../excel.constants";

/**
 * Detect the export type from a filename extension.
 *
 * Falls back to `fallback` (default XLSX) when the extension is
 * unrecognised.
 */
export function detectType(
  filename: string,
  fallback: ExcelType = ExcelType.XLSX,
): ExcelType {
  const ext = filename.split(".").pop()?.toLowerCase();
  if (ext && ext in EXTENSION_MAP) {
    return EXTENSION_MAP[ext];
  }
  return fallback;
}

/**
 * Parse a cell reference like `'C3'` into zero-based column and
 * one-based row indices.
 */
export function parseCellRef(ref: string): { col: number; row: number } {
  const match = ref.match(/^([A-Z]+)(\d+)$/i);
  if (!match) {
    throw new Error(`Invalid cell reference: "${ref}"`);
  }
  const colStr = match[1].toUpperCase();
  const row = parseInt(match[2], 10);

  let col = 0;
  for (let i = 0; i < colStr.length; i++) {
    col = col * 26 + (colStr.charCodeAt(i) - 64);
  }

  return { col, row };
}

/**
 * Convert a column letter (`'A'`, `'AB'`) to a 1-based column number.
 */
export function columnLetterToNumber(letter: string): number {
  const upper = letter.toUpperCase();
  let num = 0;
  for (let i = 0; i < upper.length; i++) {
    num = num * 26 + (upper.charCodeAt(i) - 64);
  }
  return num;
}

/**
 * Convert a 1-based column number to a column letter (`1` → `'A'`,
 * `27` → `'AA'`).
 */
export function numberToColumnLetter(num: number): string {
  let result = "";
  let n = num;
  while (n > 0) {
    const remainder = (n - 1) % 26;
    result = String.fromCharCode(65 + remainder) + result;
    n = Math.floor((n - 1) / 26);
  }
  return result;
}
