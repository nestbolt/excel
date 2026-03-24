/**
 * Apply styles to rows, columns, or cells.
 *
 * Keys can be:
 * - A row number (`1`, `2`)
 * - A column letter (`'A'`, `'B'`)
 * - A cell reference (`'A1'`, `'B2'`)
 *
 * Values are style objects applied to ExcelJS cells.
 */
export interface WithStyles {
  styles(): Record<string | number, CellStyle>;
}

export interface CellStyle {
  font?: FontStyle;
  alignment?: AlignmentStyle;
  fill?: FillStyle;
  border?: BorderStyles;
  numFmt?: string;
}

export interface FontStyle {
  name?: string;
  size?: number;
  bold?: boolean;
  italic?: boolean;
  underline?: boolean | "single" | "double";
  strike?: boolean;
  color?: string;
}

export interface AlignmentStyle {
  horizontal?: "left" | "center" | "right" | "fill" | "justify";
  vertical?: "top" | "middle" | "bottom";
  wrapText?: boolean;
  textRotation?: number;
}

export interface FillStyle {
  type?: "pattern";
  pattern?: "solid" | "none";
  fgColor?: string;
  bgColor?: string;
}

export interface BorderStyles {
  top?: BorderStyle;
  bottom?: BorderStyle;
  left?: BorderStyle;
  right?: BorderStyle;
}

export interface BorderStyle {
  style?:
    | "thin"
    | "medium"
    | "thick"
    | "dotted"
    | "dashed"
    | "double"
    | "hair"
    | "mediumDashed"
    | "dashDot"
    | "mediumDashDot"
    | "dashDotDot"
    | "mediumDashDotDot"
    | "slantDashDot";
  color?: string;
}
