import type { Worksheet } from "exceljs";
import type {
  FromCollection,
  FromArray,
  WithHeadings,
  WithMapping,
  WithTitle,
  WithColumnWidths,
  WithColumnFormatting,
  WithStyles,
  WithCustomStartCell,
  WithAutoFilter,
  WithFrozenRows,
  WithFrozenColumns,
  CellStyle,
  FontStyle,
  FillStyle,
  BorderStyle as IBorderStyle,
} from "./concerns";
import { columnLetterToNumber, numberToColumnLetter, parseCellRef } from "./helpers";

/* ------------------------------------------------------------------ */
/*  Type-guard helpers                                                 */
/* ------------------------------------------------------------------ */

function isFromCollection(obj: any): obj is FromCollection {
  return typeof obj.collection === "function";
}

function isFromArray(obj: any): obj is FromArray {
  return typeof obj.array === "function";
}

function isWithHeadings(obj: any): obj is WithHeadings {
  return typeof obj.headings === "function";
}

function isWithMapping(obj: any): obj is WithMapping {
  return typeof obj.map === "function";
}

function isWithTitle(obj: any): obj is WithTitle {
  return typeof obj.title === "function";
}

function isWithColumnWidths(obj: any): obj is WithColumnWidths {
  return typeof obj.columnWidths === "function";
}

function isWithColumnFormatting(obj: any): obj is WithColumnFormatting {
  return typeof obj.columnFormats === "function";
}

function isWithStyles(obj: any): obj is WithStyles {
  return typeof obj.styles === "function";
}

function isShouldAutoSize(obj: any): boolean {
  return obj.shouldAutoSize === true;
}

function isWithCustomStartCell(obj: any): obj is WithCustomStartCell {
  return typeof obj.startCell === "function";
}

function isWithAutoFilter(obj: any): obj is WithAutoFilter {
  return typeof obj.autoFilter === "function";
}

function isWithFrozenRows(obj: any): obj is WithFrozenRows {
  return typeof obj.frozenRows === "function";
}

function isWithFrozenColumns(obj: any): obj is WithFrozenColumns {
  return typeof obj.frozenColumns === "function";
}

/* ------------------------------------------------------------------ */
/*  Style conversion helpers                                           */
/* ------------------------------------------------------------------ */

function toArgb(hex: string): string {
  const clean = hex.replace(/^#/, "");
  return clean.length === 6 ? `FF${clean}` : clean;
}

function convertFont(f: FontStyle): Record<string, any> {
  const out: Record<string, any> = {};
  if (f.name !== undefined) out.name = f.name;
  if (f.size !== undefined) out.size = f.size;
  if (f.bold !== undefined) out.bold = f.bold;
  if (f.italic !== undefined) out.italic = f.italic;
  if (f.underline !== undefined) out.underline = f.underline;
  if (f.strike !== undefined) out.strike = f.strike;
  if (f.color !== undefined) out.color = { argb: toArgb(f.color) };
  return out;
}

function convertFill(f: FillStyle): Record<string, any> {
  return {
    type: f.type ?? "pattern",
    pattern: f.pattern ?? "solid",
    fgColor: f.fgColor ? { argb: toArgb(f.fgColor) } : undefined,
    bgColor: f.bgColor ? { argb: toArgb(f.bgColor) } : undefined,
  };
}

function convertBorderSide(
  b: IBorderStyle,
): Record<string, any> {
  const out: Record<string, any> = {};
  if (b.style) out.style = b.style;
  if (b.color) out.color = { argb: toArgb(b.color) };
  return out;
}

function convertStyle(style: CellStyle): Record<string, any> {
  const out: Record<string, any> = {};
  if (style.font) out.font = convertFont(style.font);
  if (style.alignment) out.alignment = style.alignment;
  if (style.fill) out.fill = convertFill(style.fill);
  if (style.border) {
    const border: Record<string, any> = {};
    if (style.border.top) border.top = convertBorderSide(style.border.top);
    if (style.border.bottom)
      border.bottom = convertBorderSide(style.border.bottom);
    if (style.border.left) border.left = convertBorderSide(style.border.left);
    if (style.border.right)
      border.right = convertBorderSide(style.border.right);
    out.border = border;
  }
  if (style.numFmt) out.numFmt = style.numFmt;
  return out;
}

/* ------------------------------------------------------------------ */
/*  Sheet builder                                                      */
/* ------------------------------------------------------------------ */

export async function populateSheet(
  worksheet: Worksheet,
  exportable: object,
): Promise<void> {
  // --- title --------------------------------------------------------
  if (isWithTitle(exportable)) {
    worksheet.name = exportable.title();
  }

  // --- determine start position ------------------------------------
  let startRow = 1;
  let startCol = 1;
  if (isWithCustomStartCell(exportable)) {
    const ref = parseCellRef(exportable.startCell());
    startRow = ref.row;
    startCol = ref.col;
  }

  let currentRow = startRow;
  let headingColCount = 0;
  let headingStartRow = startRow;

  // --- headings -----------------------------------------------------
  if (isWithHeadings(exportable)) {
    const headings = exportable.headings();
    const headingRows = Array.isArray(headings[0]) ? headings : [headings];
    headingStartRow = currentRow;

    for (const headingRow of headingRows as string[][]) {
      if (headingRow.length > headingColCount) {
        headingColCount = headingRow.length;
      }
      const row = worksheet.getRow(currentRow);
      headingRow.forEach((val, idx) => {
        row.getCell(startCol + idx).value = val;
      });
      row.commit();
      currentRow++;
    }
  }

  // --- data rows ----------------------------------------------------
  let rows: any[][];

  if (isFromCollection(exportable)) {
    const data = await exportable.collection();

    if (isWithMapping(exportable)) {
      rows = data.map((item) => (exportable as WithMapping).map(item));
    } else {
      rows = data.map((item) => {
        if (Array.isArray(item)) return item;
        return Object.values(item);
      });
    }
  } else if (isFromArray(exportable)) {
    rows = await exportable.array();
  } else {
    throw new Error(
      "Export must implement FromCollection or FromArray to provide data.",
    );
  }

  for (const rowData of rows) {
    const row = worksheet.getRow(currentRow);
    rowData.forEach((val, idx) => {
      row.getCell(startCol + idx).value = val;
    });
    row.commit();
    currentRow++;
  }

  // --- column widths ------------------------------------------------
  if (isWithColumnWidths(exportable)) {
    const widths = exportable.columnWidths();
    for (const [letter, width] of Object.entries(widths)) {
      const colNum = columnLetterToNumber(letter);
      worksheet.getColumn(colNum).width = width;
    }
  }

  // --- column formatting --------------------------------------------
  if (isWithColumnFormatting(exportable)) {
    const formats = exportable.columnFormats();
    for (const [letter, numFmt] of Object.entries(formats)) {
      const colNum = columnLetterToNumber(letter);
      worksheet.getColumn(colNum).numFmt = numFmt;
    }
  }

  // --- auto-size ----------------------------------------------------
  if (isShouldAutoSize(exportable)) {
    worksheet.columns.forEach((column) => {
      if (!column?.eachCell) return;
      let maxLen = 10;
      column.eachCell({ includeEmpty: false }, (cell) => {
        const val = cell.value;
        const len =
          val !== null && val !== undefined ? String(val).length : 0;
        if (len > maxLen) maxLen = len;
      });
      column.width = Math.min(maxLen + 2, 60);
    });
  }

  // --- styles -------------------------------------------------------
  if (isWithStyles(exportable)) {
    const styleMap = exportable.styles();

    for (const [key, style] of Object.entries(styleMap)) {
      const converted = convertStyle(style);
      const numKey = Number(key);

      if (!isNaN(numKey)) {
        // Row number
        const row = worksheet.getRow(numKey);
        row.eachCell({ includeEmpty: false }, (cell) => {
          Object.assign(cell, converted);
        });
      } else if (/^[A-Z]+$/i.test(key)) {
        // Column letter
        const colNum = columnLetterToNumber(key);
        worksheet.getColumn(colNum).eachCell({ includeEmpty: false }, (cell) => {
          Object.assign(cell, converted);
        });
      } else if (/^[A-Z]+\d+$/i.test(key)) {
        // Cell reference
        const cell = worksheet.getCell(key);
        Object.assign(cell, converted);
      }
    }
  }

  // --- auto-filter --------------------------------------------------
  if (isWithAutoFilter(exportable)) {
    const filterValue = exportable.autoFilter();
    if (filterValue === "auto") {
      if (headingColCount > 0) {
        const lastColLetter = numberToColumnLetter(startCol + headingColCount - 1);
        const firstColLetter = numberToColumnLetter(startCol);
        worksheet.autoFilter = `${firstColLetter}${headingStartRow}:${lastColLetter}${headingStartRow}`;
      }
    } else {
      worksheet.autoFilter = filterValue;
    }
  }

  // --- frozen rows / columns ----------------------------------------
  const frozenRowCount = isWithFrozenRows(exportable)
    ? exportable.frozenRows()
    : 0;
  const frozenColCount = isWithFrozenColumns(exportable)
    ? exportable.frozenColumns()
    : 0;

  if (frozenRowCount > 0 || frozenColCount > 0) {
    worksheet.views = [
      {
        state: "frozen",
        xSplit: frozenColCount,
        ySplit: frozenRowCount,
      },
    ];
  }
}
