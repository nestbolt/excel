/**
 * Helper script to create a test template .xlsx file programmatically.
 * Called from tests before template-related specs run.
 */
import { Workbook } from "exceljs";

export async function createTestTemplate(filePath: string): Promise<void> {
  const wb = new Workbook();
  const ws = wb.addWorksheet("Invoice");

  ws.getCell("A1").value = "Company:";
  ws.getCell("B1").value = "{{company}}";
  ws.getCell("A2").value = "Date:";
  ws.getCell("B2").value = "{{date}}";
  ws.getCell("A3").value = "Total:";
  ws.getCell("B3").value = "{{total}}";

  // Header row for data region
  ws.getCell("A5").value = "Item";
  ws.getCell("B5").value = "Qty";
  ws.getCell("C5").value = "Price";

  // Style the header bold
  ws.getRow(5).font = { bold: true };

  await wb.xlsx.writeFile(filePath);
}
