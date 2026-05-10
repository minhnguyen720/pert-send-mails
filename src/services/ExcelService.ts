import ExcelJS from "exceljs";
import path from "path";
import type { Guest } from "../types";

export async function readGuests(filePath: string): Promise<Guest[]> {
  const workbook = new ExcelJS.Workbook();
  const ext = path.extname(filePath).toLowerCase();

  if (ext === ".csv") {
    await workbook.csv.readFile(filePath);
  } else {
    await workbook.xlsx.readFile(filePath);
  }

  const worksheet = workbook.worksheets[0];
  if (!worksheet) {
    throw new Error(`No worksheet found in file: ${filePath}`);
  }

  const guests: Guest[] = [];
  let headerRow: string[] = [];

  worksheet.eachRow((row, rowIndex) => {
    const values = (row.values as ExcelJS.CellValue[]).slice(1); // index 0 is always null in exceljs

    if (rowIndex === 1) {
      headerRow = values.map((v) => String(v ?? "").trim());
      return;
    }

    const nameIdx = headerRow.indexOf("Customer Name");
    const emailIdx = headerRow.indexOf("Email sending");

    if (nameIdx === -1 || emailIdx === -1) return;

    const name = String(values[nameIdx] ?? "").trim();
    const email = String(values[emailIdx] ?? "").trim();

    if (!name || !email) {
      console.warn(
        `[ExcelService] Skipping row ${rowIndex}: missing name or email.`,
      );
      return;
    }

    guests.push({ name, email });
  });

  return guests;
}
