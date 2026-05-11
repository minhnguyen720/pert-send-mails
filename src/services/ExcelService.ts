import ExcelJS from "exceljs";
import path from "path";
import type { Guest } from "../types";

const excludeList = [
  80, 241, 242, 246, 247, 248, 249, 250, 257, 258, 259, 260, 261, 262, 268, 269,
  270, 271, 272, 273,
];

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

    const itemIdx = headerRow.indexOf("Item");
    const nameIdx = headerRow.indexOf("Customer Name");
    const emailIdx = headerRow.indexOf("Email sending");
    const ccIdx = headerRow.indexOf("CC email");

    if (nameIdx === -1 || emailIdx === -1) return;

    const name = String(values[nameIdx] ?? "").trim();
    const email = String(values[emailIdx] ?? "")
      .trim()
      .toLowerCase();
    const rawItem = itemIdx !== -1 ? values[itemIdx] : null;
    const item = rawItem != null && rawItem !== "" ? Number(rawItem) : NaN;
    const ccRaw = ccIdx !== -1 ? String(values[ccIdx] ?? "") : "";
    const ccEmails = ccRaw
      .split("/")
      .map((e) => e.trim().toLowerCase())
      .filter((e) => e.length > 0);

    if (!name || !email) {
      console.warn(
        `[ExcelService] Skipping row ${rowIndex}: missing name or email.`,
      );
      return;
    }

    if (isNaN(item)) {
      console.warn(
        `[ExcelService] Skipping row ${rowIndex}: missing or invalid Item value.`,
      );
      return;
    }

    if (!excludeList.includes(item)) {
      guests.push({ item, name, email, ccEmails });
    }
  });

  return guests;
}
