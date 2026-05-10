import ExcelJS from "exceljs";
import { readGuests } from "../ExcelService";

jest.mock("exceljs");

// ─── Helpers ──────────────────────────────────────────────────────────────────

/**
 * Builds a mock ExcelJS.Workbook whose first worksheet will replay the given
 * rows when `eachRow` is called.  Each inner array represents the cell values
 * for that row (index 0 in exceljs is always null, so we prepend it here).
 */
function setupMockWorkbook(rows: (string | null)[][]) {
  const mockEachRow = jest.fn(
    (
      callback: (row: { values: (string | null)[] }, rowIndex: number) => void,
    ) => {
      rows.forEach((rowValues, i) => {
        callback({ values: [null, ...rowValues] }, i + 1);
      });
    },
  );

  const mockWorksheet = { eachRow: mockEachRow };

  const mockWorkbook = {
    csv: { readFile: jest.fn().mockResolvedValue(undefined) },
    xlsx: { readFile: jest.fn().mockResolvedValue(undefined) },
    worksheets: [mockWorksheet],
  };

  (ExcelJS.Workbook as jest.Mock).mockImplementation(() => mockWorkbook);

  return mockWorkbook;
}

/** Minimal valid rows: header + one data row */
function validRows(overrides: Partial<{ name: string; email: string }> = {}) {
  const name = overrides.name ?? "Nguyen Van A";
  const email = overrides.email ?? "nguyenvana@example.com";
  return [
    ["Customer Name", "Email sending"],
    [name, email],
  ];
}

// ─── Tests ────────────────────────────────────────────────────────────────────

describe("ExcelService — readGuests", () => {
  beforeEach(() => jest.resetAllMocks());

  // ── File type routing ───────────────────────────────────────────────────────

  describe("file type routing", () => {
    it("should call csv.readFile for .csv files", async () => {
      const wb = setupMockWorkbook(validRows());
      await readGuests("/data/guests.csv");
      expect(wb.csv.readFile).toHaveBeenCalledWith("/data/guests.csv");
      expect(wb.xlsx.readFile).not.toHaveBeenCalled();
    });

    it("should call xlsx.readFile for .xlsx files", async () => {
      const wb = setupMockWorkbook(validRows());
      await readGuests("/data/guests.xlsx");
      expect(wb.xlsx.readFile).toHaveBeenCalledWith("/data/guests.xlsx");
      expect(wb.csv.readFile).not.toHaveBeenCalled();
    });

    it("should call xlsx.readFile for .xls files", async () => {
      const wb = setupMockWorkbook(validRows());
      await readGuests("/data/guests.xls");
      expect(wb.xlsx.readFile).toHaveBeenCalled();
    });
  });

  // ── Parsing — happy path ────────────────────────────────────────────────────

  describe("parsing — happy path", () => {
    it("should return a Guest with name and email from a CSV", async () => {
      setupMockWorkbook(
        validRows({ name: "Tran Thi B", email: "tranthib@example.com" }),
      );
      const guests = await readGuests("guests.csv");
      expect(guests).toHaveLength(1);
      expect(guests[0]).toEqual({
        name: "Tran Thi B",
        email: "tranthib@example.com",
      });
    });

    it("should parse columns regardless of their position in the header", async () => {
      // Email comes before Customer Name
      setupMockWorkbook([
        ["Region", "Email sending", "Customer Name", "Phone"],
        ["South", "le@example.com", "Le Van C", "0901"],
      ]);
      const guests = await readGuests("guests.csv");
      expect(guests[0]).toEqual({ name: "Le Van C", email: "le@example.com" });
    });

    it("should trim whitespace from name and email values", async () => {
      setupMockWorkbook([
        ["Customer Name", "Email sending"],
        ["  Nguyen Van D  ", "  nguyenvand@example.com  "],
      ]);
      const guests = await readGuests("guests.csv");
      expect(guests[0]).toEqual({
        name: "Nguyen Van D",
        email: "nguyenvand@example.com",
      });
    });

    it("should return all guests from multiple data rows", async () => {
      const dataRows = Array.from({ length: 10 }, (_, i) => [
        `Guest ${i + 1}`,
        `guest${i + 1}@example.com`,
      ]);
      setupMockWorkbook([["Customer Name", "Email sending"], ...dataRows]);

      const guests = await readGuests("guests.csv");
      expect(guests).toHaveLength(10);
      expect(guests[4]).toEqual({
        name: "Guest 5",
        email: "guest5@example.com",
      });
    });
  });

  // ── Skipping invalid rows ───────────────────────────────────────────────────

  describe("skipping invalid rows", () => {
    it("should skip rows where Customer Name is empty", async () => {
      setupMockWorkbook([
        ["Customer Name", "Email sending"],
        ["", "valid@example.com"],
        ["Valid Guest", "valid@example.com"],
      ]);
      const guests = await readGuests("guests.csv");
      expect(guests).toHaveLength(1);
      expect(guests[0]!.name).toBe("Valid Guest");
    });

    it("should skip rows where Email sending is empty", async () => {
      setupMockWorkbook([
        ["Customer Name", "Email sending"],
        ["No Email Guest", ""],
        ["Has Email Guest", "hasemail@example.com"],
      ]);
      const guests = await readGuests("guests.csv");
      expect(guests).toHaveLength(1);
      expect(guests[0]!.email).toBe("hasemail@example.com");
    });

    it("should skip rows where both name and email are empty", async () => {
      setupMockWorkbook([
        ["Customer Name", "Email sending"],
        ["", ""],
        ["Real Guest", "real@example.com"],
      ]);
      const guests = await readGuests("guests.csv");
      expect(guests).toHaveLength(1);
    });

    it("should skip rows with null cell values", async () => {
      setupMockWorkbook([
        ["Customer Name", "Email sending"],
        [null, null],
        ["Null Survivor", "survivor@example.com"],
      ]);
      const guests = await readGuests("guests.csv");
      expect(guests).toHaveLength(1);
    });

    it("should return an empty array when all data rows are invalid", async () => {
      setupMockWorkbook([
        ["Customer Name", "Email sending"],
        ["", ""],
        [null, null],
      ]);
      const guests = await readGuests("guests.csv");
      expect(guests).toHaveLength(0);
    });
  });

  // ── Error cases ─────────────────────────────────────────────────────────────

  describe("error cases", () => {
    it("should throw when the workbook has no worksheets", async () => {
      (ExcelJS.Workbook as jest.Mock).mockImplementation(() => ({
        csv: { readFile: jest.fn().mockResolvedValue(undefined) },
        xlsx: { readFile: jest.fn().mockResolvedValue(undefined) },
        worksheets: [],
      }));

      await expect(readGuests("guests.csv")).rejects.toThrow(
        "No worksheet found",
      );
    });

    it("should return empty array when header is missing required columns", async () => {
      setupMockWorkbook([
        ["Item", "Region", "Phone"], // No Customer Name / Email sending
        ["product1", "South", "0901"],
      ]);
      const guests = await readGuests("guests.csv");
      expect(guests).toHaveLength(0);
    });
  });

  // ── Scale test ──────────────────────────────────────────────────────────────

  describe("scale", () => {
    it("should correctly parse 450 guest rows", async () => {
      const dataRows = Array.from({ length: 450 }, (_, i) => [
        `Guest ${i + 1}`,
        `guest${i + 1}@company.com`,
      ]);
      setupMockWorkbook([["Customer Name", "Email sending"], ...dataRows]);

      const guests = await readGuests("guests.csv");

      expect(guests).toHaveLength(450);
      expect(guests[0]).toEqual({
        name: "Guest 1",
        email: "guest1@company.com",
      });
      expect(guests[449]).toEqual({
        name: "Guest 450",
        email: "guest450@company.com",
      });
    });

    it("should skip invalid rows and still return valid guests within 450 rows", async () => {
      // Every 10th row is invalid
      const dataRows = Array.from({ length: 450 }, (_, i) =>
        (i + 1) % 10 === 0
          ? ["", ""]
          : [`Guest ${i + 1}`, `guest${i + 1}@company.com`],
      );
      setupMockWorkbook([["Customer Name", "Email sending"], ...dataRows]);

      const guests = await readGuests("guests.csv");

      // 45 rows are invalid (rows 10, 20, ..., 450)
      expect(guests).toHaveLength(405);
    });
  });
});
