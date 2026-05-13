import ExcelJS from "exceljs";
import { readGuests } from "../ExcelService";

jest.mock("exceljs");

// ─── Helpers ──────────────────────────────────────────────────────────────────

/**
 * Builds a mock ExcelJS.Workbook whose first worksheet will replay the given
 * rows when `eachRow` is called.  Each inner array represents the cell values
 * for that row (index 0 in exceljs is always null, so we prepend it here).
 */
function setupMockWorkbook(rows: (string | number | null)[][]) {
  const mockEachRow = jest.fn(
    (
      callback: (
        row: { values: (string | number | null)[] },
        rowIndex: number,
      ) => void,
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

/** Standard header row matching the real CSV structure */
const HEADER = ["Item", "Customer Name", "Email sending", "CC email"];

/** Minimal valid rows: header + one data row */
function validRows(
  overrides: Partial<{
    item: number;
    name: string;
    email: string;
    cc: string;
  }> = {},
) {
  const item = overrides.item ?? 1;
  const name = overrides.name ?? "Nguyen Van A";
  const email = overrides.email ?? "nguyenvana@example.com";
  const cc = overrides.cc ?? "";
  return [HEADER, [item, name, email, cc]];
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
    it("should return a Guest with item, name and email from a CSV", async () => {
      setupMockWorkbook(
        validRows({
          item: 42,
          name: "Tran Thi B",
          email: "tranthib@example.com",
        }),
      );
      const guests = await readGuests("guests.csv");
      expect(guests).toHaveLength(1);
      expect(guests[0]).toEqual({
        item: 42,
        name: "Tran Thi B",
        email: "tranthib@example.com",
        ccEmails: [],
      });
    });

    it("should parse columns regardless of their position in the header", async () => {
      setupMockWorkbook([
        ["Region", "Email sending", "Customer Name", "Item", "Phone"],
        ["South", "le@example.com", "Le Van C", 7, "0901"],
      ]);
      const guests = await readGuests("guests.csv");
      expect(guests[0]).toEqual({
        item: 7,
        name: "Le Van C",
        email: "le@example.com",
        ccEmails: [],
      });
    });

    it("should trim whitespace from name and email values", async () => {
      setupMockWorkbook([
        HEADER,
        [5, "  Nguyen Van D  ", "  nguyenvand@example.com  ", ""],
      ]);
      const guests = await readGuests("guests.csv");
      expect(guests[0]).toEqual({
        item: 5,
        name: "Nguyen Van D",
        email: "nguyenvand@example.com",
        ccEmails: [],
      });
    });

    it("should return all guests from multiple data rows", async () => {
      const dataRows = Array.from({ length: 10 }, (_, i) => [
        i + 1,
        `Guest ${i + 1}`,
        `guest${i + 1}@example.com`,
      ]);
      setupMockWorkbook([HEADER, ...dataRows]);

      const guests = await readGuests("guests.csv");
      expect(guests).toHaveLength(10);
      expect(guests[4]).toEqual({
        item: 5,
        name: "Guest 5",
        email: "guest5@example.com",
        ccEmails: [],
      });
    });
  });

  // ── Item column — attachment filename pattern ────────────────────────────────

  describe("Item column — filename pattern", () => {
    it("should parse item 1 correctly (used to build 001000-01 filename)", async () => {
      setupMockWorkbook(validRows({ item: 1 }));
      const guests = await readGuests("guests.csv");
      expect(guests[0]!.item).toBe(1);
      expect(String(guests[0]!.item).padStart(3, "0")).toBe("001");
    });

    it("should parse item 83 correctly (used to build 083000-01 filename)", async () => {
      setupMockWorkbook(validRows({ item: 83 }));
      const guests = await readGuests("guests.csv");
      expect(guests[0]!.item).toBe(83);
      expect(String(guests[0]!.item).padStart(3, "0")).toBe("083");
    });

    it("should parse item 323 correctly (used to build 323000-01 filename)", async () => {
      setupMockWorkbook(validRows({ item: 323 }));
      const guests = await readGuests("guests.csv");
      expect(guests[0]!.item).toBe(323);
      expect(String(guests[0]!.item).padStart(3, "0")).toBe("323");
    });

    it("should produce the correct full filename for item 1", async () => {
      setupMockWorkbook(validRows({ item: 1 }));
      const guests = await readGuests("guests.csv");
      const itemPadded = String(guests[0]!.item).padStart(3, "0");
      expect(`E-Invitation_PARTNER 1_${itemPadded}000-01.jpg`).toBe(
        "E-Invitation_PARTNER 1_001000-01.jpg",
      );
    });

    it("should produce the correct full filename for item 83", async () => {
      setupMockWorkbook(validRows({ item: 83 }));
      const guests = await readGuests("guests.csv");
      const itemPadded = String(guests[0]!.item).padStart(3, "0");
      expect(`E-Invitation_PARTNER 1_${itemPadded}000-01.jpg`).toBe(
        "E-Invitation_PARTNER 1_083000-01.jpg",
      );
    });

    it("should produce the correct full filename for item 323", async () => {
      setupMockWorkbook(validRows({ item: 323 }));
      const guests = await readGuests("guests.csv");
      const itemPadded = String(guests[0]!.item).padStart(3, "0");
      expect(`E-Invitation_PARTNER 1_${itemPadded}000-01.jpg`).toBe(
        "E-Invitation_PARTNER 1_323000-01.jpg",
      );
    });
  });

  // ── Skipping invalid rows ───────────────────────────────────────────────────

  describe("skipping invalid rows", () => {
    it("should skip rows where Customer Name is empty", async () => {
      setupMockWorkbook([
        HEADER,
        [1, "", "valid@example.com"],
        [2, "Valid Guest", "valid@example.com"],
      ]);
      const guests = await readGuests("guests.csv");
      expect(guests).toHaveLength(1);
      expect(guests[0]!.name).toBe("Valid Guest");
    });

    it("should skip rows where Email sending is empty", async () => {
      setupMockWorkbook([
        HEADER,
        [1, "No Email Guest", ""],
        [2, "Has Email Guest", "hasemail@example.com"],
      ]);
      const guests = await readGuests("guests.csv");
      expect(guests).toHaveLength(1);
      expect(guests[0]!.email).toBe("hasemail@example.com");
    });

    it("should skip rows where Item is missing (null)", async () => {
      setupMockWorkbook([
        HEADER,
        [null, "No Item Guest", "noitem@example.com", ""],
        [5, "Has Item Guest", "hasitem@example.com", ""],
      ]);
      const guests = await readGuests("guests.csv");
      expect(guests).toHaveLength(1);
      expect(guests[0]!.item).toBe(5);
    });

    it("should skip rows where Item is not a number (NaN)", async () => {
      setupMockWorkbook([
        HEADER,
        ["invalid", "Bad Item Guest", "bad@example.com", ""],
        [10, "Good Item Guest", "good@example.com", ""],
      ]);
      const guests = await readGuests("guests.csv");
      expect(guests).toHaveLength(1);
      expect(guests[0]!.item).toBe(10);
    });

    it("should skip rows where both name and email are empty", async () => {
      setupMockWorkbook([
        HEADER,
        [1, "", ""],
        [2, "Real Guest", "real@example.com"],
      ]);
      const guests = await readGuests("guests.csv");
      expect(guests).toHaveLength(1);
    });

    it("should skip rows with null cell values", async () => {
      setupMockWorkbook([
        HEADER,
        [null, null, null],
        [3, "Null Survivor", "survivor@example.com"],
      ]);
      const guests = await readGuests("guests.csv");
      expect(guests).toHaveLength(1);
    });

    it("should return an empty array when all data rows are invalid", async () => {
      setupMockWorkbook([HEADER, [null, "", ""], [null, null, null]]);
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
        ["Region", "Phone"], // No Item / Customer Name / Email sending
        ["South", "0901"],
      ]);
      const guests = await readGuests("guests.csv");
      expect(guests).toHaveLength(0);
    });
  });

  // ── Scale test ──────────────────────────────────────────────────────────────

  describe("scale", () => {
    it("should correctly parse 450 guest rows with item numbers", async () => {
      const dataRows = Array.from({ length: 450 }, (_, i) => [
        i + 1,
        `Guest ${i + 1}`,
        `guest${i + 1}@company.com`,
      ]);
      setupMockWorkbook([HEADER, ...dataRows]);

      const guests = await readGuests("guests.csv");

      // excludeList=[80,241,242] removes 3 of the 450 rows
      expect(guests).toHaveLength(447);
      expect(guests[0]).toEqual({
        item: 1,
        name: "Guest 1",
        email: "guest1@company.com",
        ccEmails: [],
      });
      // guest at index 446 is item 450 (last valid after 3 excluded)
      expect(guests[446]).toEqual({
        item: 450,
        name: "Guest 450",
        email: "guest450@company.com",
        ccEmails: [],
      });
    });

    it("should skip invalid rows and still return valid guests within 450 rows", async () => {
      // Every 10th row has an empty string item — invalid (null → 0, so use "" to trigger NaN)
      const dataRows = Array.from({ length: 450 }, (_, i) =>
        (i + 1) % 10 === 0
          ? ["", `Guest ${i + 1}`, `guest${i + 1}@company.com`]
          : [i + 1, `Guest ${i + 1}`, `guest${i + 1}@company.com`],
      );
      setupMockWorkbook([HEADER, ...dataRows]);

      const guests = await readGuests("guests.csv");

      // 45 rows have empty item (rows 10,20,...,450) — item 80 is one of them (row 80 → index 79, (79+1)%10=0)
      // items 241 and 242 are valid rows but filtered by excludeList → 450 - 45 - 2 = 403
      expect(guests).toHaveLength(403);
    });

    it("should produce correct padded filenames for all 450 item numbers", async () => {
      const dataRows = Array.from({ length: 450 }, (_, i) => [
        i + 1,
        `Guest ${i + 1}`,
        `guest${i + 1}@company.com`,
      ]);
      setupMockWorkbook([HEADER, ...dataRows]);

      const guests = await readGuests("guests.csv");

      // Spot-check filename formula for item 1, 83, 323, 450
      const check = (item: number, expected: string) => {
        const guest = guests.find((g) => g.item === item)!;
        const padded = String(guest.item).padStart(3, "0");
        expect(`E-Invitation_PARTNER 1_${padded}000-01.jpg`).toBe(expected);
      };

      check(1, "E-Invitation_PARTNER 1_001000-01.jpg");
      check(83, "E-Invitation_PARTNER 1_083000-01.jpg");
      check(323, "E-Invitation_PARTNER 1_323000-01.jpg");
      check(450, "E-Invitation_PARTNER 1_450000-01.jpg");
    });
  });

  // ── CC email parsing ──────────────────────────────────────────────────────

  describe("CC email parsing", () => {
    it("should parse a single CC email", async () => {
      setupMockWorkbook(validRows({ cc: "anthony_nguyen@asus.com" }));
      const guests = await readGuests("guests.csv");
      expect(guests[0]!.ccEmails).toEqual(["anthony_nguyen@asus.com"]);
    });

    it("should parse multiple CC emails separated by /", async () => {
      setupMockWorkbook(
        validRows({
          cc: "anthony_nguyen@asus.com/ Vu_Bui@asus.com/ Ginny_Tran@asus.com/ Julia_Hoang@asus.com",
        }),
      );
      const guests = await readGuests("guests.csv");
      expect(guests[0]!.ccEmails).toEqual([
        "anthony_nguyen@asus.com",
        "vu_bui@asus.com",
        "ginny_tran@asus.com",
        "julia_hoang@asus.com",
      ]);
    });

    it("should trim whitespace around each CC email", async () => {
      setupMockWorkbook(validRows({ cc: "  a@asus.com  /  b@asus.com  " }));
      const guests = await readGuests("guests.csv");
      expect(guests[0]!.ccEmails).toEqual(["a@asus.com", "b@asus.com"]);
    });

    it("should lowercase all CC emails", async () => {
      setupMockWorkbook(
        validRows({ cc: "Anthony_Nguyen@ASUS.COM/ Vu_Bui@Asus.Com" }),
      );
      const guests = await readGuests("guests.csv");
      expect(guests[0]!.ccEmails).toEqual([
        "anthony_nguyen@asus.com",
        "vu_bui@asus.com",
      ]);
    });

    it("should return empty array when CC email column is empty", async () => {
      setupMockWorkbook(validRows({ cc: "" }));
      const guests = await readGuests("guests.csv");
      expect(guests[0]!.ccEmails).toEqual([]);
    });

    it("should return empty array when CC email column is null", async () => {
      setupMockWorkbook([HEADER, [1, "Guest A", "a@example.com", null]]);
      const guests = await readGuests("guests.csv");
      expect(guests[0]!.ccEmails).toEqual([]);
    });

    it("should ignore empty segments from trailing or double slashes", async () => {
      setupMockWorkbook(validRows({ cc: "a@asus.com/ / b@asus.com/" }));
      const guests = await readGuests("guests.csv");
      expect(guests[0]!.ccEmails).toEqual(["a@asus.com", "b@asus.com"]);
    });

    it("should return empty array when CC email column is not present in header", async () => {
      setupMockWorkbook([
        ["Item", "Customer Name", "Email sending"], // no CC email column
        [1, "Guest A", "a@example.com"],
      ]);
      const guests = await readGuests("guests.csv");
      expect(guests[0]!.ccEmails).toEqual([]);
    });
  });
});
