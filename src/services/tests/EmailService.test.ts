import fs from "fs";
import nodemailer from "nodemailer";
import { EmailService } from "../EmailService";
import type { Guest, SmtpConfig } from "../../types";

// ─── Module mocks ─────────────────────────────────────────────────────────────

jest.mock("nodemailer");
jest.mock("fs");

// ─── Helpers ──────────────────────────────────────────────────────────────────

const mockSmtpConfig: SmtpConfig = {
  host: "smtp.example.com",
  port: 587,
  user: "test@example.com",
  pass: "password",
};

const makeGuest = (index: number, ccEmails: string[] = []): Guest => ({
  item: index,
  name: `Guest ${index}`,
  email: `guest${index}@example.com`,
  ccEmails,
});

const HTML_BODY = "<p>Hello {{GUEST_NAME}}</p>";

// ─── Tests ────────────────────────────────────────────────────────────────────

describe("EmailService", () => {
  let mockSendMail: jest.Mock;
  let service: EmailService;

  beforeEach(() => {
    jest.resetAllMocks();

    mockSendMail = jest.fn().mockResolvedValue({ messageId: "mock-id" });

    (nodemailer.createTransport as jest.Mock).mockReturnValue({
      sendMail: mockSendMail,
    });

    // By default, attachments do not exist
    (fs.existsSync as jest.Mock).mockReturnValue(false);

    process.env["SMTP_USER"] = "test@example.com";

    service = new EmailService(mockSmtpConfig);
  });

  afterEach(() => {
    delete process.env["SMTP_USER"];
  });

  // ── Constructor ─────────────────────────────────────────────────────────────

  describe("constructor", () => {
    it("should create a transporter with the given SMTP config", () => {
      expect(nodemailer.createTransport).toHaveBeenCalledWith({
        host: mockSmtpConfig.host,
        port: mockSmtpConfig.port,
        secure: false,
        auth: { user: mockSmtpConfig.user, pass: mockSmtpConfig.pass },
      });
    });

    it("should set secure=true when port is 465", () => {
      new EmailService({ ...mockSmtpConfig, port: 465 });
      expect(nodemailer.createTransport).toHaveBeenCalledWith(
        expect.objectContaining({ secure: true }),
      );
    });
  });

  // ── sendInvitation — success paths ──────────────────────────────────────────

  describe("sendInvitation — success", () => {
    it("should return status OK on successful send", async () => {
      const guest = makeGuest(1);
      const result = await service.sendInvitation(guest, HTML_BODY, null);

      expect(result.status).toBe("OK");
      expect(result.email).toBe(guest.email);
      expect(result.timestamp).toMatch(/^\d{4}-\d{2}-\d{2}T\d{2}:\d{2}:\d{2}/);
    });

    it("should call sendMail with correct fields", async () => {
      const guest = makeGuest(1);
      await service.sendInvitation(guest, HTML_BODY, null);

      expect(mockSendMail).toHaveBeenCalledWith(
        expect.objectContaining({
          to: guest.email,
          subject:
            "THƯ MỜI THAM DỰ SỰ KIỆN ASUS EXPERTBOOK ULTRA GRAND LAUNCH - The Flagship of the Industry. Period.",
          html: HTML_BODY,
        }),
      );
    });

    it("should send without attachment when attachmentPath is null", async () => {
      await service.sendInvitation(makeGuest(1), HTML_BODY, null);

      expect(mockSendMail).toHaveBeenCalledWith(
        expect.objectContaining({ attachments: [] }),
      );
    });

    it("should not include cc when guest has no CC emails", async () => {
      await service.sendInvitation(makeGuest(1, []), HTML_BODY, null);

      expect(mockSendMail).toHaveBeenCalledWith(
        expect.objectContaining({ cc: undefined }),
      );
    });

    it("should include cc when guest has one CC email", async () => {
      await service.sendInvitation(
        makeGuest(1, ["anthony_nguyen@asus.com"]),
        HTML_BODY,
        null,
      );

      expect(mockSendMail).toHaveBeenCalledWith(
        expect.objectContaining({
          cc: "anthony_nguyen@asus.com, quynhnhu19111@gmail.com",
        }),
      );
    });

    it("should join multiple CC emails with comma-space", async () => {
      await service.sendInvitation(
        makeGuest(1, [
          "anthony_nguyen@asus.com",
          "vu_bui@asus.com",
          "ginny_tran@asus.com",
        ]),
        HTML_BODY,
        null,
      );

      expect(mockSendMail).toHaveBeenCalledWith(
        expect.objectContaining({
          cc: "anthony_nguyen@asus.com, vu_bui@asus.com, ginny_tran@asus.com, quynhnhu19111@gmail.com",
        }),
      );
    });

    it("should include inline attachment with CID when file exists", async () => {
      (fs.existsSync as jest.Mock).mockReturnValue(true);
      const attachmentPath = "/attachments/Guest 1.jpeg";

      await service.sendInvitation(makeGuest(1), HTML_BODY, attachmentPath);

      expect(mockSendMail).toHaveBeenCalledWith(
        expect.objectContaining({
          attachments: [{ path: attachmentPath, cid: "invitationImage" }],
        }),
      );
    });

    it("should send without attachment when file does not exist", async () => {
      (fs.existsSync as jest.Mock).mockReturnValue(false);

      await service.sendInvitation(
        makeGuest(1),
        HTML_BODY,
        "/attachments/missing.jpeg",
      );

      expect(mockSendMail).toHaveBeenCalledWith(
        expect.objectContaining({ attachments: [] }),
      );
    });
  });

  // ── sendInvitation — error handling ─────────────────────────────────────────

  describe("sendInvitation — error handling", () => {
    it("should return status FAILED when sendMail throws an Error", async () => {
      mockSendMail.mockRejectedValue(new Error("Connection refused"));
      const guest = makeGuest(1);

      const result = await service.sendInvitation(guest, HTML_BODY, null);

      expect(result.status).toBe("FAILED");
      expect(result.email).toBe(guest.email);
      expect(result.reason).toBe("Connection refused");
    });

    it("should return status FAILED when sendMail throws a non-Error", async () => {
      mockSendMail.mockRejectedValue("timeout");

      const result = await service.sendInvitation(
        makeGuest(1),
        HTML_BODY,
        null,
      );

      expect(result.status).toBe("FAILED");
      expect(result.reason).toBe("timeout");
    });

    it("should NOT throw even when sendMail fails", async () => {
      mockSendMail.mockRejectedValue(new Error("SMTP error"));

      await expect(
        service.sendInvitation(makeGuest(1), HTML_BODY, null),
      ).resolves.not.toThrow();
    });
  });

  // ── Bulk — 450 guests ────────────────────────────────────────────────────────

  describe("bulk sending — 450 guests", () => {
    const GUEST_COUNT = 450;

    it("should successfully send to all 450 guests and return OK for each", async () => {
      const guests = Array.from({ length: GUEST_COUNT }, (_, i) =>
        makeGuest(i + 1),
      );

      const results = await Promise.all(
        guests.map((g) => service.sendInvitation(g, HTML_BODY, null)),
      );

      expect(mockSendMail).toHaveBeenCalledTimes(GUEST_COUNT);
      expect(results).toHaveLength(GUEST_COUNT);
      results.forEach((r, i) => {
        expect(r.status).toBe("OK");
        expect(r.email).toBe(guests[i]!.email);
      });
    });

    it("should return FAILED for each guest when SMTP is down, never throwing", async () => {
      mockSendMail.mockRejectedValue(new Error("SMTP unavailable"));
      const guests = Array.from({ length: GUEST_COUNT }, (_, i) =>
        makeGuest(i + 1),
      );

      const results = await Promise.all(
        guests.map((g) => service.sendInvitation(g, HTML_BODY, null)),
      );

      expect(results).toHaveLength(GUEST_COUNT);
      results.forEach((r) => {
        expect(r.status).toBe("FAILED");
        expect(r.reason).toBe("SMTP unavailable");
      });
    });

    it("should handle mixed success and failure across 450 guests without throwing", async () => {
      // Odd-indexed calls fail, even-indexed succeed
      let callCount = 0;
      mockSendMail.mockImplementation(() => {
        callCount++;
        return callCount % 2 === 0
          ? Promise.reject(new Error("Intermittent error"))
          : Promise.resolve({ messageId: "ok" });
      });

      const guests = Array.from({ length: GUEST_COUNT }, (_, i) =>
        makeGuest(i + 1),
      );
      const results = await Promise.all(
        guests.map((g) => service.sendInvitation(g, HTML_BODY, null)),
      );

      const successes = results.filter((r) => r.status === "OK");
      const failures = results.filter((r) => r.status === "FAILED");

      expect(successes.length + failures.length).toBe(GUEST_COUNT);
      expect(successes.length).toBe(GUEST_COUNT / 2);
      expect(failures.length).toBe(GUEST_COUNT / 2);
    });

    it("should attach inline images for all 450 guests when files exist", async () => {
      (fs.existsSync as jest.Mock).mockReturnValue(true);
      const guests = Array.from({ length: GUEST_COUNT }, (_, i) =>
        makeGuest(i + 1),
      );

      await Promise.all(
        guests.map((g) =>
          service.sendInvitation(g, HTML_BODY, `/attachments/${g.name}.jpeg`),
        ),
      );

      expect(mockSendMail).toHaveBeenCalledTimes(GUEST_COUNT);
      // Every call should have included the inline attachment
      mockSendMail.mock.calls.forEach((call) => {
        expect(call[0].attachments[0]).toMatchObject({
          cid: "invitationImage",
        });
      });
    });
  });
});
