import nodemailer from "nodemailer";
import fs from "fs";
import type { Guest, SendResult, SmtpConfig } from "../types";
import moment from "moment";

export class EmailService {
  private transporter: nodemailer.Transporter;

  constructor(config: SmtpConfig) {
    this.transporter = nodemailer.createTransport({
      host: config.host,
      port: config.port,
      secure: true,
      auth: {
        user: config.user,
        pass: config.pass,
      },
    });
  }

  async sendInvitation(
    guest: Guest,
    htmlBody: string,
    attachmentPath: string | null,
  ): Promise<SendResult> {
    const timestamp = moment().format("DD/MM/YYYY h:mm a");

    const attachments: nodemailer.SendMailOptions["attachments"] = [];
    if (attachmentPath) {
      if (fs.existsSync(attachmentPath)) {
        attachments.push({ path: attachmentPath, cid: "invitationImage" });
      } else {
        console.warn(
          `[EmailService] Attachment not found for ${guest.email}: ${attachmentPath}`,
        );
      }
    }

    try {
      await this.transporter.sendMail({
        from: process.env["SMTP_USER"],
        to: guest.email,
        cc: guest.ccEmails.length > 0 ? guest.ccEmails.join(", ") : undefined,
        subject:
          "THƯ MỜI THAM DỰ SỰ KIỆN ASUS EXPERTBOOK ULTRA GRAND LAUNCH - The Flagship of the Industry. Period.",
        html: htmlBody,
        attachments,
      });

      return {
        email: guest.email,
        status: "OK",
        timestamp,
        sentBy: process.env["SMTP_USER"] as string,
        ccEmails: (guest.ccEmails ?? []).join(", "),
      };
    } catch (err) {
      const reason = err instanceof Error ? err.message : String(err);
      console.error(
        `[EmailService] Failed to send to ${guest.email}: ${reason}`,
      );
      return {
        email: guest.email,
        status: "FAILED",
        reason,
        timestamp,
        sentBy: process.env["SMTP_USER"] as string,
        ccEmails: (guest.ccEmails ?? []).join(", "),
      };
    }
  }
}
