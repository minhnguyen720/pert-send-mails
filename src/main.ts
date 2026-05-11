import dotenv from "dotenv";
import fs from "fs";
import path from "path";
import Handlebars from "handlebars";
import { readGuests } from "./services/ExcelService";
import { EmailService } from "./services/EmailService";
import { LoggerService } from "./services/LoggerService";
import { sleep } from "./utils/sleep";
import type { SmtpConfig } from "./types";

dotenv.config();

// ─── Sheet config mapping ────────────────────────────────────────────────────
// Add new options here: { data: "data/file.csv", template: "templates/file.html" }
const SHEET_CONFIG: Record<string, { data: string; template: string }> = {
  testEmail: { data: "data/test.csv", template: "templates/invitation.html" },
  disty: {
    data: "data/disty.csv",
    template: "templates/invitation_disty.html",
  },
  partner: { data: "data/partner.csv", template: "templates/invitation.html" },
};

// ─── Pre-flight checks ────────────────────────────────────────────────────────

function preflight(option: string): { dataFile: string; templateFile: string } {
  const requiredEnvVars = ["SMTP_HOST", "SMTP_PORT", "SMTP_USER", "SMTP_PASS"];
  const missingEnv = requiredEnvVars.filter((v) => !process.env[v]);
  if (missingEnv.length > 0) {
    console.error(
      `[Preflight] Missing required .env variables: ${missingEnv.join(", ")}`,
    );
    process.exit(1);
  }

  const availableOptions = Object.keys(SHEET_CONFIG).join(", ");
  const config = SHEET_CONFIG[option];
  if (!config) {
    console.error(
      `[Preflight] Unknown option "${option}". Available: ${availableOptions}`,
    );
    process.exit(1);
  }

  const dataFile = path.resolve(config.data);
  if (!fs.existsSync(dataFile)) {
    console.error(`[Preflight] Data file not found: ${dataFile}`);
    process.exit(1);
  }

  const templateFile = path.resolve(config.template);
  if (!fs.existsSync(templateFile)) {
    console.error(`[Preflight] Template not found: ${templateFile}`);
    process.exit(1);
  }

  return { dataFile, templateFile };
}

// ─── Main ─────────────────────────────────────────────────────────────────────

async function main(): Promise<void> {
  const option = process.argv[2];
  if (!option) {
    const availableOptions = Object.keys(SHEET_CONFIG).join(", ");
    console.error(
      `[Main] Usage: npm run send -- <option>\n  Available options: ${availableOptions}`,
    );
    process.exit(1);
  }

  const { dataFile, templateFile } = preflight(option);

  console.log(`[Main] Reading guests from: ${dataFile}`);
  const guests = await readGuests(dataFile);
  console.log(`[Main] ${guests.length} guests loaded.`);

  const templateSource = fs.readFileSync(templateFile, "utf-8");
  const template = Handlebars.compile(templateSource);

  const smtpConfig: SmtpConfig = {
    host: process.env["SMTP_HOST"]!,
    port: Number(process.env["SMTP_PORT"] ?? 587),
    user: process.env["SMTP_USER"]!,
    pass: process.env["SMTP_PASS"]!,
  };

  const emailService = new EmailService(smtpConfig);
  const logger = new LoggerService(path.resolve("logs"));

  let successCount = 0;
  let failedCount = 0;

  for (let i = 0; i < guests.length; i++) {
    const guest = guests[i]!;
    console.log(`[Main] Sending ${i + 1}/${guests.length}: ${guest.email}`);

    const type = "1"; // partner
    // const type = "2"; // disty

    // const folder = "test";
    const folder = "partner";
    // const folder = "disty";

    const htmlBody = template({ GUEST_NAME: guest.name });
    const itemPadded = String(guest.item).padStart(3, "0");
    const attachmentFileName = `E-Invitation_PARTNER ${type}_${itemPadded}000-01.jpg`;
    const attachmentPath = path.resolve(
      `attachments/${folder}/${attachmentFileName}`,
    );

    const result = await emailService.sendInvitation(
      guest,
      htmlBody,
      attachmentPath,
    );
    logger.appendResult(result);

    if (result.status === "OK") {
      successCount++;
    } else {
      failedCount++;
    }

    if (i < guests.length - 1) {
      await sleep(10000);
    }
  }

  console.log("\n─── Summary ───────────────────────────────");
  console.log(`  Total Processed : ${guests.length}`);
  console.log(`  Total Success   : ${successCount}`);
  console.log(`  Total Failed    : ${failedCount}`);
  console.log(`  Log saved to    : ${logger.getLogPath()}`);
  console.log("───────────────────────────────────────────\n");
}

main().catch((err: unknown) => {
  console.error("[Main] Fatal error:", err);
  process.exit(1);
});
