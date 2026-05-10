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

// ─── Data file mapping ────────────────────────────────────────────────────────
// Add new sheets here as needed: "option_name": "data/filename.csv"
const DATA_FILES: Record<string, string> = {
  test: "data/test.csv",
  //   partner: "data/partner.csv",
  //   disty: "data/disty.csv",
};

// ─── Pre-flight checks ────────────────────────────────────────────────────────

function preflight(option: string): string {
  const requiredEnvVars = ["SMTP_HOST", "SMTP_PORT", "SMTP_USER", "SMTP_PASS"];
  const missingEnv = requiredEnvVars.filter((v) => !process.env[v]);
  if (missingEnv.length > 0) {
    console.error(
      `[Preflight] Missing required .env variables: ${missingEnv.join(", ")}`,
    );
    process.exit(1);
  }

  const availableOptions = Object.keys(DATA_FILES).join(", ");
  const relPath = DATA_FILES[option];
  if (!relPath) {
    console.error(
      `[Preflight] Unknown option "${option}". Available: ${availableOptions}`,
    );
    process.exit(1);
  }

  const dataFile = path.resolve(relPath);
  if (!fs.existsSync(dataFile)) {
    console.error(`[Preflight] Data file not found: ${dataFile}`);
    process.exit(1);
  }

  const templateFile = path.resolve("templates/invitation.html");
  if (!fs.existsSync(templateFile)) {
    console.error(`[Preflight] Template not found: ${templateFile}`);
    process.exit(1);
  }

  return dataFile;
}

// ─── Main ─────────────────────────────────────────────────────────────────────

async function main(): Promise<void> {
  const option = process.argv[2];
  if (!option) {
    const availableOptions = Object.keys(DATA_FILES).join(", ");
    console.error(
      `[Main] Usage: npm run send -- <option>\n  Available options: ${availableOptions}`,
    );
    process.exit(1);
  }

  const dataFile = preflight(option);
  const templateFile = path.resolve("templates/invitation.html");
  //   const distyTemplateFile = path.resolve("templates/invitation_disty.html");

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

    const htmlBody = template({ GUEST_NAME: guest.name });
    const jpegPath = path.resolve(`attachments/${guest.name}.jpeg`);
    const attachmentPath = fs.existsSync(jpegPath) ? jpegPath : null;

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
      await sleep(2000);
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
