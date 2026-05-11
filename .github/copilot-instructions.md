# 🛠️ Copilot Skill: Batch Invitation Email Sender

## 🎯 Objective

You are instructed to act as an automated orchestration agent. Your task is to execute the batch email invitation process using the provided Node.js/TypeScript architecture.

## 📋 Context & Rules

- **Environment:** Node.js with TypeScript (`ts-node`).
- **Data Source:** `data/guests.xlsx` (Requires columns: `Customer Name`, `Email sending`).
- **Attachments:** Images mapped by email address (e.g., `attachments/user@domain.com.jpg`).
- **Rate Limiting:** A mandatory 2000ms delay must be respected between each email sent to prevent SMTP throttling.
- **Error Handling:** Missing attachments or SMTP errors must NOT halt the entire process. They must be caught, logged as `FAILED`, and the loop must continue.

## ⚙️ Execution Environment

- **Language/Runtime:** Node.js + TypeScript
- **Core Libraries:** `nodemailer`, `exceljs`, `handlebars`, `dotenv`

## 📂 Project Structure

The agent must expect and adhere strictly to the following directory structure:

pert-send-mails/
├── data/
│ └── guests.xlsx # Source data
├── attachments/
│ └── [E-Invitation_PARTNER 1_001000-01].jpg # ~450 images mapped by customer order in the list
├── templates/
│ └── invitation.html # HTML template with {{GUEST_NAME}}
├── logs/
│ └── [auto-generated-csv] # Audit logs
├── src/
│ ├── types.ts # TS interfaces
│ ├── services/
│ ├──── **tests**/ # New: Test files go here
│ │ ├──── EmailService.test.ts
│ │ ├──── ExcelService.test.ts
│ │ └──── LoggerService.test.ts
│ │ ├── ExcelService.ts # exceljs logic
│ │ ├── EmailService.ts # nodemailer logic
│ │ └── LoggerService.ts # fs logic for CSV appending
│ ├── utils/
│ │ └── sleep.ts # Delay function
│ └── main.ts # Entry point and loop orchestration
├── .env # SMTP credentials
├── tsconfig.json
└── package.json

## 🚀 Execution Steps

When the user asks you to "Send the invitations", you must execute the following protocol strictly in order:

1. **Pre-flight Check:**
   - Verify that `.env` exists and contains `SMTP_HOST`, `SMTP_USER`, and `SMTP_PASS`.
   - Verify `data/guests.xlsx` exists.
   - Verify `templates/invitation.html` exists.
     _If any are missing, halt execution and alert the user._

2. **Trigger Execution:**
   - Suggest the user run the following command in their integrated terminal:
     ```bash
     npx ts-node src/main.ts
     ```

3. **Post-Execution Verification:**
   - Once the script finishes, locate the newly generated CSV file in the `logs/` directory.
   - Read the CSV file.
   - Provide the user with a summary table in the chat, formatted as follows:
     - Total Processed: [X]
     - Total Success: [Y]
     - Total Failed: [Z]
   - List any specific email addresses that failed and the reason for failure.

## ⚠️ Constraints

- Do not attempt to rewrite the core logic of `src/main.ts` unless explicitly asked to debug an error.
- Do not execute the terminal command yourself without the user's explicit permission.
