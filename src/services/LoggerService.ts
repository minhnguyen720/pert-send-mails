import fs from "fs";
import path from "path";
import type { SendResult } from "../types";

export class LoggerService {
  private logPath: string;

  constructor(logsDir: string) {
    fs.mkdirSync(logsDir, { recursive: true });
    const timestamp = new Date().toISOString().replace(/[:.]/g, "-");
    this.logPath = path.join(logsDir, `send-log-${timestamp}.csv`);
    fs.writeFileSync(this.logPath, "Email,Status,Reason,Timestamp\n", "utf-8");
    console.log(`[LoggerService] Logging to: ${this.logPath}`);
  }

  appendResult(result: SendResult): void {
    const reason = (result.reason ?? "").replace(/"/g, '""');
    const line = `"${result.email}","${result.status}","${reason}","${result.timestamp}"\n`;
    fs.appendFileSync(this.logPath, line, "utf-8");
  }

  getLogPath(): string {
    return this.logPath;
  }
}
