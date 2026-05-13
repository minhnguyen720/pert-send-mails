export interface Guest {
  item: number;
  name: string;
  email: string;
  ccEmails: string[];
}

export interface SendResult {
  email: string;
  status: "OK" | "FAILED";
  reason?: string;
  timestamp: string;
  sentBy?: string;
  ccEmails?: string;
}

export interface SmtpConfig {
  host: string;
  port: number;
  user: string;
  pass: string;
}
