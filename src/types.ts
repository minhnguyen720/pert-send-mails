export interface Guest {
  name: string;
  email: string;
}

export interface SendResult {
  email: string;
  status: "OK" | "FAILED";
  reason?: string;
  timestamp: string;
}

export interface SmtpConfig {
  host: string;
  port: number;
  user: string;
  pass: string;
}
