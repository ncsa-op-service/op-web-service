export type LinkStatus =
  | "WAITING"
  | "FOUND"
  | "NOT_LINE"
  | "INVALID_URL"
  | "NOT_FOUND"
  | "TIMEOUT"
  | "ERROR";

export type ExcelRow = {
  name?: string;
  line_url?: string;
  url?: string;
};