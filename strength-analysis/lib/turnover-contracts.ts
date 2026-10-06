export const TURNOVER_HEADERS = ["Department", "Unit", "Emp #", "Employee Name", "Designation", "Gender", "Rest Day", "Grade", "Shift", "Pay Sheet", "Org Group", "Joining Date"] as const;
export type Employee = Record<(typeof TURNOVER_HEADERS)[number], string>;
export type ReportRow = { category: string; subcategory: string; designation: string; grade: string; cadre: string; approvedStrength: number; onRoll: number; status: string };
export type EmployeeRecord = { employee: Employee; sheet: string; row: number; match: number | null; issue: string };
export type ReportIssue = { sheet: string; row: number; employeeId: string; message: string };
export type TurnoverResult = { version: 1; fileName: string; generatedAt: string; templateSignature: string; categories: string[]; rows: ReportRow[]; employees: EmployeeRecord[]; issues: ReportIssue[]; scannedRows: number; duplicates: number; skippedRows: number; matched: number; unmatched: number };
export type ScanProgress = (step: string, percent: number) => void;
export const MAX_FILE_SIZE = 20 * 1024 * 1024;

export function validateUpload(name: string, size: number) {
  if (!/\.(xlsx|xls)$/i.test(name)) throw new Error("Choose an Excel workbook (.xlsx or .xls). CSV, PDF, and renamed files are not supported.");
  if (!size) throw new Error("This file is empty. Export the report from Excel and upload it again.");
  if (size > MAX_FILE_SIZE) throw new Error("This workbook exceeds 20 MB. Split the report into a smaller workbook.");
}

