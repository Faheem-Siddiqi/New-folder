import * as XLSX from "xlsx";
import type { StrengthData } from "./strength-data";

import { TURNOVER_HEADERS, validateUpload, type Employee, type EmployeeRecord, type ReportRow, type ScanProgress, type TurnoverResult } from "./turnover-contracts.ts";
export * from "./turnover-contracts.ts";
const MAX_ROWS = 100_000;
const MAX_CELLS = 5_000_000;
export const normalizeMatch = (value: string) => value.normalize("NFKC").toLowerCase().replace(/[^a-z0-9]/g, "");
const aliases = [
  ["department", "dept", "departmentname"], ["unit", "unitname"], ["emp", "empno", "empnumber", "employeeid", "employeeno", "employeenumber", "empid"],
  ["employeename", "empname", "name"], ["designation", "jobtitle"], ["gender", "sex"], ["restday", "weeklyoff", "offday"],
  ["grade", "employeegrade"], ["shift", "shiftname"], ["paysheet", "payroll", "paysheetname"], ["orggroup", "organizationgroup", "organisationgroup"], ["joiningdate", "dateofjoining", "doj"],
];


function validateSignature(bytes: Uint8Array, name: string) {
  const zip = bytes[0] === 0x50 && bytes[1] === 0x4b && bytes[2] === 3 && bytes[3] === 4;
  const ole = [0xd0, 0xcf, 0x11, 0xe0, 0xa1, 0xb1, 0x1a, 0xe1].every((value, i) => bytes[i] === value);
  if ((/\.xlsx$/i.test(name) && !zip) || (/\.xls$/i.test(name) && !ole)) throw new Error("The file contents do not match its Excel extension. Open it in Excel and use Save As to create a valid .xlsx workbook. Password-protected files must be unlocked first.");
  if (zip) {
    // Inspect central-directory sizes before inflating a potentially enormous workbook.
    const view = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
    let expanded = 0;
    for (let offset = 0; offset + 46 <= bytes.length; offset++) {
      if (view.getUint32(offset, true) !== 0x02014b50) continue;
      const length = view.getUint32(offset + 24, true);
      expanded += length;
      if (expanded > 200 * 1024 * 1024 || length === 0xffffffff) throw new Error("The workbook expands beyond the supported size. Save a smaller, unencrypted workbook and try again.");
      offset += 45 + view.getUint16(offset + 28, true) + view.getUint16(offset + 30, true) + view.getUint16(offset + 32, true);
    }
  }
}

function detectHeader(row: string[]) {
  const indexes = aliases.map((names) => row.map((value) => normalizeMatch(String(value ?? ""))).reduce<number[]>((matches, value, index) => names.includes(value) ? [...matches, index] : matches, []));
  return { count: indexes.filter((index) => index.length > 0).length, indexes, complete: indexes.every((index) => index.length === 1) };
}

export function templateRows(template: StrengthData): ReportRow[] {
  return template.strengthStructure.flatMap((category) => category.subcategories.flatMap((subcategory) => subcategory.designations.map((designation) => ({ ...designation, category: category.category, subcategory: subcategory.subcategory, onRoll: 0 }))));
}

export function matchEmployee(employee: Employee, rows: ReportRow[]): { match: number | null; issue: string } {
  const designation = normalizeMatch(employee.Designation);
  const grade = normalizeMatch(employee.Grade);
  const contexts = new Set([employee.Department, employee.Unit, employee["Org Group"], `${employee.Department} ${employee.Unit}`, `${employee.Department} ${employee["Org Group"]}`].map(normalizeMatch).filter(Boolean));
  let candidates = rows.map((row, index) => ({ row, index })).filter(({ row }) => normalizeMatch(row.designation) === designation && normalizeMatch(row.grade) === grade && contexts.has(normalizeMatch(row.subcategory)));
  if (candidates.length > 1) {
    const narrowed = candidates.filter(({ row }) => [employee.Department, employee["Org Group"]].some((value) => normalizeMatch(value) === normalizeMatch(row.category)));
    if (narrowed.length) candidates = narrowed;
  }
  if (candidates.length === 1) return { match: candidates[0].index, issue: "" };
  return { match: null, issue: candidates.length ? "Ambiguous template match. Review department/unit, designation, and grade." : "No matching template department/unit, designation, and grade." };
}

export function scanTurnover(buffer: ArrayBuffer, fileName: string, template: StrengthData, progress: ScanProgress = () => {}): TurnoverResult {
  validateUpload(fileName, buffer.byteLength);
  progress("Validating file format", 8);
  validateSignature(new Uint8Array(buffer), fileName);
  progress("Reading Excel worksheets", 18);
  let workbook: XLSX.WorkBook;
  try { workbook = XLSX.read(buffer, { type: "array", cellDates: true, cellFormula: false, cellHTML: false, cellStyles: false, WTF: false }); }
  catch { throw new Error("This workbook is unreadable, corrupted, or password-protected. Open it in Excel, repair it if prompted, and save a new .xlsx copy."); }
  if (!workbook.SheetNames.length || workbook.SheetNames.length > 100) throw new Error("The workbook must contain between 1 and 100 worksheets.");
  const rows = templateRows(template);
  if (!rows.length) throw new Error("The strength template has no designations. Add them on the Configuration page before scanning.");
  const result: TurnoverResult = { version: 1, fileName, generatedAt: new Date().toISOString(), templateSignature: JSON.stringify(template), categories: template.strengthStructure.map((category) => category.category), rows, employees: [], issues: [], scannedRows: 0, duplicates: 0, skippedRows: 0, matched: 0, unmatched: 0 };
  const seen = new Map<string, EmployeeRecord>();
  let completeSheets = 0;
  let totalCells = 0;
  let bestHeader: { sheet: string; count: number; missing: string[] } | null = null;
  workbook.SheetNames.forEach((sheetName, sheetIndex) => {
    progress(`Detecting headers: ${sheetName}`, 25 + Math.round(40 * sheetIndex / workbook.SheetNames.length));
    const sheet = workbook.Sheets[sheetName];
    if (!sheet?.["!ref"]) { result.issues.push({ sheet: sheetName, row: 0, employeeId: "", message: "Empty or unreadable worksheet skipped." }); return; }
    const range = XLSX.utils.decode_range(sheet["!ref"]);
    const cellCount = (range.e.r - range.s.r + 1) * (range.e.c - range.s.c + 1);
    totalCells += cellCount;
    if (range.e.r >= MAX_ROWS || range.e.c >= 256 || totalCells > MAX_CELLS) throw new Error("The workbook is too large or has excessive empty formatting. Limit it to 100,000 rows per sheet, 256 columns, and 5 million cells, then save again.");
    let matrix: string[][];
    try { matrix = XLSX.utils.sheet_to_json<string[]>(sheet, { header: 1, raw: false, defval: "", blankrows: true, range: 0 }); }
    catch { result.issues.push({ sheet: sheetName, row: 0, employeeId: "", message: "Unreadable worksheet skipped; review the original workbook." }); return; }
    let header: number[] | null = null;
    let found = false;
    for (let index = 0; index < matrix.length; index++) {
      if (index % 500 === 0) progress(`Scanning ${sheetName}: row ${index + 1} of ${matrix.length}`, 25 + Math.round(40 * (sheetIndex + index / matrix.length) / workbook.SheetNames.length));
      const values = matrix[index] ?? [];
      const detected = detectHeader(values);
      if (!detected.complete && detected.count >= 4 && matrix[index + 1]) {
        const next = matrix[index + 1];
        const combined = values.map((value, column) => {
          const lower = String(next[column] ?? "").trim();
          const joined = `${value} ${lower}`;
          if (aliases.some((names) => names.includes(normalizeMatch(joined)))) return joined;
          if (!lower && aliases.some((names) => names.includes(normalizeMatch(String(value))))) return value;
          return "";
        });
        const stacked = detectHeader(combined);
        if (stacked.complete) { header = stacked.indexes.map((indexes) => indexes[0]); found = true; index++; continue; }
      }
      if (detected.count >= 4 && (!bestHeader || detected.count > bestHeader.count)) bestHeader = { sheet: sheetName, count: detected.count, missing: TURNOVER_HEADERS.filter((_, i) => detected.indexes[i].length !== 1) };
      if (detected.complete) { header = detected.indexes.map((indexes) => indexes[0]); found = true; continue; }
      if (detected.count >= 6) { header = null; result.issues.push({ sheet: sheetName, row: index + 1, employeeId: "", message: `Incomplete or duplicate headers: ${TURNOVER_HEADERS.filter((_, i) => detected.indexes[i].length !== 1).join(", ")}. Section skipped.` }); continue; }
      if (!header || values.every((value) => !String(value).trim())) continue;
      const employee = Object.fromEntries(TURNOVER_HEADERS.map((name, i) => {
        let value = values[header![i]] ?? "";
        // Carry department/unit context only when the source explicitly merges those cells.
        if (!value && [0, 1, 10].includes(i)) {
          const merge = sheet["!merges"]?.find((item) => item.s.c === header![i] && item.s.r < index && item.e.r >= index);
          if (merge) value = matrix[merge.s.r]?.[header![i]] ?? "";
        }
        return [name, String(value).trim()];
      })) as Employee;
      if (Object.values(employee).every((value) => !value)) continue;
      const footer = /^(grand\s*total|sub\s*total|total)\b/i.test(employee.Department) || /^(grand\s*total|sub\s*total|total)\b/i.test(employee["Emp #"]);
      if (footer) continue;
      result.scannedRows++;
      if (result.scannedRows > MAX_ROWS) throw new Error("The report contains more than 100,000 employee rows. Split it into a smaller workbook.");
      if (!employee["Emp #"] || !employee.Designation || !employee.Grade || (!employee.Department && !employee.Unit)) {
        result.skippedRows++;
        result.issues.push({ sheet: sheetName, row: index + 1, employeeId: employee["Emp #"], message: "Row excluded: employee number, designation, grade, and department or unit are required." });
        continue;
      }
      const record: EmployeeRecord = { employee, sheet: sheetName, row: index + 1, ...matchEmployee(employee, rows) };
      const identity = `${normalizeMatch(employee.Unit)}|${employee["Emp #"].toLowerCase()}`;
      const previous = seen.get(identity);
      if (previous) {
        result.duplicates++;
        const identical = TURNOVER_HEADERS.every((name) => normalizeMatch(previous.employee[name]) === normalizeMatch(employee[name]));
        if (!identical) { previous.match = null; previous.issue = "Conflicting records for the same employee number and unit; excluded from counts."; record.issue = previous.issue; record.match = null; result.employees.push(record); }
        result.issues.push({ sheet: sheetName, row: index + 1, employeeId: employee["Emp #"], message: identical ? "Duplicate employee number/unit record skipped." : record.issue });
        continue;
      }
      seen.set(identity, record);
      result.employees.push(record);
      if (!employee["Employee Name"]) result.issues.push({ sheet: sheetName, row: index + 1, employeeId: employee["Emp #"], message: "Employee name is blank in the source." });
    }
    if (found) completeSheets++;
    else result.issues.push({ sheet: sheetName, row: 0, employeeId: "", message: "No complete turnover header found; worksheet skipped." });
  });
  if (!completeSheets) {
    const detected = bestHeader as { sheet: string; count: number; missing: string[] } | null;
    throw new Error(detected ? `Missing or duplicate headers in ${detected.sheet}: ${detected.missing.join(", ")}. All 12 turnover headers are required.` : "No turnover table found. Include all 12 required column headers; title rows and reordered columns are supported.");
  }
  if (!result.employees.length) throw new Error("No valid employee records were found below the headers. Check employee numbers, designation, grade, and department/unit values.");
  progress("Matching employees to the strength template", 80);
  for (const record of result.employees) {
    if (record.match !== null) { result.rows[record.match].onRoll++; result.matched++; }
    else { result.unmatched++; result.issues.push({ sheet: record.sheet, row: record.row, employeeId: record.employee["Emp #"], message: record.issue }); }
  }
  progress("Preparing report and validation summary", 94);
  return result;
}
