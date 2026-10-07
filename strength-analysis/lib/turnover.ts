import * as XLSX from "xlsx";
import type { StrengthData } from "./strength-data";
import { createEmployeeMatcher, normalizeMatch, sameEmployeeIdentity } from "./employee-matching.ts";
import { strengthRows } from "./strength-rows.ts";
export { matchEmployee, normalizeMatch } from "./employee-matching.ts";

import { MATCHING_VERSION, TURNOVER_HEADERS, validateUpload, type Employee, type EmployeeRecord, type ReportRow, type ScanProgress, type TurnoverResult } from "./turnover-contracts.ts";
export * from "./turnover-contracts.ts";
const MAX_ROWS = 100_000;
const MAX_CELLS = 5_000_000;
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
  return { count: indexes.filter((index) => index.length > 0).length, indexes, complete: indexes.every((index, column) => column === 7 || index.length === 1) };
}

export function templateRows(template: StrengthData): ReportRow[] {
  return strengthRows(template, true);
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
  const matchEmployee = createEmployeeMatcher(rows);
  if (!rows.length) throw new Error("The strength template has no designations. Add them on the Configuration page before scanning.");
  const result: TurnoverResult = { version: 1, matchingVersion: MATCHING_VERSION, fileName, generatedAt: new Date().toISOString(), templateSignature: JSON.stringify(template), categories: template.strengthStructure.map((category) => category.category), rows, employees: [], issues: [], scannedRows: 0, duplicates: 0, skippedRows: 0, matched: 0, unmatched: 0 };
  const seen = new Map<string, EmployeeRecord>();
  const conflicts = new Set<string>();
  const identityNames = new Map<string, string>();
  const repeated: { identity: string; sheet: string; row: number; employeeId: string }[] = [];
  let apprenticeRows = 0;
  let nonApprenticeRows = 0;
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
        if (stacked.complete) { header = stacked.indexes.map((indexes) => indexes[0] ?? -1); found = true; index++; continue; }
      }
      if (detected.count >= 4 && (!bestHeader || detected.count > bestHeader.count)) bestHeader = { sheet: sheetName, count: detected.count, missing: TURNOVER_HEADERS.filter((_, i) => i !== 7 && detected.indexes[i].length !== 1) };
      if (detected.complete) {
        header = detected.indexes.map((indexes) => indexes[0] ?? -1); found = true;
        if (detected.indexes[7].length > 1) result.issues.push({ sheet: sheetName, row: index + 1, employeeId: "", severity: "info", outcome: "notice", message: "Multiple Grade columns found. The first is displayed as employee information; grade does not affect strength matching." });
        continue;
      }
      if (detected.count >= 6) { header = null; result.issues.push({ sheet: sheetName, row: index + 1, employeeId: "", message: `Incomplete or duplicate headers: ${TURNOVER_HEADERS.filter((_, i) => i !== 7 && detected.indexes[i].length !== 1).join(", ")}. Section skipped.` }); continue; }
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
      if (normalizeMatch(employee.Designation) === "apprentice") { apprenticeRows++; continue; }
      nonApprenticeRows++;
      result.scannedRows++;
      if (result.scannedRows > MAX_ROWS) throw new Error("The report contains more than 100,000 employee rows. Split it into a smaller workbook.");
      if (!employee["Emp #"] || !employee.Designation || (!employee.Department && !employee.Unit)) {
        result.skippedRows++;
        result.issues.push({ sheet: sheetName, row: index + 1, employeeId: employee["Emp #"], message: "Row excluded: employee number, designation, and department or unit are required.", severity: "error", outcome: "excluded" });
        continue;
      }
      const record: EmployeeRecord = { employee, sheet: sheetName, row: index + 1, ...matchEmployee(employee) };
      const identity = `${normalizeMatch(employee.Unit)}|${employee["Emp #"].toLowerCase()}`;
      const previous = seen.get(identity);
      if (previous) {
        result.duplicates++;
        const name = normalizeMatch(employee["Employee Name"]);
        const identical = sameEmployeeIdentity(previous.employee, employee) && (!name || !identityNames.has(identity) || identityNames.get(identity) === name);
        if (name && !identityNames.has(identity)) identityNames.set(identity, name);
        if (!identical || conflicts.has(identity)) {
          conflicts.add(identity);
          previous.match = null;
          previous.issue = `Conflicting identity or strength fields for the same employee number/unit. Compare ${previous.sheet} row ${previous.row} with ${sheetName} row ${index + 1}: employee name, department, or designation differs. These records are excluded until corrected.`;
          record.issue = previous.issue; record.match = null; result.employees.push(record);
        } else repeated.push({ identity, sheet: sheetName, row: index + 1, employeeId: employee["Emp #"] });
        continue;
      }
      seen.set(identity, record);
      if (employee["Employee Name"]) identityNames.set(identity, normalizeMatch(employee["Employee Name"]));
      result.employees.push(record);
      if (!employee["Employee Name"]) result.issues.push({ sheet: sheetName, row: index + 1, employeeId: employee["Emp #"], message: "Employee name is blank in the source; this does not prevent strength matching.", severity: "info", outcome: "notice" });
    }
    if (found) completeSheets++;
    else result.issues.push({ sheet: sheetName, row: 0, employeeId: "", message: "No complete turnover header found; worksheet skipped." });
  });
  if (!completeSheets) {
    const detected = bestHeader as { sheet: string; count: number; missing: string[] } | null;
    throw new Error(detected ? `Missing or duplicate headers in ${detected.sheet}: ${detected.missing.join(", ")}. Grade is optional; the other turnover headers are required.` : "No turnover table found. Include the turnover column headers; Grade is optional. Title rows and reordered columns are supported.");
  }
  if (!result.employees.length && (apprenticeRows === 0 || nonApprenticeRows > 0)) throw new Error("No valid employee records were found below the headers. Check employee numbers, designation, and department/unit values.");
  progress("Matching employees to the strength template", 80);
  for (const repeat of repeated) {
    const first = seen.get(repeat.identity)!;
    const counted = first.match !== null;
    result.issues.push({ sheet: repeat.sheet, row: repeat.row, employeeId: repeat.employeeId, severity: counted ? "info" : "warning", outcome: counted ? "counted" : "excluded", message: counted ? `Repeated employee number/unit: counted once from ${first.sheet} row ${first.row}. Grade, shift, rest day, and other employee-detail differences do not exclude this employee.` : `Repeated employee number/unit; its first record is excluded. ${first.issue}` });
  }
  for (const record of result.employees) {
    if (record.match !== null) { result.rows[record.match].onRoll++; result.matched++; }
    else { result.unmatched++; result.issues.push({ sheet: record.sheet, row: record.row, employeeId: record.employee["Emp #"], message: record.issue, severity: "error", outcome: "excluded" }); }
  }
  progress("Preparing report and validation summary", 94);
  return result;
}
