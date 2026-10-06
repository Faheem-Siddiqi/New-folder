import test from "node:test";
import assert from "node:assert/strict";
import { readFile } from "node:fs/promises";
import * as XLSX from "xlsx";
import ExcelJS from "exceljs";
import { scanTurnover, TURNOVER_HEADERS, validateUpload, matchEmployee, templateRows } from "./turnover.ts";
import { createReportWorkbook } from "./report-workbook.ts";

const template = {
  strengthStructure: [
    { category: "RING", subcategories: [{ subcategory: "SPINNING - RING - 1", designations: [{ designation: "OPERATOR", grade: "E-03", cadre: "Worker", approvedStrength: 8, onRoll: 999, status: "Active" }, { designation: "SUPERVISOR", grade: "E-04", cadre: "Staff", approvedStrength: 1, onRoll: 12, status: "Inactive" }] }] },
    { category: "HR / ADMIN", subcategories: [{ subcategory: "Human Resource", designations: [{ designation: "OFFICER", grade: "M-11", cadre: "Staff", approvedStrength: 2, onRoll: 0, status: "Active" }] }] },
  ],
};
const employee = (changes = {}) => ({ Department: "Spinning Ring", Unit: "1", "Emp #": "0001", "Employee Name": "Sample Employee", Designation: "Operator", Gender: "M", "Rest Day": "Sunday", Grade: "E03", Shift: "A", "Pay Sheet": "Workers", "Org Group": "RING", "Joining Date": "2024-01-15", ...changes });
const employeeRow = (entry, headers = TURNOVER_HEADERS) => headers.map((header) => entry[header] ?? "");
function workbook(sheets, format = "xlsx") {
  const wb = XLSX.utils.book_new();
  for (const { name, rows, merges } of sheets) { const ws = XLSX.utils.aoa_to_sheet(rows); if (merges) ws["!merges"] = merges; XLSX.utils.book_append_sheet(wb, ws, name); }
  return XLSX.write(wb, { type: "array", bookType: format });
}
const scan = (entries = [employee()], format = "xlsx") => scanTurnover(workbook([{ name: "Employees", rows: [TURNOVER_HEADERS, ...entries.map((entry) => employeeRow(entry))] }], format), `turnover.${format}`, template);

test("detects shifted/reordered headers and title rows, preserves IDs, and never mutates template", () => {
  const original = JSON.stringify(template);
  const reordered = [...TURNOVER_HEADERS].reverse();
  const steps = [];
  const result = scanTurnover(workbook([{ name: "Staff", rows: [["Monthly turnover report"], [], ["", "", ...reordered], ["", "", ...employeeRow(employee(), reordered)]] }]), "report.xlsx", template, (step, percent) => steps.push(percent));
  assert.equal(result.matched, 1);
  assert.equal(result.employees[0].employee["Emp #"], "0001");
  assert.equal(result.rows[0].onRoll, 1);
  assert.equal(result.rows[1].onRoll, 0);
  assert.equal(result.rows[0].approvedStrength, 8);
  assert.equal(result.rows[1].status, "Inactive");
  assert.equal(JSON.stringify(template), original);
  assert.ok(steps.every((value, index) => index === 0 || value >= steps[index - 1]));
});

test("legacy .xls workbooks are supported", () => {
  assert.equal(scan([employee()], "xls").matched, 1);
});

test("reads multiple worksheets and repeated tables with different column order", () => {
  const reversed = [...TURNOVER_HEADERS].reverse();
  const result = scanTurnover(workbook([
    { name: "Cover", rows: [["Report summary"]] },
    { name: "First", rows: [TURNOVER_HEADERS, employeeRow(employee()), reversed, employeeRow(employee({ "Emp #": "0002" }), reversed)] },
    { name: "Second", rows: [TURNOVER_HEADERS, employeeRow(employee({ Department: "Human Resource", Unit: "KGM", Designation: "OFFICER", Grade: "M-11", "Emp #": "0003" }))] },
  ]), "report.xlsx", template);
  assert.equal(result.matched, 3);
  assert.equal(result.rows[0].onRoll, 2);
  assert.equal(result.rows[2].onRoll, 1);
  assert.ok(result.issues.some((issue) => issue.sheet === "Cover"));
});

test("identical duplicates are counted once and conflicting identities are excluded", () => {
  const duplicate = scan([employee(), employee()]);
  assert.equal(duplicate.matched, 1);
  assert.equal(duplicate.duplicates, 1);
  const conflict = scan([employee(), employee({ Grade: "E-04", Designation: "SUPERVISOR" })]);
  assert.equal(conflict.matched, 0);
  assert.equal(conflict.unmatched, 2);
  assert.equal(conflict.rows[0].onRoll, 0);
  assert.equal(conflict.rows[1].onRoll, 0);
});

test("matching respects grade and department; ambiguity is never guessed", () => {
  const result = scan([employee({ Department: "Unknown", Unit: "Unknown" }), employee({ "Emp #": "0002", Grade: "E-99" })]);
  assert.equal(result.matched, 0);
  assert.equal(result.unmatched, 2);
  const ambiguous = [...templateRows(template), ...templateRows(template)];
  assert.equal(matchEmployee(employee(), ambiguous).match, null);
});

test("malformed records are excluded and explicit merged department cells are supported", () => {
  const result = scanTurnover(workbook([{ name: "Employees", rows: [TURNOVER_HEADERS, employeeRow(employee({ Department: "SPINNING - RING - 1" })), employeeRow(employee({ Department: "", "Emp #": "0002" })), employeeRow(employee({ "Emp #": "", Designation: "", Grade: "" }))], merges: [XLSX.utils.decode_range("A2:A3")] }]), "report.xlsx", template);
  assert.equal(result.matched, 2);
  assert.equal(result.skippedRows, 1);
});

test("rejects unsupported, empty, oversized, disguised, corrupt, missing-header and empty-data uploads", () => {
  assert.throws(() => validateUpload("report.csv", 20), /Excel workbook/);
  assert.throws(() => validateUpload("report.xlsx", 0), /empty/);
  assert.throws(() => validateUpload("report.xlsx", 21 * 1024 * 1024), /20 MB/);
  assert.throws(() => scanTurnover(new TextEncoder().encode("not an Excel file").buffer, "report.xlsx", template), /extension/);
  assert.throws(() => scanTurnover(new Uint8Array([0x50, 0x4b, 3, 4, 0, 0, 0, 0]).buffer, "report.xlsx", template), /corrupted/);
  const missing = TURNOVER_HEADERS.filter((header) => header !== "Grade");
  assert.throws(() => scanTurnover(workbook([{ name: "Data", rows: [missing, employeeRow(employee(), missing)] }]), "report.xlsx", template), /Grade/);
  assert.throws(() => scanTurnover(workbook([{ name: "Data", rows: [TURNOVER_HEADERS] }]), "report.xlsx", template), /No valid employee/);
});

test("partially malformed sections are flagged without losing valid worksheet records", () => {
  const incomplete = TURNOVER_HEADERS.filter((header) => header !== "Joining Date");
  const result = scanTurnover(workbook([{ name: "Data", rows: [TURNOVER_HEADERS, employeeRow(employee()), incomplete, employeeRow(employee({ "Emp #": "0002" }), incomplete)] }]), "report.xlsx", template);
  assert.equal(result.matched, 1);
  assert.ok(result.issues.some((issue) => /Joining Date/.test(issue.message)));
});

test("generated XLSX round-trips with separate categories, formula totals, correct header sequence, styles, and logo", async () => {
  const before = await readFile(new URL("../strength-master-seed-0.json", import.meta.url));
  const result = scan();
  const png = await readFile(new URL("../public/kohinoor-logo.png", import.meta.url));
  const buffer = await createReportWorkbook(result, `data:image/png;base64,${png.toString("base64")}`);
  assert.equal(new Uint8Array(buffer)[0], 0x50);
  assert.equal(new Uint8Array(buffer)[1], 0x4b);
  const wb = new ExcelJS.Workbook();
  await wb.xlsx.load(buffer);
  assert.equal(wb.worksheets.length, 6);
  assert.ok(wb.getWorksheet("RING"));
  assert.ok(wb.getWorksheet("HR   ADMIN"));
  assert.deepEqual(wb.getWorksheet("Employees").getRow(4).values.slice(1, 13), [...TURNOVER_HEADERS]);
  assert.equal(wb.getWorksheet("Strength Detail").getCell("G5").value, 1);
  assert.equal(wb.getWorksheet("Strength Detail").getCell("F5").value, 8);
  assert.equal(wb.getWorksheet("Strength Detail").getCell("H5").value.result, 7);
  assert.equal(wb.getWorksheet("Summary").getCell("D5").value.result, 1);
  assert.equal(wb.getWorksheet("Employees").getCell("C5").value, "0001");
  assert.equal(wb.getWorksheet("RING").views[0].ySplit, 4);
  assert.ok(wb.getWorksheet("RING").autoFilter);
  assert.equal(wb.getWorksheet("RING").getRow(4).getCell(1).font.bold, true);
  assert.equal(wb.getWorksheet("RING").getImages().length, 1);
  assert.deepEqual(await readFile(new URL("../strength-master-seed-0.json", import.meta.url)), before);
});

test("worksheet names are sanitized, bounded and collision-safe", async () => {
  const result = scan();
  result.rows.push({ ...result.rows[0], category: "Summary" }, { ...result.rows[0], category: "x".repeat(40) }, { ...result.rows[0], category: "X".repeat(40) });
  const wb = new ExcelJS.Workbook();
  await wb.xlsx.load(await createReportWorkbook(result));
  const names = wb.worksheets.map((sheet) => sheet.name);
  assert.equal(names.length, new Set(names.map((name) => name.toLowerCase())).size);
  assert.ok(names.every((name) => name.length <= 31 && !/[\\/*?:\[\]]/.test(name)));
});

test("configured categories without designations still receive their own worksheet", async () => {
  const withEmptyCategory = structuredClone(template);
  withEmptyCategory.strengthStructure.push({ category: "EMPTY", subcategories: [] });
  const result = scanTurnover(workbook([{ name: "Data", rows: [TURNOVER_HEADERS, employeeRow(employee())] }]), "report.xlsx", withEmptyCategory);
  const wb = new ExcelJS.Workbook();
  await wb.xlsx.load(await createReportWorkbook(result));
  assert.ok(wb.getWorksheet("EMPTY"));
  assert.equal(wb.getWorksheet("Summary").getCell("C7").value, 0);
});

test("two-row headers are detected without treating header fragments as employees", () => {
  const top = TURNOVER_HEADERS.map((header) => header === "Employee Name" ? "Employee" : header === "Joining Date" ? "Joining" : header);
  const lower = TURNOVER_HEADERS.map((header) => header === "Employee Name" ? "Name" : header === "Joining Date" ? "Date" : "");
  const result = scanTurnover(workbook([{ name: "Data", rows: [top, lower, employeeRow(employee())] }]), "report.xlsx", template);
  assert.equal(result.matched, 1);
  assert.equal(result.scannedRows, 1);
});
