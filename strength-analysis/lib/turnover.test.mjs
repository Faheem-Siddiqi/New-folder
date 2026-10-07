import test from "node:test";
import assert from "node:assert/strict";
import { readFile } from "node:fs/promises";
import * as XLSX from "xlsx";
import ExcelJS from "exceljs";
import { scanTurnover, TURNOVER_HEADERS, validateUpload, matchEmployee, templateRows } from "./turnover.ts";
import { createReportWorkbook } from "./report-workbook.ts";
import { STRENGTH_FILE } from "./strength-file.ts";

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

test("apprentice designation rows are ignored in validation and all employee counts", () => {
  const result = scan([
    employee({ Designation: "  apprentice  ", Department: "Unknown", Unit: "Unknown" }),
    employee({ Designation: "APPRENTICE", "Emp #": "0002", "Employee Name": "" }),
    employee(),
  ]);
  assert.equal(result.scannedRows, 1);
  assert.equal(result.matched, 1);
  assert.equal(result.unmatched, 0);
  assert.equal(result.duplicates, 0);
  assert.equal(result.skippedRows, 0);
  assert.equal(result.rows[0].onRoll, 1);
  assert.equal(result.employees.length, 1);
  assert.deepEqual(result.issues, []);
});

test("an apprentice-only workbook produces a zero-count report without validation notes", () => {
  const result = scan([employee({ Designation: "ApPrEnTiCe", "Emp #": "0002" })]);
  assert.equal(result.scannedRows, 0);
  assert.equal(result.matched, 0);
  assert.equal(result.unmatched, 0);
  assert.equal(result.rows[0].onRoll, 0);
  assert.equal(result.employees.length, 0);
  assert.deepEqual(result.issues, []);
});

test("matching respects department, ignores grade, and does not guess ambiguous departments", () => {
  const result = scan([employee({ Department: "Unknown", Unit: "Unknown" }), employee({ "Emp #": "0002", Grade: "E-99" })]);
  assert.equal(result.matched, 1);
  assert.equal(result.unmatched, 1);
  assert.equal(result.employees[1].employee.Grade, "E-99");
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
  const missing = TURNOVER_HEADERS.filter((header) => header !== "Emp #");
  assert.throws(() => scanTurnover(workbook([{ name: "Data", rows: [missing, employeeRow(employee(), missing)] }]), "report.xlsx", template), /Emp #/);
  assert.throws(() => scanTurnover(workbook([{ name: "Data", rows: [TURNOVER_HEADERS] }]), "report.xlsx", template), /No valid employee/);
});

test("partially malformed sections are flagged without losing valid worksheet records", () => {
  const incomplete = TURNOVER_HEADERS.filter((header) => header !== "Joining Date");
  const result = scanTurnover(workbook([{ name: "Data", rows: [TURNOVER_HEADERS, employeeRow(employee()), incomplete, employeeRow(employee({ "Emp #": "0002" }), incomplete)] }]), "report.xlsx", template);
  assert.equal(result.matched, 1);
  assert.ok(result.issues.some((issue) => /Joining Date/.test(issue.message)));
});

test("generated XLSX round-trips with separate categories, formula totals, correct header sequence, styles, and logo", async () => {
  const before = await readFile(STRENGTH_FILE);
  const result = scan();
  const png = await readFile(new URL("../public/kohinoor-logo.png", import.meta.url));
  const buffer = await createReportWorkbook(result, `data:image/png;base64,${png.toString("base64")}`);
  assert.equal(new Uint8Array(buffer)[0], 0x50);
  assert.equal(new Uint8Array(buffer)[1], 0x4b);
  const wb = new ExcelJS.Workbook();
  await wb.xlsx.load(buffer);
  assert.equal(wb.worksheets.length, 7);
  assert.ok(wb.getWorksheet("RING"));
  assert.ok(wb.getWorksheet("HR   ADMIN"));
  assert.deepEqual(wb.getWorksheet("Employees").getRow(6).values.slice(1, 13), [...TURNOVER_HEADERS]);
  assert.equal(wb.getWorksheet("Strength Detail").getCell("G7").value, 1);
  assert.equal(wb.getWorksheet("Strength Detail").getCell("F7").value, 8);
  assert.equal(wb.getWorksheet("Strength Detail").getCell("H7").value.result, 7);
  assert.equal(wb.getWorksheet("Summary").getCell("H9").value.result, 1);
  assert.equal(wb.getWorksheet("Employees").getCell("C7").value, "0001");
  assert.ok(wb.worksheets.every(sheet => sheet.views.every(view => view.state !== "frozen" && !view.xSplit && !view.ySplit)));
  for (const sheet of wb.worksheets) { assert.equal(sheet.getRow(4).hasValues, false); assert.equal(sheet.getRow(5).hasValues, false); }
  assert.ok(wb.getWorksheet("RING").autoFilter);
  assert.equal(wb.getWorksheet("RING").getRow(6).getCell(1).font.bold, true);
  assert.equal(wb.getWorksheet("RING").getImages().length, 1);
  for (const sheet of wb.worksheets) {
    const headers = sheet.getRow(6).values.slice(1);
    assert.ok(headers.every((header) => !/status/i.test(String(header))), `${sheet.name} should not export status`);
    sheet.eachRow((row) => row.eachCell((cell) => assert.ok(!["Active", "Inactive"].includes(cell.value))));
  }
  assert.deepEqual(await readFile(STRENGTH_FILE), before);
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
  assert.equal(wb.getWorksheet("Summary").getCell("F9").value, "EMPTY");
});

test("two-row headers are detected without treating header fragments as employees", () => {
  const top = TURNOVER_HEADERS.map((header) => header === "Employee Name" ? "Employee" : header === "Joining Date" ? "Joining" : header);
  const lower = TURNOVER_HEADERS.map((header) => header === "Employee Name" ? "Name" : header === "Joining Date" ? "Date" : "");
  const result = scanTurnover(workbook([{ name: "Data", rows: [top, lower, employeeRow(employee())] }]), "report.xlsx", template);
  assert.equal(result.matched, 1);
  assert.equal(result.scannedRows, 1);
});

test("consolidated departments match without crossing production lines even when grades differ", () => {
  const departments = ["RING-5 - PRODUCTION", "SPINNING - RING - 2 + 3 MAINTENANCE", "AUTO CONE - MAINTENANCE-7", "AUTO CONE - PRODUCTION-7", "ELECTRIONICS SPINNING", "ELECTRIONICS-1", "ELECTRIC SPINNING", "AIR CONDITIONING&COMPRESSOR-1", "COMPRESSOR/XORELLA", "SPINNING - CARD ROOM", "SPINNING - DRAWING & SIMPLEX"];
  const tree = { strengthStructure: [{ category: "Spinning", subcategories: departments.map((subcategory) => ({ subcategory, designations: [{ designation: "FITTER", grade: "E-03", cadre: "Worker", approvedStrength: 4, onRoll: 0, status: "Active" }] })) }] };
  const rows = templateRows(tree);
  const pairs = [
    ["Spinning - Ring - 5", "RING-5 - PRODUCTION"],
    ["Spinning - Ring - 2 Maintenance", "SPINNING - RING - 2 + 3 MAINTENANCE"],
    ["Spinning - Ring - 3 Maintenance", "SPINNING - RING - 2 + 3 MAINTENANCE"],
    ["SPINNING - AUTO CONE - MAINTENANCE - 7", "AUTO CONE - MAINTENANCE-7"],
    ["SPINNING - AUTO CONE - PRODUCTION - 7", "AUTO CONE - PRODUCTION-7"],
    ["Engineering Services - Electronics", "ELECTRIONICS SPINNING"],
    ["Engineering Services - Electronics 1", "ELECTRIONICS-1"],
    ["Engineering Services - Electric - Spinning", "ELECTRIC SPINNING"],
    ["Spinning - Air Conditioning 1", "AIR CONDITIONING&COMPRESSOR-1"],
    ["~COMPRESSOR 1", "AIR CONDITIONING&COMPRESSOR-1"],
    ["COMPRESSOR", "COMPRESSOR/XORELLA"],
    ["Spinning - Card Room - 1 Maintenance", "SPINNING - CARD ROOM"],
    ["Spinning - Drawing & Simplex 1", "SPINNING - DRAWING & SIMPLEX"],
  ];
  for (const [Department, expected] of pairs) {
    const result = matchEmployee(employee({ Department, Unit: "KGM", Designation: "Fitter", Grade: "E-3" }), rows);
    assert.notEqual(result.match, null, Department);
    assert.equal(rows[result.match].subcategory, expected);
  }
  assert.equal(matchEmployee(employee({ Department: "Spinning - Ring - 6", Unit: "KGM", Designation: "Fitter" }), rows).match, null);
  assert.equal(matchEmployee(employee({ Department: "Spinning - Ring - 4 Maintenance", Unit: "KGM", Designation: "Fitter" }), rows).match, null);
  const wrongGrade = matchEmployee(employee({ Department: "Spinning - Ring - 5", Unit: "KGM", Designation: "Fitter", Grade: "E-04" }), rows);
  assert.notEqual(wrongGrade.match, null);
  assert.equal(rows[wrongGrade.match].subcategory, "RING-5 - PRODUCTION");
  assert.equal(wrongGrade.issue, "");
});

test("common designation abbreviations and word order match within the resolved department", () => {
  const rows = templateRows(template);
  rows[0].designation = "ASSISTANT FOREMAN";
  assert.equal(matchEmployee(employee({ Designation: "ASSTT. FORMAN" }), rows).match, 0);
  rows[0].designation = "INCHARGE MAINTENANCE";
  assert.equal(matchEmployee(employee({ Designation: "MAINTENANCE INCHARGE" }), rows).match, 0);
  assert.equal(matchEmployee(employee({ Department: "Unconfigured workshop", Unit: "KGM", Designation: "MAINTENANCE INCHARGE" }), rows).match, null);
});

test("shift, payroll and rest-day differences count duplicate employees once", () => {
  const result = scan([employee(), employee({ Shift: "B", "Rest Day": "Friday", "Pay Sheet": "Staff", "Joining Date": "2025-01-01" })]);
  assert.equal(result.matched, 1);
  assert.equal(result.unmatched, 0);
  assert.equal(result.duplicates, 1);
  assert.equal(result.rows[0].onRoll, 1);
  assert.equal(result.issues[0].outcome, "counted");
  assert.equal(result.issues[0].severity, "info");
});

test("later identity conflicts invalidate the group including earlier duplicate notes", () => {
  const result = scan([employee(), employee({ Shift: "B" }), employee({ "Employee Name": "Another Person" }), employee()]);
  assert.equal(result.matched, 0);
  assert.equal(result.rows[0].onRoll, 0);
  assert.ok(result.issues.every((issue) => issue.outcome === "excluded"));
  const initiallyBlank = scan([employee({ "Employee Name": "" }), employee(), employee({ "Employee Name": "Another Person" })]);
  assert.equal(initiallyBlank.matched, 0);
});

test("equivalent template duplicates combine approvals and count each employee once without modifying JSON", () => {
  const tree = structuredClone(template);
  tree.strengthStructure[0].subcategories[0].designations.push({ ...tree.strengthStructure[0].subcategories[0].designations[0], approvedStrength: 3 });
  const before = JSON.stringify(tree);
  const result = scanTurnover(workbook([{ name: "Data", rows: [TURNOVER_HEADERS, employeeRow(employee())] }]), "turnover.xlsx", tree);
  assert.equal(result.matched, 1);
  assert.equal(result.rows.length, 3);
  assert.equal(result.rows[0].approvedStrength, 11);
  assert.equal(result.rows[0].onRoll, 1);
  assert.equal(JSON.stringify(tree), before);
});

test("Excel vacancy formulas retain positive, negative and zero values with sign-based highlights", async () => {
  const result = scan();
  result.rows[0].approvedStrength = 0;
  result.rows[2].approvedStrength = 0;
  const wb = new ExcelJS.Workbook();
  await wb.xlsx.load(await createReportWorkbook(result));
  const detail = wb.getWorksheet("Strength Detail");
  assert.equal(detail.getCell("H7").value.result, -1);
  assert.equal(detail.getCell("H8").value.result, 1);
  const reread = XLSX.read(await wb.xlsx.writeBuffer(), { type: "buffer" });
  assert.equal(reread.Sheets["Strength Detail"].H9.v, 0);
  for (const [sheet, ref] of [[detail, "H7:H9"], [wb.getWorksheet("RING"), "G7:G8"]]) {
    assert.equal(sheet.conditionalFormattings[0].ref, ref);
    const rules = sheet.conditionalFormattings[0].rules;
    assert.equal(rules.length, 2);
    assert.equal(rules[0].operator, "greaterThan");
    assert.deepEqual(rules[0].formulae, ["0"]);
    assert.equal(rules[0].style.font.color.argb, "FF047857");
    assert.equal(rules[1].operator, "lessThan");
    assert.deepEqual(rules[1].formulae, ["0"]);
    assert.equal(rules[1].style.font.color.argb, "FFB91C1C");
  }
});

test("different, blank and missing grades are accepted and retained only as employee metadata", async () => {
  const result = scan([employee({ Grade: "M-99" }), employee({ "Emp #": "0002", Grade: "" })]);
  assert.equal(result.matched, 2);
  assert.equal(result.unmatched, 0);
  assert.equal(result.skippedRows, 0);
  assert.equal(result.rows[0].onRoll, 2);
  assert.equal(result.rows[0].grade, "E-03");
  assert.deepEqual(result.employees.map((record) => record.employee.Grade), ["M-99", ""]);
  const exported = new ExcelJS.Workbook();
  await exported.xlsx.load(await createReportWorkbook(result));
  assert.equal(exported.getWorksheet("Employees").getCell("H7").value, "M-99");
  assert.equal(exported.getWorksheet("Employees").getCell("P7").value, "E-03");
  assert.equal(exported.getWorksheet("Strength Detail").getCell("G7").value, 2);
  const headers = TURNOVER_HEADERS.filter((header) => header !== "Grade");
  const missing = scanTurnover(workbook([{ name: "Data", rows: [headers, employeeRow(employee(), headers)] }]), "no-grade.xlsx", template);
  assert.equal(missing.matched, 1);
  assert.equal(missing.employees[0].employee.Grade, "");
  const repeated = scanTurnover(workbook([{ name: "Data", rows: [[...TURNOVER_HEADERS, "Grade"], [...employeeRow(employee()), "M-99"]] }]), "extra-grade.xlsx", template);
  assert.equal(repeated.matched, 1);
  assert.equal(repeated.employees[0].employee.Grade, "E03");
  assert.equal(repeated.issues[0].severity, "info");
});

test("grade-only differences between duplicate employees do not create conflicts", () => {
  const result = scan([employee(), employee({ Grade: "M-99" }), employee({ Grade: "" })]);
  assert.equal(result.matched, 1);
  assert.equal(result.duplicates, 2);
  assert.equal(result.unmatched, 0);
  assert.ok(result.issues.every((issue) => issue.outcome === "counted"));
});

test("the same template designation across grades is one strength group with summed approvals", () => {
  const tree = structuredClone(template);
  tree.strengthStructure[0].subcategories[0].designations.push({ ...tree.strengthStructure[0].subcategories[0].designations[0], grade: "M-11", cadre: "Staff", approvedStrength: 3 });
  const original = JSON.stringify(tree);
  const result = scanTurnover(workbook([{ name: "Data", rows: [TURNOVER_HEADERS, employeeRow(employee({ Grade: "E-99" }))] }]), "mixed-grades.xlsx", tree);
  assert.equal(result.matched, 1);
  assert.equal(result.rows.length, 3);
  assert.equal(result.rows[0].approvedStrength, 11);
  assert.equal(result.rows[0].onRoll, 1);
  assert.equal(result.rows[0].grade, "E-03 / M-11");
  assert.equal(result.rows[0].cadre, "Worker / Staff");
  assert.equal(JSON.stringify(tree), original);
});
