import test from "node:test";
import assert from "node:assert/strict";
import { mkdtemp, readFile, writeFile, rm } from "node:fs/promises";
import { tmpdir } from "node:os";
import { join } from "node:path";
import ExcelJS from "exceljs";
import * as XLSX from "xlsx";
import { buildReviewCases, caseIdentity, againstPostOptions, resolutionAllowed, vacancyFor } from "./review-cases.ts";
import { appendAdjustment, readAdjustmentHistory } from "./adjustment-history.ts";
import { createReportWorkbook } from "./report-workbook.ts";
import { TURNOVER_HEADERS } from "./turnover-contracts.ts";

const row = { category: "BACKPROCESS", subcategory: "Spinning-BlowRoom", designation: "MIXING SORTER", grade: "E-03", cadre: "Worker", approvedStrength: 1, onRoll: 4, status: "Active" };
const makeReport = () => ({ version: 1, generatedAt: "2026-10-07T10:00:00Z", fileName: "sample.xlsx", templateSignature: JSON.stringify({ strengthStructure: [] }), categories: ["BACKPROCESS"], rows: [{ ...row }, { ...row, designation: "ASSISTANT MANAGER", onRoll: 0 }, { ...row, subcategory: "Another Department", designation: "OTHER POST", onRoll: 0 }], employees: ["001", "002", "003", "004"].map((id, index) => ({ employee: Object.fromEntries(TURNOVER_HEADERS.map((header) => [header, header === "Emp #" ? id : header === "Employee Name" ? `Employee ${id}` : header === "Unit" ? "1" : ""])), match: 0, sheet: "Employees", row: index + 2, issue: "" })), issues: [], scannedRows: 4, matched: 4, unmatched: 0, duplicates: 0, skippedRows: 0 });
const resolution = (report, caseNumber, overrides = {}) => ({ ...caseIdentity(report, 0, caseNumber), id: `resolution-${caseNumber}`, caseType: "Against Post", assignedAgainstDesignation: "ASSISTANT MANAGER", resolvedAt: "2026-10-07T11:00:00Z", ...overrides });

test("negative three creates three designation cases, with no employee selection", () => {
  const report = makeReport();
  const cases = buildReviewCases(report, []);
  assert.deepEqual(cases.map(entry => entry.caseNumber), [1,2,3]);
  assert.ok(cases.every(entry => entry.rowIndex === 0 && !("recordIndex" in entry)));
  assert.deepEqual(againstPostOptions(report.rows, row), ["ASSISTANT MANAGER"]);
  assert.equal(resolutionAllowed(report.rows, row, resolution(report, 1, { assignedAgainstDesignation: "OTHER POST" })), false);
  assert.equal(resolutionAllowed(report.rows, row, resolution(report, 1, { caseType: "Social Security Leave", assignedAgainstDesignation: null })), true);
  assert.equal(resolutionAllowed(report.rows, row, resolution(report, 1, { caseType: "Other", assignedAgainstDesignation: null, otherReason: " " })), false);
  assert.equal(resolutionAllowed(report.rows, row, resolution(report, 1, { caseType: "Other", assignedAgainstDesignation: null, otherReason: "Temporary duty" })), true);
});

test("one Against Post, one Social Security and one Other clear -3 to zero without editing strength or headcount", () => {
  const report = makeReport();
  report.reviewCases = buildReviewCases(report, []);
  assert.equal(vacancyFor(report,0), -3);
  report.reviewCases[0].resolution = resolution(report,1);
  assert.equal(vacancyFor(report,0), -2);
  report.reviewCases[1].resolution = resolution(report,2,{caseType:"Social Security Leave",assignedAgainstDesignation:null});
  assert.equal(vacancyFor(report,0), -1);
  report.reviewCases[2].resolution = resolution(report,3,{caseType:"Other",assignedAgainstDesignation:null,otherReason:"Temporary duty"});
  assert.equal(vacancyFor(report,0), 0);
  assert.equal(report.rows[0].onRoll,4);
  assert.equal(report.rows[0].approvedStrength,1);
  assert.equal(report.employees.length,4);
});

test("history is suggested by designation and case number; only same-upload confirmations carry forward", () => {
  const report = makeReport();
  const saved = resolution(report,1);
  let cases = buildReviewCases(report,[saved]);
  assert.deepEqual(cases[0].suggestion,saved);
  assert.equal(cases[0].resolution,undefined);
  report.employees[0].employee["Emp #"]="NEW EMPLOYEE";
  report.rows[0].grade="M-99";
  assert.deepEqual(buildReviewCases(report,[saved])[0].suggestion,saved);
  cases[0].resolution=saved;
  assert.equal(buildReviewCases(report,[saved],cases)[0].resolution.id,saved.id);
  assert.equal(buildReviewCases(report,[saved])[0].resolution,undefined);
  const removed = {...report,rows:[report.rows[0],report.rows[2]]};
  assert.equal(buildReviewCases(removed,[saved],cases)[0].resolution,undefined);
  assert.equal(buildReviewCases(removed,[saved],cases)[0].suggestion,undefined);
  assert.ok(buildReviewCases(report,[{...saved,subcategory:"Another Department"}]).every(entry=>!entry.suggestion));
  report.rows[0].approvedStrength=3;
  assert.equal(buildReviewCases(report,[saved],cases).length,1);
});

test("separate history preserves all resolutions, serializes writes, retries safely and preserves corruption", async () => {
  const directory=await mkdtemp(join(tmpdir(),"strength-cases-"));
  const path=join(directory,"strength-adjustment.json");
  try {
    const report=makeReport(); const first=resolution(report,1); const second=resolution(report,2,{caseType:"Other",assignedAgainstDesignation:null,otherReason:"Temporary duty"});
    await Promise.all([appendAdjustment(first,path),appendAdjustment(second,path)]);
    assert.equal((await readAdjustmentHistory(path)).length,2);
    await appendAdjustment(first,path);
    assert.equal((await readAdjustmentHistory(path)).length,2);
    await assert.rejects(appendAdjustment({...first,assignedAgainstDesignation:"CHANGED"},path),/already been used/);
    await writeFile(path,"{broken");
    await assert.rejects(appendAdjustment(first,path),/history could not be read/);
    assert.equal(await readFile(path,"utf8"),"{broken");
  } finally { await rm(directory,{recursive:true,force:true}); }
});

test("Excel includes all saved case explanations beside adjusted vacancies with only category and summary sheets", async () => {
  const report=makeReport();
  report.reviewCases=buildReviewCases(report,[]);
  report.reviewCases[0].resolution=resolution(report,1);
  report.reviewCases[1].resolution=resolution(report,2,{caseType:"Social Security Leave",assignedAgainstDesignation:null});
  report.reviewCases[2].resolution=resolution(report,3,{caseType:"Other",assignedAgainstDesignation:null,otherReason:"Temporary duty"});
  const snapshot=JSON.stringify(report);
  const workbook=new ExcelJS.Workbook();
  const bytes = await createReportWorkbook(report);
  await workbook.xlsx.load(bytes);
  const cells = XLSX.read(bytes, {type:"array"});
  assert.deepEqual(workbook.worksheets.map((sheet) => sheet.name), ["Summary", "BACKPROCESS"]);
  const detail=workbook.getWorksheet("BACKPROCESS");
  assert.equal(detail.columnCount,8);
  assert.equal(detail.getCell("F8").value,4);
  assert.equal(cells.Sheets.BACKPROCESS.G8.v,0);
  assert.equal(detail.getCell("G8").value.formula,"E8-F8+3");
  assert.match(detail.getCell("H8").value,/Against Post/);
  assert.match(detail.getCell("H8").value,/Social Security/);
  assert.match(detail.getCell("H8").value,/Other.*Temporary duty/);
  assert.ok(detail.getRow(8).height >= 58);
  assert.equal(JSON.stringify(report),snapshot);
});

test("Excel shows unresolved excess as pending cases and keeps negative vacancies", async () => {
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.load(await createReportWorkbook(makeReport()));
  const sheet = workbook.getWorksheet("BACKPROCESS");
  assert.equal(sheet.getCell("G8").value.result, -3);
  assert.equal(sheet.getCell("H8").value, "Case 1: Pending review\nCase 2: Pending review\nCase 3: Pending review");
  assert.equal(sheet.getCell("H9").value, "");
});
