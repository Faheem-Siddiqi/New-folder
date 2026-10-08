import test from "node:test";
import assert from "node:assert/strict";
import * as XLSX from "xlsx";
import { scanTurnover, TURNOVER_HEADERS } from "./turnover.ts";
import { refreshReport } from "./refresh-report.ts";

const post = (designation = "Operator", approvedStrength = 1) => ({ designation, approvedStrength, grade: "E-03", cadre: "Worker", onRoll: 999, status: "Active" });
const template = () => ({ strengthStructure: [{ category: "Production", subcategories: [{ subcategory: "Ring", designations: [post()] }] }] });
const employee = (id, designation = "Operator", name = `Employee ${id}`, shift = "A") => ["Ring", "KGM", id, name, designation, "M", "Sunday", "M-99", shift, "Worker", "Production", "2024-01-01"];
function scan(entries, tree = template()) {
  const workbook = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(workbook, XLSX.utils.aoa_to_sheet([[...TURNOVER_HEADERS], ...entries]), "Employees");
  return scanTurnover(XLSX.write(workbook, { type: "array", bookType: "xlsx" }), "turnover.xlsx", tree);
}

test("session records rematch employees and updates approvals and duplicate notes without the workbook", () => {
  const previous = scan([employee("001", "Fitter"), employee("001", "Fitter", "Employee 001", "B"), employee("002")]);
  assert.equal(previous.matched, 1);
  const snapshot = JSON.stringify(previous);
  const latest = template();
  latest.strengthStructure[0].subcategories[0].designations[0].approvedStrength = 7;
  latest.strengthStructure[0].subcategories[0].designations.push(post("Fitter", 3));
  const next = refreshReport(previous, latest);
  assert.equal(next.matched, 2);
  assert.equal(next.unmatched, 0);
  assert.equal(next.rows[0].approvedStrength, 7);
  assert.equal(next.rows[0].onRoll, 1);
  assert.equal(next.rows[1].onRoll, 1);
  assert.equal(next.duplicates, 1);
  assert.equal(next.issues.find((issue) => issue.message.startsWith("Repeated employee")).outcome, "counted");
  assert.ok(next.issues.every((issue) => !issue.message.includes("not configured")));
  assert.equal(JSON.stringify(previous), snapshot);
  const removed = refreshReport(next, template());
  assert.equal(removed.matched, 1);
  assert.equal(removed.issues.find((issue) => issue.message.startsWith("Repeated employee")).outcome, "excluded");
});

test("conflicting employee identities stay excluded when refreshing session records", () => {
  const previous = scan([employee("001", "Operator", "First"), employee("001", "Operator", "Second"), employee("002")]);
  const next = refreshReport(previous, template());
  assert.equal(next.matched, 1);
  assert.equal(next.unmatched, 2);
  assert.equal(next.rows[0].onRoll, 1);
  assert.equal(next.issues.filter((issue) => issue.message.startsWith("Conflicting identity")).length, 2);
});

