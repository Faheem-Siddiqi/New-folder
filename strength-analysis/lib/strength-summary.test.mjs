import test from "node:test";
import assert from "node:assert/strict";
import { readFile } from "node:fs/promises";
import ExcelJS from "exceljs";
import * as XLSX from "xlsx";
import { strengthView } from "./strength-summary.ts";
import { STRENGTH_FILE } from "./strength-file.ts";
import { scanTurnover, TURNOVER_HEADERS } from "./turnover.ts";
import { createReportWorkbook } from "./report-workbook.ts";

test("management summary includes every HR subcategory and the five spinning categories", async () => {
  const template = JSON.parse(await readFile(STRENGTH_FILE, "utf8"));
  const view = strengthView(template);
  const admin = template.strengthStructure.find((category) => category.category === "HR & ADMIN");
  assert.deepEqual(view.admin.map((row) => row.name), admin.subcategories.map((row) => row.subcategory));
  assert.deepEqual(view.spinning.map((row) => row.name), ["Back Process", "Ring", "Autocone / Machcone / RC", "Spinning General / Lab", "Services"]);
  assert.equal([...view.admin, ...view.spinning].reduce((sum, row) => sum + row.approved, 0), view.total.approved);
  assert.equal([...view.admin, ...view.spinning].reduce((sum, row) => sum + row.onRoll, 0), view.total.onRoll);
});

test("updated hierarchy and approval recalculate UI and Excel while uploaded employees determine on-roll", async () => {
  const designation = { designation: "Officer", grade: "M-11", cadre: "Staff", approvedStrength: 2, onRoll: 999, status: "Active" };
  const template = { strengthStructure: [{ category: "HR & ADMIN", subcategories: [{ subcategory: "Finance", designations: [designation] }] }] };
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet([TURNOVER_HEADERS, ["Finance", "KGM", "001", "Employee", "Officer", "M", "Sunday", "M-11", "A", "Staff", "HR & ADMIN", "2024-01-01"]]), "Employees");
  const bytes = XLSX.write(wb, { type: "array", bookType: "xlsx" });
  const previous = scanTurnover(bytes, "turnover.xlsx", template);
  const latest = structuredClone(template);
  latest.strengthStructure[0].subcategories[0].designations[0].approvedStrength = 7;
  latest.strengthStructure[0].subcategories.push({ subcategory: "New finance section", designations: [{ ...designation, designation: "Manager", approvedStrength: 3 }] }, { subcategory: "Empty section", designations: [] });
  latest.strengthStructure.push({ category: "SERVICES", subcategories: [{ subcategory: "Workshop", designations: [{ ...designation, designation: "Fitter", approvedStrength: 5 }] }] });
  const original = JSON.stringify(latest);
  const result = scanTurnover(bytes, "turnover.xlsx", latest);
  const view = strengthView(result);
  assert.equal(strengthView(previous).total.approved, 2);
  assert.deepEqual(view.total, { approved: 15, onRoll: 1, designations: 3 });
  assert.deepEqual(view.admin.map((row) => [row.name, row.approved, row.onRoll]), [["Finance", 7, 1], ["New finance section", 3, 0], ["Empty section", 0, 0]]);
  assert.deepEqual(view.spinning.map((row) => [row.name, row.approved, row.onRoll]), [["Services", 5, 0]]);
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.load(await createReportWorkbook(result));
  assert.equal(workbook.getWorksheet("Summary").getCell("B5").value.result, 7);
  assert.equal(workbook.getWorksheet("Summary").getCell("C5").value.result, 1);
  assert.equal(workbook.getWorksheet("Summary").getCell("B8").value.result, 10);
  assert.equal(workbook.getWorksheet("SERVICES").getCell("E5").value, 5);
  assert.equal(JSON.stringify(latest), original);
});
