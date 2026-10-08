import test from "node:test";
import assert from "node:assert/strict";
import { mkdtemp, readFile, writeFile, rm } from "node:fs/promises";
import { tmpdir } from "node:os";
import { join } from "node:path";
import { saveProcessedReport, saveLatestCase, readLatestReport } from "./latest-report.ts";
import { caseIdentity, vacancyFor } from "./review-cases.ts";

const report = () => ({ version: 1, generatedAt: "2026-10-08T10:00:00Z", fileName: "turnover.xlsx", templateSignature: JSON.stringify({ strengthStructure: [] }), categories: ["RING"], rows: [
  { category: "RING", subcategory: "Spinning", designation: "Labour", grade: "E-01", cadre: "Worker", status: "Active", approvedStrength: 1, onRoll: 4 },
  { category: "RING", subcategory: "Spinning", designation: "Fitter", grade: "E-03", cadre: "Worker", status: "Active", approvedStrength: 3, onRoll: 0 },
], employees: [], issues: [], scannedRows: 4, matched: 4, unmatched: 0, duplicates: 0, skippedRows: 0 });
const choice = (result, number, extra = {}) => ({ ...caseIdentity(result, 0, number), id: `case-${number}`, caseType: "Against Post", assignedAgainstDesignation: "Fitter", resolvedAt: "2026-10-08T11:00:00Z", ...extra });
async function withFile(work) {
  const directory = await mkdtemp(join(tmpdir(), "latest-strength-"));
  try { await work(join(directory, "strength-last-result.json")); }
  finally { await rm(directory, { recursive: true, force: true }); }
}

test("upload, case addition, edit and approval persist the entire latest report", () => withFile(async (path) => {
  let latest = await saveProcessedReport(report(), undefined, path);
  assert.deepEqual(await readLatestReport(path), latest);
  latest = await saveLatestCase(choice(latest, 1), latest.reportId, latest.revision, path);
  latest = await saveLatestCase(choice(latest, 2, { caseType: "Social Security Leave", assignedAgainstDesignation: null }), latest.reportId, latest.revision, path);
  latest = await saveLatestCase(choice(latest, 3, { caseType: "Other", assignedAgainstDesignation: null, otherReason: "Temporary duty" }), latest.reportId, latest.revision, path);
  assert.equal(vacancyFor(await readLatestReport(path), 0), 0);
  latest = await saveLatestCase(choice(latest, 1, { id: "edited", caseType: "Other", assignedAgainstDesignation: null, otherReason: "Updated reason" }), latest.reportId, latest.revision, path);
  assert.equal((await readLatestReport(path)).reviewCases[0].resolution.otherReason, "Updated reason");
  const next = await saveProcessedReport(report(), undefined, path);
  assert.ok(next.reviewCases.every((entry) => entry.suggestion && !entry.resolution));
  assert.equal(vacancyFor(next, 0), -3);
  assert.deepEqual(next.reviewCases.map((entry) => entry.suggestion.caseType), ["Other", "Social Security Leave", "Other"]);
  // Another upload before approval must retain the pending suggestions.
  const again = await saveProcessedReport(report(), undefined, path);
  assert.ok(again.reviewCases.every((entry) => entry.suggestion && !entry.resolution));
  const approved = await saveLatestCase({ ...again.reviewCases[0].suggestion, id: "approved" }, again.reportId, again.revision, path);
  assert.equal(vacancyFor(await readLatestReport(path), 0), -2);
  assert.equal(approved.reviewCases[1].suggestion.caseType, "Social Security Leave");
}));

test("only identical raw negative vacancy and designation identity reuse the previous choices", () => withFile(async (path) => {
  let latest = await saveProcessedReport(report(), undefined, path);
  latest = await saveLatestCase(choice(latest, 1), latest.reportId, latest.revision, path);
  for (const field of ["category", "subcategory", "designation", "onRoll"]) {
    const changed = report();
    changed.rows[0][field] = field === "onRoll" ? 3 : "Different";
    const next = await saveProcessedReport(changed, undefined, path);
    assert.ok(next.reviewCases.every((entry) => !entry.suggestion));
    latest = await saveProcessedReport(report(), undefined, path);
    latest = await saveLatestCase(choice(latest, 1), latest.reportId, latest.revision, path);
  }
  const positive = report(); positive.rows[0].onRoll = 0;
  assert.equal((await saveProcessedReport(positive, undefined, path)).reviewCases.length, 0);
  assert.ok((await saveProcessedReport(report(), undefined, path)).reviewCases.every((entry) => !entry.suggestion));
}));

test("stale operations, invalid posts and corrupt files never replace the latest result", () => withFile(async (path) => {
  let latest = await saveProcessedReport(report(), undefined, path);
  const stale = latest;
  latest = await saveLatestCase(choice(latest, 1), latest.reportId, latest.revision, path);
  await assert.rejects(saveLatestCase(choice(stale, 2), stale.reportId, stale.revision, path), /changed/);
  await assert.rejects(saveLatestCase(choice(latest, 2, { assignedAgainstDesignation: "Missing" }), latest.reportId, latest.revision, path), /valid post/);
  assert.deepEqual(await readLatestReport(path), latest);
  const newer = await saveProcessedReport(report(), undefined, path);
  await assert.rejects(saveProcessedReport(report(), latest.reportId, path), /newer turnover/);
  assert.deepEqual(await readLatestReport(path), newer);
  await writeFile(path, "{broken");
  await assert.rejects(saveProcessedReport(report(), undefined, path), /could not be read/);
  assert.equal(await readFile(path, "utf8"), "{broken");
}));
