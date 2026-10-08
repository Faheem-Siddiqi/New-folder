import { buildReviewCases, scenarioKey, validResolution, type Resolution } from "./review-cases.ts";
import { TURNOVER_HEADERS, type TurnoverResult } from "./turnover-contracts.ts";

export const SESSION_REPORT_KEY = "strength-analysis:session-report:v1";
export function readSessionReport(): TurnoverResult | null {
  const contents = sessionStorage.getItem(SESSION_REPORT_KEY);
  if (!contents) return null;
  try {
    const result = JSON.parse(contents) as TurnoverResult;
    if (result.version !== 1 || !result.reportId || typeof result.fileName !== "string" || typeof result.templateSignature !== "string" || !Number.isSafeInteger(result.revision) || !Array.isArray(result.categories) || !Array.isArray(result.rows) || !Array.isArray(result.issues) || !Array.isArray(result.employees) || !result.rows.every((row) => row && [row.category, row.subcategory, row.designation, row.grade, row.cadre, row.status].every((value) => typeof value === "string") && [row.approvedStrength, row.onRoll].every((value) => Number.isSafeInteger(value) && value >= 0)) || !result.employees.every((record) => record?.employee && TURNOVER_HEADERS.every((header) => typeof record.employee[header] === "string") && (record.match === null || Number.isInteger(record.match) && record.match >= 0 && record.match < result.rows.length))) throw new Error();
    JSON.parse(result.templateSignature);
    result.reviewCases = buildReviewCases(result, result.reviewCases ?? []);
    return result;
  } catch { throw new Error("The report in this tab could not be restored. Upload the workbook again."); }
}
export function storeSessionReport(result: TurnoverResult) {
  try { sessionStorage.setItem(SESSION_REPORT_KEY, JSON.stringify(result)); }
  catch { throw new Error("This tab could not save the report in session storage. Storage may be full or disabled. Try a smaller workbook or allow session storage."); }
}
export function createSessionReport(result: TurnoverResult): TurnoverResult {
  return { ...result, reportId: crypto.randomUUID(), revision: 0, reviewCases: buildReviewCases(result) };
}
export function updateSessionCase(report: TurnoverResult, resolution: Resolution): TurnoverResult {
  if (!validResolution(resolution)) throw new Error("Choose a valid case assignment.");
  const key = scenarioKey(resolution);
  const cases = buildReviewCases(report, report.reviewCases ?? []);
  if (!cases.some((entry) => entry.key === key)) throw new Error("This negative vacancy no longer requires a case.");
  const updated = cases.map((entry) => entry.key === key ? { ...entry, resolution } : entry);
  const next = { ...report, revision: (report.revision ?? 0) + 1, reviewCases: buildReviewCases(report, updated) };
  if (!next.reviewCases.find((entry) => entry.key === key)?.resolution) throw new Error("Choose a valid post in the same subcategory.");
  return next;
}
