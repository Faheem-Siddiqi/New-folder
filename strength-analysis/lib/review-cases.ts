import type { TurnoverResult, ReportRow } from "./turnover-contracts.ts";

export type Resolution = { id: string; category: string; subcategory: string; designation: string; caseNumber: number; caseType: "Against Post" | "Social Security Leave" | "Other"; assignedAgainstDesignation: string | null; otherReason?: string; resolvedAt: string };
export type ReviewCase = { key: string; rowIndex: number; caseNumber: number; resolution?: Resolution; suggestion?: Resolution };
const normalized = (value: string) => value.normalize("NFKC").trim().toLowerCase().replace(/\s+/g, " ");
export function scenarioKey(value: Pick<Resolution, "category" | "subcategory" | "designation" | "caseNumber">) {
  return JSON.stringify([normalized(value.category), normalized(value.subcategory), normalized(value.designation), value.caseNumber]);
}
export function caseIdentity(report: TurnoverResult, rowIndex: number, caseNumber: number) {
  const row = report.rows[rowIndex];
  if (!row || !Number.isSafeInteger(caseNumber) || caseNumber < 1) throw new Error("This designation case is no longer valid. Reprocess the upload.");
  return { category: row.category, subcategory: row.subcategory, designation: row.designation, caseNumber };
}
export function againstPostOptions(rows: ReportRow[], row: ReportRow) {
  return [...new Set(rows.filter((target) => target.category === row.category && target.subcategory === row.subcategory && target.designation !== row.designation).map((target) => target.designation))];
}
export function validResolution(value: unknown): value is Resolution {
  if (!value || typeof value !== "object") return false;
  const record = value as Resolution;
  return [record.id, record.category, record.subcategory, record.designation, record.resolvedAt].every((field) => typeof field === "string" && field.length > 0 && field.length <= 500)
    && Number.isSafeInteger(record.caseNumber) && record.caseNumber > 0 && Number.isFinite(Date.parse(record.resolvedAt))
    && (record.caseType === "Social Security Leave" ? record.assignedAgainstDesignation === null : record.caseType === "Other" ? record.assignedAgainstDesignation === null && typeof record.otherReason === "string" && !!record.otherReason.trim() && record.otherReason.length <= 200 : record.caseType === "Against Post" && typeof record.assignedAgainstDesignation === "string" && !!record.assignedAgainstDesignation && record.assignedAgainstDesignation.length <= 500);
}
// Preserve older employee-level history without treating it as a designation adjustment.
export function validHistoryRecord(value: unknown) {
  if (validResolution(value)) return true;
  if (!value || typeof value !== "object") return false;
  const record = value as Record<string, unknown>;
  return record.caseNumber === undefined && ["id", "employeeId", "unit", "employeeName", "category", "subcategory", "designation", "resolvedAt"].every((field) => typeof record[field] === "string") && !!record.employeeId && Number.isFinite(Date.parse(record.resolvedAt as string)) && (record.caseType === "Social Security Leave" ? record.assignedAgainstDesignation === null : record.caseType === "Against Post" && typeof record.assignedAgainstDesignation === "string");
}
export function resolutionAllowed(rows: ReportRow[], row: ReportRow, resolution: Resolution) {
  return validResolution(resolution) && (resolution.caseType !== "Against Post" || againstPostOptions(rows, row).includes(resolution.assignedAgainstDesignation ?? ""));
}
export function buildReviewCases(report: TurnoverResult, history: Resolution[], confirmed: ReviewCase[] = []): ReviewCase[] {
  const previous = new Map(history.filter(validResolution).map((entry) => [scenarioKey(entry), entry]));
  const current = new Map((Array.isArray(confirmed) ? confirmed : []).filter((entry) => entry && typeof entry.key === "string" && entry.resolution && validResolution(entry.resolution)).map((entry) => [entry.key, entry.resolution!]));
  const cases: ReviewCase[] = [];
  report.rows.forEach((row, rowIndex) => {
    const excess = Math.max(0, row.onRoll - row.approvedStrength);
    for (let caseNumber = 1; caseNumber <= excess; caseNumber++) {
      const key = scenarioKey(caseIdentity(report, rowIndex, caseNumber));
      const resolution = current.get(key);
      const suggestion = previous.get(key);
      cases.push({ key, rowIndex, caseNumber, ...(resolution && scenarioKey(resolution) === key && resolutionAllowed(report.rows, row, resolution) ? { resolution } : {}), ...(suggestion && resolutionAllowed(report.rows, row, suggestion) ? { suggestion } : {}) });
    }
  });
  return cases;
}
export function confirmedAdjustments(report: TurnoverResult, rowIndex?: number) {
  const keys = new Set<string>();
  for (const entry of report.reviewCases ?? []) {
    const row = report.rows[entry.rowIndex];
    if (row && (rowIndex === undefined || entry.rowIndex === rowIndex) && entry.caseNumber <= Math.max(0, row.onRoll - row.approvedStrength) && entry.resolution && scenarioKey(entry.resolution) === entry.key && scenarioKey(caseIdentity(report, entry.rowIndex, entry.caseNumber)) === entry.key && resolutionAllowed(report.rows, row, entry.resolution)) keys.add(entry.key);
  }
  return keys.size;
}
export function vacancyFor(report: TurnoverResult, rowIndex: number) {
  const row = report.rows[rowIndex];
  return row.approvedStrength - row.onRoll + confirmedAdjustments(report, rowIndex);
}
export function resolutionLabel(resolution: Resolution) {
  return resolution.caseType === "Against Post" ? `Against Post ? ${resolution.assignedAgainstDesignation}` : resolution.caseType === "Other" ? `Other ? ${resolution.otherReason}` : "Social Security";
}
export const reviewColors = { "Against Post": { fill: "FFDBEAFE", text: "FF1D4ED8", className: "bg-blue-50 text-blue-800 border-blue-200" }, "Social Security Leave": { fill: "FFEEE5DA", text: "FF785B42", className: "bg-[#eee5da] text-[#785b42] border-[#d7c6b2]" }, "Other": { fill: "FFF3F4F6", text: "FF4B5563", className: "bg-neutral-100 text-neutral-700 border-neutral-200" } };
