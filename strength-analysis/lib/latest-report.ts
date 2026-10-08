import { readFile, writeFile, rename, unlink } from "node:fs/promises";
import { resolve } from "node:path";
import { randomUUID } from "node:crypto";
import { buildReviewCases, scenarioKey, validResolution, type Resolution } from "./review-cases.ts";
import { TURNOVER_HEADERS, type TurnoverResult } from "./turnover-contracts.ts";
import { refreshReport } from "./refresh-report.ts";
import type { StrengthData } from "./strength-data.ts";

export const LATEST_REPORT_FILE = resolve(process.cwd(), "..", "strength-last-result.json");
let queue: Promise<unknown> = Promise.resolve();
function serialized<T>(operation: () => Promise<T>): Promise<T> {
  const job = queue.catch(() => {}).then(operation);
  queue = job;
  return job;
}
function validateResult(value: TurnoverResult) {
  if (!value || value.version !== 1 || typeof value.fileName !== "string" || typeof value.templateSignature !== "string" || !Number.isFinite(Date.parse(value.generatedAt)) ||
    !Array.isArray(value.categories) || !value.categories.every((name) => typeof name === "string") ||
    !Array.isArray(value.rows) || !value.rows.every((row) => row && [row.category, row.subcategory, row.designation, row.grade, row.cadre, row.status].every((item) => typeof item === "string") && [row.approvedStrength, row.onRoll].every((item) => Number.isSafeInteger(item) && item >= 0)) ||
    !Array.isArray(value.employees) || !value.employees.every((record) => record?.employee && TURNOVER_HEADERS.every((header) => typeof record.employee[header] === "string") && (record.match === null || Number.isInteger(record.match) && record.match >= 0 && record.match < value.rows.length) && typeof record.sheet === "string" && typeof record.issue === "string" && Number.isInteger(record.row)) ||
    !Array.isArray(value.issues) || !value.issues.every((issue) => issue && typeof issue.sheet === "string" && typeof issue.employeeId === "string" && typeof issue.message === "string" && Number.isInteger(issue.row)) || ![value.scannedRows, value.duplicates, value.skippedRows, value.matched, value.unmatched].every((item) => Number.isSafeInteger(item) && item >= 0)) throw new Error("Invalid latest report. Upload the turnover workbook again.");
  const template = JSON.parse(value.templateSignature);
  if (!Array.isArray(template?.strengthStructure) || !template.strengthStructure.every((category: { category?: unknown; subcategories?: unknown }) => typeof category?.category === "string" && Array.isArray(category.subcategories) && category.subcategories.every((section: { subcategory?: unknown; designations?: unknown }) => typeof section?.subcategory === "string" && Array.isArray(section.designations)))) throw new Error("Invalid saved strength snapshot.");
  if (value.reportId !== undefined && (typeof value.reportId !== "string" || !value.reportId || !Number.isSafeInteger(value.revision) || value.revision! < 0)) throw new Error("Invalid report identity.");
  if (value.reviewCases !== undefined && (!Array.isArray(value.reviewCases) || value.reviewCases.some((entry) => !entry || !Number.isInteger(entry.rowIndex) || !value.rows[entry.rowIndex] || !Number.isSafeInteger(entry.caseNumber) || entry.caseNumber < 1 || entry.key !== scenarioKey({ ...value.rows[entry.rowIndex], caseNumber: entry.caseNumber }) || [entry.resolution, entry.suggestion].some((item) => item && (!validResolution(item) || scenarioKey(item) !== entry.key))))) throw new Error("Invalid saved report cases.");
}
export async function readLatestReport(path = LATEST_REPORT_FILE): Promise<TurnoverResult | null> {
  try {
    const saved = JSON.parse(await readFile(path, "utf8"));
    if (saved.version !== 1 || !("result" in saved)) throw new Error("Invalid latest report file");
    if (saved.result === null) return null;
    validateResult(saved.result);
    return saved.result;
  } catch (error) {
    if ((error as NodeJS.ErrnoException).code === "ENOENT") return null;
    throw new Error("strength-last-result.json could not be read. The existing file has been preserved; repair it before saving another report.");
  }
}
async function writeLatestReport(result: TurnoverResult, path: string) {
  const temporary = `${path}.${randomUUID()}.tmp`;
  try {
    await writeFile(temporary, `${JSON.stringify({ version: 1, updatedAt: new Date().toISOString(), result }, null, 2)}\n`, { encoding: "utf8", flag: "wx" });
    await rename(temporary, path);
  } catch { throw new Error("The latest report could not be saved. Check that strength-last-result.json is writable and retry."); }
  finally { await unlink(temporary).catch(() => {}); }
}
export function previousCases(report: TurnoverResult, previous: TurnoverResult | null): Resolution[] {
  if (!previous) return [];
  const oldRows = new Map(previous.rows.map((row) => [scenarioKey({ ...row, caseNumber: 1 }), row]));
  const eligible = new Set(report.rows.filter((row) => {
    const old = oldRows.get(scenarioKey({ ...row, caseNumber: 1 }));
    return old && row.approvedStrength - row.onRoll < 0 && row.approvedStrength - row.onRoll === old.approvedStrength - old.onRoll;
  }).map((row) => scenarioKey({ ...row, caseNumber: 1 })));
  return (previous.reviewCases ?? []).flatMap((entry) => {
    const choice = entry.resolution ?? entry.suggestion;
    return choice && validResolution(choice) && eligible.has(scenarioKey({ ...choice, caseNumber: 1 })) ? [choice] : [];
  });
}
export function saveProcessedReport(report: TurnoverResult, expectedReportId?: string, path = LATEST_REPORT_FILE) {
  return serialized(async () => {
    validateResult(report);
    const previous = await readLatestReport(path);
    if (expectedReportId && previous?.reportId !== expectedReportId) throw new Error("A newer turnover report is available. Refresh before reprocessing.");
    const next: TurnoverResult = { ...report, reportId: randomUUID(), revision: 0, reviewCases: buildReviewCases(report, previousCases(report, previous)) };
    await writeLatestReport(next, path);
    return next;
  });
}
export function saveLatestCase(resolution: Resolution, reportId: string, revision: number, path = LATEST_REPORT_FILE) {
  return serialized(async () => {
    const latest = await readLatestReport(path);
    if (!latest || latest.reportId !== reportId || latest.revision !== revision) throw new Error("The latest report changed. Refresh before saving this case.");
    if (!validResolution(resolution)) throw new Error("Invalid case assignment.");
    const key = scenarioKey(resolution);
    const cases = buildReviewCases(latest, [], latest.reviewCases ?? []);
    if (!cases.some((entry) => entry.key === key)) throw new Error("This negative vacancy no longer requires that case.");
    const updated = cases.map((entry) => entry.key === key ? { ...entry, resolution } : entry);
    const suggestions = new Map((latest.reviewCases ?? []).map((entry) => [entry.key, entry.suggestion]));
    const next = { ...latest, revision: revision + 1, reviewCases: buildReviewCases(latest, [], updated).map((entry) => {
      const suggestion = suggestions.get(entry.key);
      return { ...entry, ...(suggestion ? { suggestion } : {}) };
    }) };
    if (!next.reviewCases.find((entry) => entry.key === key)?.resolution) throw new Error("Choose a valid post in the same subcategory.");
    await writeLatestReport(next, path);
    return next;
  });
}

export function refreshLatestReport(template: StrengthData, expectedReportId: string, path = LATEST_REPORT_FILE) {
  return serialized(async () => {
    const previous = await readLatestReport(path);
    if (!previous || previous.reportId !== expectedReportId) throw new Error("The saved report changed. Reload Home before refreshing it.");
    const refreshed = refreshReport(previous, template);
    const next: TurnoverResult = { ...refreshed, reportId: randomUUID(), revision: 0, reviewCases: buildReviewCases(refreshed, previousCases(refreshed, previous)) };
    await writeLatestReport(next, path);
    return next;
  });
}
