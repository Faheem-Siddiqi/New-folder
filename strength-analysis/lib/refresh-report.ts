import type { StrengthData } from "./strength-data";
import { strengthRows } from "./strength-rows.ts";
import { createEmployeeMatcher, normalizeMatch, sameEmployeeIdentity } from "./employee-matching.ts";
import { MATCHING_VERSION, type TurnoverResult } from "./turnover-contracts.ts";

export function refreshReport(previous: TurnoverResult, template: StrengthData): TurnoverResult {
  const rows = strengthRows(template, true);
  if (!rows.length) throw new Error("The strength template has no designations. Add them in Configuration before refreshing the report.");
  const match = createEmployeeMatcher(rows);
  const identity = (record: TurnoverResult["employees"][number]) => `${normalizeMatch(record.employee.Unit)}|${record.employee["Emp #"].toLowerCase()}`;
  const groups = new Map<string, typeof previous.employees>();
  for (const record of previous.employees) {
    const key = identity(record);
    const group = groups.get(key) ?? [];
    group.push(record);
    groups.set(key, group);
  }
  const conflicts = new Set([...groups].filter(([, group]) => {
    const names = new Set(group.map((record) => normalizeMatch(record.employee["Employee Name"])).filter(Boolean));
    return names.size > 1 || group.some((record) => !sameEmployeeIdentity(group[0].employee, record.employee) || record.issue.startsWith("Conflicting identity or strength fields"));
  }).map(([key]) => key));
  const employees = previous.employees.map((record) => ({ ...record, employee: { ...record.employee }, ...(conflicts.has(identity(record))
    ? { match: null, issue: record.issue || "Conflicting identity or strength fields for the same employee number/unit. Correct the source workbook and upload it again." }
    : match(record.employee)) }));
  const byEmployee = new Map<string, typeof employees>();
  const oldIssues = new Map(previous.employees.map((record) => [JSON.stringify([record.sheet, record.row]), record.issue]));
  for (const record of employees) {
    const records = byEmployee.get(record.employee["Emp #"]) ?? [];
    records.push(record);
    byEmployee.set(record.employee["Emp #"], records);
  }
  // Keep source/header/skipped-row notes; regenerate notes that depended on matching.
  const oldMappingNotes = new Set(previous.employees.filter((record) => record.issue).map((record) => JSON.stringify([record.sheet, record.row, record.issue])));
  const issues = previous.issues.filter((issue) => !oldMappingNotes.has(JSON.stringify([issue.sheet, issue.row, issue.message]))).map((issue) => {
    if (!issue.message.startsWith("Repeated employee number/unit")) return { ...issue };
    const first = (byEmployee.get(issue.employeeId) ?? []).find((record) => {
      const oldIssue = oldIssues.get(JSON.stringify([record.sheet, record.row]));
      return issue.message.includes(`${record.sheet} row ${record.row}`) || oldIssue && issue.message.endsWith(oldIssue);
    });
    if (!first) return { ...issue };
    const counted = first.match !== null;
    return { ...issue, severity: counted ? "info" as const : "warning" as const, outcome: counted ? "counted" as const : "excluded" as const, message: counted
      ? `Repeated employee number/unit: counted once from ${first.sheet} row ${first.row}. Grade, shift, rest day, and other employee-detail differences do not exclude this employee.`
      : `Repeated employee number/unit; its first record is excluded. ${first.issue}` };
  });
  let matched = 0;
  for (const record of employees) {
    if (record.match !== null) { rows[record.match].onRoll++; matched++; }
    else issues.push({ sheet: record.sheet, row: record.row, employeeId: record.employee["Emp #"], message: record.issue, severity: "error", outcome: "excluded" });
  }
  return { ...previous, matchingVersion: MATCHING_VERSION, generatedAt: new Date().toISOString(), templateSignature: JSON.stringify(template), categories: template.strengthStructure.map((category) => category.category), rows, employees, issues, matched, unmatched: employees.length - matched, reviewCases: undefined };
}
