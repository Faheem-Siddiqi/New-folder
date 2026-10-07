import type { Employee, ReportRow } from "./turnover-contracts";

export const normalizeMatch = (value: string) => value.normalize("NFKC").toLowerCase().replace(/[^a-z0-9]/g, "");
const words = (value: string) => value.normalize("NFKC").toLowerCase()
  .replace(/\belectrionics\b/g, "electronics")
  .replace(/\b(asstt|asst)\b/g, "assistant")
  .replace(/\bforman\b/g, "foreman").replace(/\bsr\b/g, "senior")
  .match(/[a-z]+|\d+[a-z]?/g) ?? [];
export const gradeKey = (value: string) => normalizeMatch(value).replace(/\d+/g, (number) => String(Number(number)));
export const designationKey = (value: string) => words(value).sort().join("");

// These equivalents describe the consolidated departments in the root template.
// Keep line numbers and production/maintenance distinctions; never choose by job alone.
const departmentAliases: Record<string, string> = {
  ring5: "ring5production",
  ring2maintenance: "ring2and3maintenance",
  ring3maintenance: "ring2and3maintenance",
  ring23maintenance: "ring2and3maintenance",
  ring2and3maintenance: "ring2and3maintenance",
  compressor: "compressorxorella",
  airconditioning1: "airconditioningcompressor1",
  compressor1: "airconditioningcompressor1",
  electricspinning: "electricspinning",
  electronics: "electronicsspinning",
  electronicsspinning: "electronicsspinning",
  electronics1: "electronics1",
  blowroom1: "blowroom",
  blowroom1maintenance: "blowroom",
  blowroommixing1: "blowroom",
  cardroom1: "cardroom",
  cardroom1maintenance: "cardroom",
  comber1: "comber",
  comber1maintenance: "comber",
  drawingsimplex1: "drawingsimplex",
  drawingsimplex1maintenance: "drawingsimplex",
};
export function departmentKey(value: string) {
  const normalized = normalizeMatch(words(value).join(" ").replace(/\band\b/g, ""));
  const base = normalized.replace(/^engineeringservices/, "").replace(/^spinning/, "");
  return departmentAliases[base] ?? base;
}

export function createEmployeeMatcher(rows: ReportRow[]) {
  const indexed = rows.map((row, index) => ({ row, index }));
  type Entry = (typeof indexed)[number];
  const exactDepartments = new Map<string, Entry[]>();
  const equivalentDepartments = new Map<string, Entry[]>();
  for (const entry of indexed) for (const [map, key] of [[exactDepartments, normalizeMatch(entry.row.subcategory)], [equivalentDepartments, departmentKey(entry.row.subcategory)]] as const) {
    const group = map.get(key) ?? [];
    group.push(entry);
    map.set(key, group);
  }
  const lookup = (map: Map<string, Entry[]>, keys: string[]) => keys.map((key) => map.get(key) ?? []).find((entries) => entries.length) ?? [];
  return (employee: Employee): { match: number | null; issue: string } => {
  const contexts = [employee.Department, `${employee.Department} ${employee.Unit}`, employee.Unit].filter((value) => value.trim());
  // Resolve the department before the designation. Grade is display-only metadata.
  let departments = lookup(exactDepartments, contexts.map(normalizeMatch));
  if (!departments.length) departments = lookup(equivalentDepartments, contexts.map(departmentKey));
  const categoryContexts = [employee.Department, employee["Org Group"]].map(normalizeMatch).filter(Boolean);
  const narrowed = departments.filter(({ row }) => categoryContexts.includes(normalizeMatch(row.category)));
  if (narrowed.length) departments = narrowed;
  if (!departments.length) return { match: null, issue: `Department not found in the strength template: "${employee.Department}" (Unit "${employee.Unit}"). Designation "${employee.Designation}". Add the missing department or correct its source label.` };
  let designations = departments.filter(({ row }) => normalizeMatch(row.designation) === normalizeMatch(employee.Designation));
  if (!designations.length) designations = departments.filter(({ row }) => designationKey(row.designation) === designationKey(employee.Designation));
  if (!designations.length) return { match: null, issue: `Designation "${employee.Designation}" is not configured under ${[...new Set(departments.map(({ row }) => row.subcategory))].join(" / ")}.` };
  const candidates = designations;
  if (candidates.length === 1) return { match: candidates[0].index, issue: "" };
  return { match: null, issue: `Multiple template departments match "${employee.Designation}": ${[...new Set(candidates.map(({ row }) => `${row.category} > ${row.subcategory}`))].join("; ")}. Review the duplicate configuration entries or category.` };
  };
}

export function matchEmployee(employee: Employee, rows: ReportRow[]) {
  return createEmployeeMatcher(rows)(employee);
}

export function sameEmployeeIdentity(first: Employee, next: Employee) {
  // Grade, shift, rest day, payroll and joining-date differences do not change strength.
  return (!first["Employee Name"] || !next["Employee Name"] || normalizeMatch(first["Employee Name"]) === normalizeMatch(next["Employee Name"])) &&
    designationKey(first.Designation) === designationKey(next.Designation) &&
    departmentKey(first.Department) === departmentKey(next.Department);
}
