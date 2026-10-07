import { strengthRows } from "./strength-rows.ts";
import type { StrengthData } from "./strength-data";
import type { ReportRow, TurnoverResult } from "./turnover-contracts";

export type StrengthTotal = { name: string; category: string; approved: number; onRoll: number; designations: number };
const key = (value: string) => value.toLowerCase().replace(/[^a-z0-9]/g, "");
export const isAdminCategory = (category: string) => ["hradmin", "adminhr", "hrandadmin", "adminandhr"].includes(key(category));
const spinningNames: Record<string, string> = { backprocess: "Back Process", ring: "Ring", autoconemachconerc: "Autocone / Machcone / RC", lab: "Spinning General / Lab", spinninggenerallab: "Spinning General / Lab", services: "Services" };
const order = ["Back Process", "Ring", "Autocone / Machcone / RC", "Spinning General / Lab", "Services"];
export function sumStrength(rows: ReportRow[]) {
  return { approved: rows.reduce((sum, row) => sum + row.approvedStrength, 0), onRoll: rows.reduce((sum, row) => sum + row.onRoll, 0), designations: rows.length };
}
export function strengthView(data: StrengthData | TurnoverResult) {
  const report = "rows" in data ? data : null;
  const template: StrengthData = report ? JSON.parse(report.templateSignature) : data as StrengthData;
  const rows: ReportRow[] = report ? report.rows : strengthRows(template);
  const categories = [...new Set([...(report?.categories ?? template.strengthStructure.map((category) => category.category)), ...rows.map((row) => row.category)])];
  const detailed = categories.map((category) => ({ category, rows: rows.filter((row) => row.category === category), ...sumStrength(rows.filter((row) => row.category === category)) }));
  const admin: StrengthTotal[] = [];
  const spinning: StrengthTotal[] = [];
  for (const detail of detailed) {
    if (isAdminCategory(detail.category)) {
      const subcategories = [...new Set([...(template.strengthStructure.find((category) => category.category === detail.category)?.subcategories.map((subcategory) => subcategory.subcategory) ?? []), ...detail.rows.map((row) => row.subcategory)])];
      for (const subcategory of subcategories) admin.push({ name: subcategory, category: detail.category, ...sumStrength(detail.rows.filter((row) => row.subcategory === subcategory)) });
    } else spinning.push({ name: spinningNames[key(detail.category)] ?? detail.category, category: detail.category, ...sumStrength(detail.rows) });
  }
  spinning.sort((a, b) => (order.includes(a.name) ? order.indexOf(a.name) : 99) - (order.includes(b.name) ? order.indexOf(b.name) : 99));
  return { rows, detailed, admin, spinning, total: sumStrength(rows) };
}
