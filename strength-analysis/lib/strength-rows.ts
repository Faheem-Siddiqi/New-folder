import type { StrengthData } from "./strength-data";
import type { ReportRow } from "./turnover-contracts";
import { designationKey, gradeKey } from "./employee-matching.ts";

export function strengthRows(template: StrengthData, uploaded = false): ReportRow[] {
  const rows: ReportRow[] = [];
  const equivalent = new Map<string, ReportRow>();
  const key = (value: string) => value.normalize("NFKC").toLowerCase().replace(/[^a-z0-9]/g, "");
  const mergeValues = (first: string, next: string, normalize: (value: string) => string) => [...new Map([first, next].flatMap((value) => value.split(" / ")).filter(Boolean).map((value) => [normalize(value), value])).values()].join(" / ");
  for (const category of template.strengthStructure) for (const subcategory of category.subcategories) for (const designation of subcategory.designations) {
    const identity = [key(category.category), key(subcategory.subcategory), designationKey(designation.designation)].join("|");
    const existing = equivalent.get(identity);
    if (existing) {
      existing.approvedStrength += designation.approvedStrength;
      if (!uploaded) existing.onRoll += designation.onRoll;
      existing.grade = mergeValues(existing.grade, designation.grade, gradeKey);
      existing.cadre = mergeValues(existing.cadre, designation.cadre, key);
    }
    else {
      const row = { ...designation, category: category.category, subcategory: subcategory.subcategory, onRoll: uploaded ? 0 : designation.onRoll };
      rows.push(row);
      equivalent.set(identity, row);
    }
  }
  return rows;
}
