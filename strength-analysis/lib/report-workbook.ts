import ExcelJS from "exceljs";
import { buildReviewCases, caseIdentity, confirmedAdjustments, vacancyFor, resolutionLabel, reviewColors } from "./review-cases.ts";
import { strengthView } from "./strength-summary.ts";
import { TURNOVER_HEADERS, type TurnoverResult } from "./turnover-contracts.ts";

export async function createReportWorkbook(result: TurnoverResult, logo?: string, progress: (step: string, percent: number) => void = () => {}) {
  const workbook = new ExcelJS.Workbook();
  workbook.creator = "Kohinoor Textile Mills Limited";
  workbook.created = new Date(result.generatedAt);
  workbook.calcProperties.fullCalcOnLoad = true;
  const image = logo ? workbook.addImage({ base64: logo, extension: "png" }) : undefined;
  const names = new Set<string>();
  const sheetName = (name: string) => {
    const base = name.replace(/[\\/*?:\[\]]/g, " ").replace(/^'+|'+$/g, "").trim().slice(0, 31) || "Category";
    let candidate = base;
    let suffix = 2;
    while (names.has(candidate.toLowerCase())) { const ending = ` (${suffix++})`; candidate = base.slice(0, 31 - ending.length) + ending; }
    names.add(candidate.toLowerCase());
    return candidate;
  };
  const makeSheet = (name: string, headers: string[], widths: number[], note: string) => {
    const sheet = workbook.addWorksheet(sheetName(name), { views: [{ state: "frozen", ySplit: 4 }], pageSetup: { orientation: "landscape", paperSize: 9, fitToPage: true, fitToWidth: 1, fitToHeight: 0, printTitlesRow: "1:4" } });
    sheet.columns = headers.map((_, index) => ({ width: widths[index] ?? 22 }));
    sheet.mergeCells(1, 3, 1, Math.max(4, headers.length));
    sheet.getCell(1, 3).value = name;
    sheet.getCell(1, 3).font = { name: "Calibri", size: 18, bold: true, color: { argb: "FF171717" } };
    sheet.getRow(1).height = 32;
    sheet.getRow(2).height = 24;
    sheet.mergeCells(2, 3, 2, Math.max(4, headers.length));
    sheet.getCell(2, 3).value = `Kohinoor Textile Mills Limited | ${result.fileName} | ${new Date(result.generatedAt).toLocaleString("en-GB", { timeZone: "Asia/Karachi" })} PKT`;
    sheet.getCell(2, 3).font = { size: 10, color: { argb: "FF737373" } };
    if (image !== undefined) sheet.addImage(image, { tl: { col: 0, row: 0 }, ext: { width: 224, height: 52 } });
    sheet.mergeCells(3, 1, 3, headers.length);
    sheet.getCell(3, 1).value = note;
    sheet.getCell(3, 1).font = { size: 10, color: { argb: "FF525252" } };
    sheet.getCell(3, 1).alignment = { wrapText: true, vertical: "middle" };
    sheet.getRow(3).height = 30;
    sheet.getRow(4).values = headers;
    sheet.getRow(4).height = 30;
    sheet.getRow(4).eachCell((cell) => { cell.font = { name: "Calibri", bold: true, size: 11, color: { argb: "FFFFFFFF" } }; cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FF262626" } }; cell.alignment = { vertical: "middle", wrapText: true }; });
    sheet.headerFooter.oddFooter = "&LKohinoor Textile Mills Limited&RPage &P of &N";
    return sheet;
  };
  const finish = (sheet: ExcelJS.Worksheet, numeric: number[] = [], vacancyColumn?: number) => {
    sheet.autoFilter = { from: { row: 4, column: 1 }, to: { row: Math.max(4, sheet.rowCount), column: sheet.columnCount } };
    sheet.eachRow((row, index) => {
      if (index < 5) return;
      row.height = 28;
      row.eachCell((cell, column) => {
        cell.font = { name: "Calibri", size: 11, color: { argb: "FF262626" } };
        cell.alignment = { vertical: "middle", wrapText: true, horizontal: numeric.includes(column) ? "right" : "left" };
        if (index % 2) cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FFF5F5F5" } };
        if (numeric.includes(column)) cell.numFmt = "#,##0;[Red](#,##0)";
      });
    });
    if (vacancyColumn && sheet.rowCount >= 5) {
      const column = sheet.getColumn(vacancyColumn).letter;
      sheet.addConditionalFormatting({ ref: `${column}5:${column}${sheet.rowCount}`, rules: [
        { type: "cellIs", operator: "greaterThan", formulae: ["0"], priority: 1, style: { font: { color: { argb: "FF047857" } }, fill: { type: "pattern", pattern: "solid", fgColor: { argb: "FFD1FAE5" } } } },
        { type: "cellIs", operator: "lessThan", formulae: ["0"], priority: 2, style: { font: { color: { argb: "FFB91C1C" } }, fill: { type: "pattern", pattern: "solid", fgColor: { argb: "FFFEE2E2" } } } },
      ] });
    }
  };
  const cases = result.reviewCases ?? buildReviewCases(result, []);
  const tint = (cell: ExcelJS.Cell, kind: keyof typeof reviewColors) => {
    cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: reviewColors[kind].fill } };
    cell.font = { name: "Calibri", size: 11, color: { argb: reviewColors[kind].text } };
  };
  const annotateStrength = (sheet: ExcelJS.Worksheet, rows: typeof result.rows, column: number) => {
    const headers = ["Against Post Cases", "Social Security Cases", "Other Cases", "Saved Cases", "Raw Vacancies"];
    headers.forEach((header, offset) => { sheet.getCell(4, column + offset).value = header; sheet.getColumn(column + offset).width = 25; sheet.getCell(4, column + offset).style = { ...sheet.getCell(4, 1).style }; });
    rows.forEach((row, index) => {
      const rowIndex = result.rows.indexOf(row);
      const related = cases.filter((entry) => entry.rowIndex === rowIndex && entry.resolution);
      (["Against Post", "Social Security Leave", "Other"] as const).forEach((kind, offset) => { const count = related.filter((entry) => entry.resolution?.caseType === kind).length; sheet.getCell(index + 5, column + offset).value = count; if (count) tint(sheet.getCell(index + 5, column + offset), kind); });
      const saved = confirmedAdjustments({ ...result, reviewCases: cases }, rowIndex);
      sheet.getCell(index + 5, column + 3).value = { formula: `SUM(${sheet.getColumn(column).letter}${index + 5}:${sheet.getColumn(column + 2).letter}${index + 5})`, result: saved };
      sheet.getCell(index + 5, column + 4).value = row.approvedStrength - row.onRoll;
      sheet.getCell(index + 5, column - 1).value = { formula: `${sheet.getColumn(column - 3).letter}${index + 5}-${sheet.getColumn(column - 2).letter}${index + 5}+${sheet.getColumn(column + 3).letter}${index + 5}`, result: vacancyFor({ ...result, reviewCases: cases }, rowIndex) };
    });
    sheet.autoFilter = { from: { row: 4, column: 1 }, to: { row: Math.max(4, sheet.rowCount), column: column + 4 } };
  };
  const view = strengthView(result);
  const summary = makeSheet("Summary", ["Subcategory", "Approved Strength", "On-Roll Strength"], [52, 25, 28], "Admin & HR: subcategory totals. On-roll counts use matched employees from the uploaded turnover report.");
  const details = makeSheet("Strength Detail", ["Category", "Subcategory", "Designation", "Grade", "Cadre", "Approved Strength", "On-Roll Employee Count", "Vacancies"], [35, 48, 38, 14, 16, 23, 27, 18], "Designation metadata and approved strength come from the Strength JSON. On-roll counts come only from uniquely matched uploaded employees. Vacancies include saved designation cases. Approved strength and uploaded on-roll counts remain unchanged; raw vacancies are shown separately.");
  const employeeSheet = makeSheet("Employees", [...TURNOVER_HEADERS, "Template Category", "Template Subcategory", "Template Designation", "Template Grade", "Template Cadre", "Template Approved Strength", "Mapping Result", "Source Sheet", "Source Row"], [38, 18, 17, 32, 32, 13, 15, 13, 16, 20, 24, 20, 32, 45, 35, 16, 18, 26, 64, 26, 14], "The first 12 columns preserve uploaded employee information in the required sequence. Template fields are shown separately; the JSON does not contain employee names or IDs. Repeated employee number/unit records count once; differences in shift or rest day do not exclude employees.");
  const validation = makeSheet("Validation", ["Source Sheet", "Source Row", "Emp #", "Issue", "Effect on Count"], [30, 16, 20, 100, 24], "Employees matching department and designation are counted; uploaded grade is not checked. Repeated employee number/unit records count once. Only conflicting identity or strength fields exclude duplicates. Each note explains whether a record is counted or excluded.");
  result.rows.forEach((row, index) => details.addRow([row.category, row.subcategory, row.designation, row.grade, row.cadre, row.approvedStrength, row.onRoll, { formula: `F${index + 5}-G${index + 5}`, result: row.approvedStrength - row.onRoll }]));
  finish(details, [6, 7, 8], 8);
  annotateStrength(details, result.rows, 9);
  const categories = [...new Set([...result.categories, ...result.rows.map((row) => row.category)])];
  categories.forEach((category, categoryIndex) => {
    progress(`Building ${category} worksheet`, 15 + Math.round(categoryIndex / categories.length * 40));
    const categoryRows = result.rows.filter((row) => row.category === category);
    const sheet = makeSheet(category, ["Subcategory", "Designation", "Grade", "Cadre", "Approved Strength", "On-Roll Employee Count", "Vacancies"], [48, 38, 14, 16, 24, 28, 18], "Vacancies = approved minus uploaded on-roll plus saved designation cases. Raw vacancies and each case type are shown separately.");
    categoryRows.forEach((row, index) => sheet.addRow([row.subcategory, row.designation, row.grade, row.cadre, row.approvedStrength, row.onRoll, { formula: `E${index + 5}-F${index + 5}`, result: row.approvedStrength - row.onRoll }]));
    finish(sheet, [5, 6, 7], 7);
    annotateStrength(sheet, categoryRows, 8);
  });
  const endRow = Math.max(5, result.rows.length + 4);
  const quoted = (value: string) => `"${value.replace(/~/g, "~~").replace(/\*/g, "~*").replace(/\?/g, "~?").replace(/"/g, '""')}"`;
  const addSummaryRows = (rows: typeof view.admin, subcategories: boolean) => {
    const first = summary.rowCount + 1;
    for (const row of rows) {
      const criteria = `'Strength Detail'!A5:A${endRow},${quoted(row.category)}${subcategories ? `,'Strength Detail'!B5:B${endRow},${quoted(row.name)}` : ""}`;
      summary.addRow([row.name,
        { formula: `SUMIFS('Strength Detail'!F5:F${endRow},${criteria})`, result: row.approved },
        { formula: `SUMIFS('Strength Detail'!G5:G${endRow},${criteria})`, result: row.onRoll }]);
    }
    const last = summary.rowCount;
    const total = summary.addRow(["Total", { formula: rows.length ? `SUM(B${first}:B${last})` : "0", result: rows.reduce((sum, row) => sum + row.approved, 0) }, { formula: rows.length ? `SUM(C${first}:C${last})` : "0", result: rows.reduce((sum, row) => sum + row.onRoll, 0) }]);
    return total.number;
  };
  const adminTotal = addSummaryRows(view.admin, true);
  const spinningTitle = adminTotal + 2;
  summary.mergeCells(spinningTitle, 1, spinningTitle, 3);
  summary.getCell(spinningTitle, 1).value = "Spinning";
  const spinningHeader = summary.getRow(spinningTitle + 1);
  spinningHeader.values = ["Category", "Approved Strength", "On-Roll Strength"];
  const spinningTotal = addSummaryRows(view.spinning, false);
  finish(summary, [2, 3]);
  summary.autoFilter = undefined;
  for (const rowNumber of [adminTotal, spinningTotal, spinningTitle]) {
    summary.getRow(rowNumber).eachCell((cell) => { cell.font = { name: "Calibri", bold: true, size: 12 }; cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FFE5E5E5" } }; });
  }
  spinningHeader.eachCell((cell) => { cell.font = { name: "Calibri", bold: true, size: 11, color: { argb: "FFFFFFFF" } }; cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FF262626" } }; });
  progress("Adding employee records and validation notes", 65);
  for (const record of result.employees) {
    const match = record.match === null ? undefined : result.rows[record.match];
    employeeSheet.addRow([...TURNOVER_HEADERS.map((header) => record.employee[header]), match?.category ?? "", match?.subcategory ?? "", match?.designation ?? "", match?.grade ?? "", match?.cadre ?? "", match?.approvedStrength ?? "", match ? "Matched" : record.issue, record.sheet, record.row]);
  }
  finish(employeeSheet, [18, 21]);
  const adjustments = makeSheet("Review Cases Adjustments", ["Category", "Subcategory", "Designation", "Case Number", "Case Type", "Against Post / Other Reason", "Resolution", "Previous Suggestion", "Saved At"], [32, 42, 38, 18, 25, 45, 24, 45, 26], "Designation-level cases. Each saved case offsets one negative vacancy. Blue: Against Post. Brown: Social Security. Grey: Other. Uploaded employee counts and approved strength are unchanged.");
  cases.forEach((entry) => {
    const identity = caseIdentity(result, entry.rowIndex, entry.caseNumber);
    adjustments.addRow([identity.category, identity.subcategory, identity.designation, identity.caseNumber, entry.resolution?.caseType === "Social Security Leave" ? "Social Security" : entry.resolution?.caseType ?? "Pending", entry.resolution?.assignedAgainstDesignation ?? entry.resolution?.otherReason ?? "", entry.resolution ? "Saved" : "Awaiting case", entry.suggestion ? resolutionLabel(entry.suggestion) : "", entry.resolution?.resolvedAt ?? ""]);
  });
  if (!cases.length) adjustments.addRow(["", "", "No negative vacancies require review."]);
  finish(adjustments, [4]);
  cases.forEach((entry, index) => { const row = adjustments.getRow(index + 5); row.height = 42; if (entry.resolution) row.eachCell((cell) => tint(cell, entry.resolution!.caseType)); });
  result.issues.forEach((issue) => validation.addRow([issue.sheet, issue.row || "", issue.employeeId, issue.message, issue.outcome === "counted" ? "Counted once" : issue.outcome === "excluded" ? "Excluded from on-roll" : "Review note"]));
  if (!result.issues.length) validation.addRow(["", "", "", "No validation issues found."]);
  finish(validation, [2]);
  progress("Writing valid Excel workbook", 85);
  const bytes = await workbook.xlsx.writeBuffer();
  return new Uint8Array(bytes).buffer;
}
