import ExcelJS from "exceljs";
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
    if (image !== undefined) sheet.addImage(image, { tl: { col: 0, row: 0 }, ext: { width: 110, height: 52 } });
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
  const finish = (sheet: ExcelJS.Worksheet, numeric: number[] = []) => {
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
  };
  const summary = makeSheet("Summary", ["Category", "Designations", "Approved Strength", "On-Roll Employee Count", "Vacancies"], [42, 20, 24, 27, 22], `Matched employees: ${result.matched} | Unmatched: ${result.unmatched} | Duplicate rows: ${result.duplicates} | Invalid rows: ${result.skippedRows}. See Validation before using totals.`);
  const details = makeSheet("Strength Detail", ["Category", "Subcategory", "Designation", "Grade", "Cadre", "Approved Strength", "On-Roll Employee Count", "Vacancies"], [35, 48, 38, 14, 16, 23, 27, 18], "Designation metadata and approved strength come from the Strength JSON. On-roll counts come only from uniquely matched uploaded employees. All template rows are retained.");
  const employeeSheet = makeSheet("Employees", [...TURNOVER_HEADERS, "Template Category", "Template Subcategory", "Template Designation", "Template Grade", "Template Cadre", "Template Approved Strength", "Mapping Result", "Source Sheet", "Source Row"], [38, 18, 17, 32, 32, 13, 15, 13, 16, 20, 24, 20, 32, 45, 35, 16, 18, 26, 64, 26, 14], "The first 12 columns preserve uploaded employee information in the required sequence. Template fields are shown separately; the JSON does not contain employee names or IDs. Conflicting duplicates are excluded from counts.");
  const validation = makeSheet("Validation", ["Source Sheet", "Source Row", "Emp #", "Issue"], [30, 16, 20, 100], "Unmatched and conflicting records do not contribute to on-roll counts. Identical duplicate employee number/unit records are counted once. Review excluded sections and records in the original upload.");
  result.rows.forEach((row, index) => details.addRow([row.category, row.subcategory, row.designation, row.grade, row.cadre, row.approvedStrength, row.onRoll, { formula: `F${index + 5}-G${index + 5}`, result: row.approvedStrength - row.onRoll }]));
  finish(details, [6, 7, 8]);
  const categories = [...new Set([...result.categories, ...result.rows.map((row) => row.category)])];
  categories.forEach((category, categoryIndex) => {
    progress(`Building ${category} worksheet`, 15 + Math.round(categoryIndex / categories.length * 40));
    const categoryRows = result.rows.filter((row) => row.category === category);
    const sheet = makeSheet(category, ["Subcategory", "Designation", "Grade", "Cadre", "Approved Strength", "On-Roll Employee Count", "Vacancies"], [48, 38, 14, 16, 24, 28, 18], "All configured designations are included. On-roll counts are calculated from matched employees in this upload; negative vacancies indicate strength above approval.");
    categoryRows.forEach((row, index) => sheet.addRow([row.subcategory, row.designation, row.grade, row.cadre, row.approvedStrength, row.onRoll, { formula: `E${index + 5}-F${index + 5}`, result: row.approvedStrength - row.onRoll }]));
    finish(sheet, [5, 6, 7]);
    const approved = categoryRows.reduce((sum, row) => sum + row.approvedStrength, 0);
    const onRoll = categoryRows.reduce((sum, row) => sum + row.onRoll, 0);
    const ref = `'${sheet.name.replace(/'/g, "''")}'`;
    summary.addRow([category, categoryRows.length, categoryRows.length ? { formula: `SUM(${ref}!E5:E${categoryRows.length + 4})`, result: approved } : 0, categoryRows.length ? { formula: `SUM(${ref}!F5:F${categoryRows.length + 4})`, result: onRoll } : 0, { formula: `C${categoryIndex + 5}-D${categoryIndex + 5}`, result: approved - onRoll }]);
  });
  finish(summary, [2, 3, 4, 5]);
  progress("Adding employee records and validation notes", 65);
  for (const record of result.employees) {
    const match = record.match === null ? undefined : result.rows[record.match];
    employeeSheet.addRow([...TURNOVER_HEADERS.map((header) => record.employee[header]), match?.category ?? "", match?.subcategory ?? "", match?.designation ?? "", match?.grade ?? "", match?.cadre ?? "", match?.approvedStrength ?? "", match ? "Matched" : record.issue, record.sheet, record.row]);
  }
  finish(employeeSheet, [18, 21]);
  result.issues.forEach((issue) => validation.addRow([issue.sheet, issue.row || "", issue.employeeId, issue.message]));
  if (!result.issues.length) validation.addRow(["", "", "", "No validation issues found."]);
  finish(validation, [2]);
  progress("Writing valid Excel workbook", 85);
  const bytes = await workbook.xlsx.writeBuffer();
  return new Uint8Array(bytes).buffer;
}
