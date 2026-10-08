import ExcelJS from "exceljs";
import { buildReviewCases, confirmedAdjustments, vacancyFor, resolutionLabel } from "./review-cases.ts";
import { strengthView } from "./strength-summary.ts";
import { type TurnoverResult } from "./turnover-contracts.ts";

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
    const sheet = workbook.addWorksheet(sheetName(name), { views: [{ state: "normal", showGridLines: false }], pageSetup: { orientation: "landscape", paperSize: 9, fitToPage: true, fitToWidth: 1, fitToHeight: 0, printTitlesRow: "6:6" } });
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
    sheet.getRow(4).height = 14;
    sheet.getRow(5).height = 14;
    sheet.getRow(6).values = headers;
    sheet.getRow(6).height = 30;
    sheet.getRow(6).eachCell((cell) => { cell.font = { name: "Calibri", bold: true, size: 11, color: { argb: "FF525252" } }; cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FFF5F5F5" } }; cell.border = { bottom: { style: "thin", color: { argb: "FFE5E5E5" } } }; cell.alignment = { vertical: "middle", wrapText: true }; });
    sheet.headerFooter.oddFooter = "&LKohinoor Textile Mills Limited&RPage &P of &N";
    return sheet;
  };
  const finish = (sheet: ExcelJS.Worksheet, numeric: number[] = [], vacancyColumn?: number) => {
    sheet.autoFilter = { from: { row: 6, column: 1 }, to: { row: Math.max(6, sheet.rowCount), column: sheet.columnCount } };
    sheet.eachRow((row, index) => {
      if (index < 7) return;
      row.height = 28;
      row.eachCell((cell, column) => {
        cell.font = { name: "Calibri", size: 11, color: { argb: "FF262626" } };
        cell.alignment = { vertical: "middle", wrapText: true, horizontal: numeric.includes(column) ? "right" : "left" };
        if (index % 2) cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FFF5F5F5" } };
        if (numeric.includes(column)) cell.numFmt = "#,##0;[Red](#,##0)";
      });
    });
    if (vacancyColumn && sheet.rowCount >= 7) {
      const column = sheet.getColumn(vacancyColumn).letter;
      sheet.addConditionalFormatting({ ref: `${column}7:${column}${sheet.rowCount}`, rules: [
        { type: "cellIs", operator: "greaterThan", formulae: ["0"], priority: 1, style: { font: { color: { argb: "FF047857" } }, fill: { type: "pattern", pattern: "solid", fgColor: { argb: "FFD1FAE5" } } } },
        { type: "cellIs", operator: "lessThan", formulae: ["0"], priority: 2, style: { font: { color: { argb: "FFB91C1C" } }, fill: { type: "pattern", pattern: "solid", fgColor: { argb: "FFFEE2E2" } } } },
      ] });
    }
  };
  const cases = buildReviewCases(result, [], result.reviewCases ?? []);
  const view = strengthView(result);
  const summary = makeSheet("Summary", Array(8).fill(""), [3, 43, 17, 17, 4, 44, 17, 17], "Approved and on-roll strength by subcategory and category.");
  summary.getCell("C1").value = "Strength Summary";
  const totals = new Map<string, { sheet: string; subcategory: string; row: number }[]>();
  const categories = [...new Set([...result.categories, ...result.rows.map((row) => row.category)])];
  categories.forEach((category, categoryIndex) => {
    progress(`Building ${category} worksheet`, 15 + Math.round(categoryIndex / categories.length * 60));
    const categoryRows = result.rows.filter((row) => row.category === category);
    const sheet = makeSheet(category, ["Sr#", "Designation", "Grade", "Cadre Staff/Worker", "Approved", "Onroll", "Vacancies", "Cases"], [7, 38, 12, 18, 14, 14, 14, 60], "Vacancies = approved minus on-roll plus saved cases. Cases explain excess headcount; pending cases still need review.");
    sheet.getCell("C1").value = "KOHINOOR TEXTILE MILLS LIMITED";
    sheet.getCell("C2").value = "Gujar Khan Division";
    sheet.getCell("A3").value = `STRENGTH CHART — ${category} | ${result.generatedAt.slice(0, 10)}`;
    const sectionTotals: { sheet: string; subcategory: string; row: number }[] = [];
    const headings: number[] = [];
    const dataRows: number[] = [];
    const caseText = new Map<number, string>();
    for (const subcategory of new Set(categoryRows.map((row) => row.subcategory))) {
      const heading = sheet.addRow([subcategory]);
      sheet.mergeCells(heading.number, 1, heading.number, 8);
      headings.push(heading.number);
      const section = categoryRows.filter((row) => row.subcategory === subcategory);
      const first = sheet.rowCount + 1;
      section.forEach((row, index) => {
        const rowIndex = result.rows.indexOf(row);
        const related = cases.filter((entry) => entry.rowIndex === rowIndex);
        const text = related.map((entry) => `Case ${entry.caseNumber}: ${entry.resolution ? resolutionLabel(entry.resolution) : "Pending review"}`).join("\n");
        const number = sheet.rowCount + 1;
        const saved = confirmedAdjustments({ ...result, reviewCases: cases }, rowIndex);
        sheet.addRow([index + 1, row.designation, row.grade, row.cadre, row.approvedStrength, row.onRoll, { formula: `E${number}-F${number}+${saved}`, result: vacancyFor({ ...result, reviewCases: cases }, rowIndex) }, text]);
        dataRows.push(number);
        if (text) caseText.set(number, text);
      });
      const last = sheet.rowCount;
      const total = sheet.addRow(["", "Total", "", "", ...["E", "F", "G"].map((column, index) => ({ formula: `SUM(${column}${first}:${column}${last})`, result: section.reduce((sum, row) => sum + (index === 0 ? row.approvedStrength : index === 1 ? row.onRoll : vacancyFor({ ...result, reviewCases: cases }, result.rows.indexOf(row))), 0) }))]);
      sectionTotals.push({ sheet: sheet.name, subcategory, row: total.number });
    }
    finish(sheet, [1, 5, 6, 7], 7);
    sheet.autoFilter = undefined;
    headings.forEach((number) => {
      const cell = sheet.getCell(number, 1);
      cell.font = { name: "Calibri", size: 12, bold: true, color: { argb: "FFFFFFFF" } };
      cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FF404040" } };
      sheet.getRow(number).height = 30;
    });
    sectionTotals.forEach(({ row }) => sheet.getRow(row).eachCell((cell) => { cell.font = { name: "Calibri", size: 11, bold: true }; cell.border = { top: { style: "thin", color: { argb: "FF737373" } } }; }));
    caseText.forEach((text, number) => { sheet.getRow(number).height = Math.max(28, text.split("\n").reduce((lines, line) => lines + Math.ceil(line.length / 55), 0) * 16 + 10); });
    if (!dataRows.length) sheet.addRow(["", "No designations configured."]);
    totals.set(category, sectionTotals);
  });
  summary.getRow(6).eachCell({ includeEmpty: true }, (cell) => { cell.value = null; cell.style = {}; });
  const summaryTable = (title: string, label: string, rows: typeof view.admin, startColumn: number, subcategories: boolean) => {
    summary.mergeCells(6, startColumn, 6, startColumn + 2);
    const titleCell = summary.getCell(6, startColumn);
    titleCell.value = title;
    titleCell.font = { name: "Calibri", size: 14, bold: true, color: { argb: "FF171717" } };
    titleCell.alignment = { vertical: "middle" };
    [label, "Approved", "On-roll"].forEach((header, index) => {
      const cell = summary.getCell(7, startColumn + index);
      cell.value = header;
      cell.font = { name: "Calibri", size: 11, bold: true, color: { argb: "FF525252" } };
      cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FFF5F5F5" } };
      cell.alignment = { vertical: "middle", horizontal: index ? "right" : "left" };
      cell.border = { bottom: { style: "thin", color: { argb: "FFE5E5E5" } } };
    });
    rows.forEach((row, index) => {
      const number = index + 8;
      const source = (totals.get(row.category) ?? []).filter((entry) => !subcategories || entry.subcategory === row.name);
      const formula = (column: string) => source.length ? source.map((entry) => `'${entry.sheet.replace(/'/g, "''")}'!${column}${entry.row}`).join("+") : "0";
      summary.getCell(number, startColumn).value = row.name;
      summary.getCell(number, startColumn + 1).value = { formula: formula("E"), result: row.approved };
      summary.getCell(number, startColumn + 2).value = { formula: formula("F"), result: row.onRoll };
    });
    const totalRow = rows.length + 8;
    const approvedColumn = summary.getColumn(startColumn + 1).letter;
    const onRollColumn = summary.getColumn(startColumn + 2).letter;
    summary.getCell(totalRow, startColumn).value = "Total";
    summary.getCell(totalRow, startColumn + 1).value = { formula: rows.length ? `SUM(${approvedColumn}8:${approvedColumn}${totalRow - 1})` : "0", result: rows.reduce((sum, row) => sum + row.approved, 0) };
    summary.getCell(totalRow, startColumn + 2).value = { formula: rows.length ? `SUM(${onRollColumn}8:${onRollColumn}${totalRow - 1})` : "0", result: rows.reduce((sum, row) => sum + row.onRoll, 0) };
    for (let number = 8; number <= totalRow; number++) {
      summary.getRow(number).height = 30;
      for (let offset = 0; offset < 3; offset++) {
        const cell = summary.getCell(number, startColumn + offset);
        cell.font = { name: "Calibri", size: 11, bold: number === totalRow, color: { argb: "FF262626" } };
        cell.alignment = { vertical: "middle", horizontal: offset ? "right" : "left", wrapText: true };
        cell.border = { bottom: { style: "hair", color: { argb: "FFEEEEEE" } } };
        if (offset) cell.numFmt = "#,##0";
        if (number === totalRow) cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FFF5F5F5" } };
      }
    }
  };
  summary.getRow(6).height = 32;
  summary.getRow(7).height = 30;
  summaryTable("Admin & HR", "Subcategory", view.admin, 2, true);
  summaryTable("Spinning", "Category", view.spinning, 6, false);
  summary.pageSetup.printTitlesRow = "6:7";
  summary.autoFilter = undefined;
  progress("Writing valid Excel workbook", 85);
  const bytes = await workbook.xlsx.writeBuffer();
  return new Uint8Array(bytes).buffer;
}
