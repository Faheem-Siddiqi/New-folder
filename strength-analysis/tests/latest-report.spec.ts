import { test, expect } from "@playwright/test";
import { readFile, writeFile } from "node:fs/promises";
import * as XLSX from "xlsx";
import { TURNOVER_HEADERS } from "../lib/turnover-contracts";
import { STRENGTH_FILE } from "../lib/strength-file";

test("latest JSON saves cases, restores across browsers, and requires approval on a matching new upload", async ({ page, browser }) => {
  const path = "../strength-last-result.json";
  const before = await readFile(path, "utf8");
  const templateBefore = await readFile(STRENGTH_FILE, "utf8");
  try {
    await writeFile(path, JSON.stringify({ version: 1, result: null }));
    const template = JSON.parse(templateBefore);
    const category = template.strengthStructure[0];
    const section = category.subcategories[0];
    const designation = section.designations[0];
    const target = section.designations.find((entry: { designation: string }) => entry.designation !== designation.designation);
    const workbook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(workbook, XLSX.utils.aoa_to_sheet([[...TURNOVER_HEADERS], ...Array.from({ length: designation.approvedStrength + 3 }, (_, index) => [section.subcategory, "KGM", String(index), `Employee ${index}`, designation.designation, "M", "Sunday", designation.grade, "A", designation.cadre, category.category, "2024-01-01"])]), "Employees");
    const upload = { name: "latest-report.xlsx", mimeType: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", buffer: XLSX.write(workbook, { type: "buffer", bookType: "xlsx" }) };
    const scan = async () => {
      await page.getByRole("button", { name: /Upload (new )?report/, exact: true }).click();
      await page.getByLabel("Turnover workbook", { exact: true }).setInputFiles(upload);
      await page.getByRole("button", { name: "Start scanning" }).click();
      await expect(page.getByRole("dialog", { name: "Report ready", exact: true })).toBeVisible();
      await page.getByRole("button", { name: "Got it" }).click();
    };
    const readResult = async () => JSON.parse(await readFile(path, "utf8")).result;
    await page.goto("/");
    await scan();
    expect((await readResult()).reviewCases).toHaveLength(3);
    await page.getByRole("button", { name: "Review cases", exact: true }).click();
    const dialog = page.getByRole("dialog", { name: "Review Cases", exact: true });
    const cases = dialog.getByTestId("designation-case");
    let holdSave = true;
    let failSave = false;
    let releaseSave: (() => void) | undefined;
    await page.route("**/*", async (route) => {
      const request = route.request();
      if (request.method() === "POST" && request.headers()["next-action"] && (request.postData() ?? "").includes('"caseType"')) {
        if (failSave) return route.abort();
        if (holdSave) { holdSave = false; await new Promise<void>((resolve) => { releaseSave = resolve; }); }
      }
      await route.continue();
    });
    await cases.nth(0).getByRole("combobox").selectOption(`post:${target.designation}`);
    await cases.nth(0).getByRole("button", { name: "Save case", exact: true }).click();
    await expect(cases.nth(0).getByRole("button", { name: "Saving...", exact: true })).toBeDisabled();
    await expect(dialog.getByRole("button", { name: "Saving...", exact: true })).toHaveCount(1);
    await expect(dialog.getByRole("button", { name: "Close review cases" })).toBeDisabled();
    await page.keyboard.press("Escape");
    await expect(dialog).toBeVisible();
    await expect.poll(() => !!releaseSave).toBe(true);
    releaseSave!();
    await expect(page.getByRole("dialog", { name: "Case saved", exact: true })).toBeVisible();
    await page.getByRole("button", { name: "Got it", exact: true }).click();
    await expect(cases.nth(0).getByRole("button", { name: "Edit", exact: true })).toBeVisible();
    await expect(dialog.getByText("Vacancy: -2", { exact: true })).toBeVisible();
    await cases.nth(1).getByRole("combobox").selectOption("__social_security__");
    await cases.nth(1).getByRole("button", { name: "Save case", exact: true }).click();
    await page.getByRole("dialog", { name: "Case saved", exact: true }).getByRole("button", { name: "Got it" }).click();
    await expect(dialog.getByText("Vacancy: -1", { exact: true })).toBeVisible();
    await cases.nth(2).getByRole("combobox").selectOption("__other__");
    await cases.nth(2).getByRole("textbox").fill("Temporary duty");
    await cases.nth(2).getByRole("button", { name: "Save case", exact: true }).click();
    await page.getByRole("dialog", { name: "Case saved", exact: true }).getByRole("button", { name: "Got it" }).click();
    await expect(dialog.getByText("Vacancy: 0", { exact: true })).toBeVisible();
    await cases.nth(2).getByRole("button", { name: "Edit", exact: true }).click();
    await cases.nth(2).getByRole("textbox").fill("Updated duty");
    failSave = true;
    await cases.nth(2).getByRole("button", { name: "Save case", exact: true }).click();
    await page.getByRole("dialog", { name: "Case could not be saved" }).getByRole("button", { name: "Got it" }).click();
    await expect(cases.nth(2).getByRole("textbox")).toHaveValue("Updated duty");
    await expect(cases.nth(2).getByRole("button", { name: "Save case", exact: true })).toBeEnabled();
    failSave = false;
    await cases.nth(2).getByRole("button", { name: "Save case", exact: true }).click();
    await page.getByRole("dialog", { name: "Case updated" }).getByRole("button", { name: "Got it" }).click();
    await expect(cases.nth(2).getByRole("button", { name: "Edit", exact: true })).toBeVisible();
    expect((await readResult()).reviewCases.map((entry: { resolution: { caseType: string } }) => entry.resolution.caseType)).toEqual(["Against Post", "Social Security Leave", "Other"]);
    // Simulate a newer operation on the server while this dialog remains open.
    const external = JSON.parse(await readFile(path, "utf8"));
    external.result.reviewCases[2].resolution.otherReason = "Updated on server";
    external.result.revision += 1;
    await writeFile(path, JSON.stringify(external));
    await dialog.getByRole("button", { name: "Close review cases" }).click();
    await expect(page.getByText("Loaded from the latest saved report.", { exact: true })).toBeVisible();
    await page.getByRole("button", { name: "Review cases", exact: true }).click();
    await expect(cases.nth(2).getByText(/Updated on server/)).toBeVisible();
    await dialog.getByRole("button", { name: "Close review cases" }).click();
    const another = await browser.newContext();
    try {
      const second = await another.newPage();
      await second.goto("http://127.0.0.1:3100/");
      await expect(second.getByText("Restored from the latest saved report.", { exact: true })).toBeVisible();
      await second.getByRole("button", { name: "Review cases", exact: true }).click();
      await expect(second.getByText("Vacancy: 0", { exact: true })).toBeVisible();
    } finally { await another.close(); }
    await scan();
    await page.getByRole("button", { name: "Review cases", exact: true }).click();
    await expect(dialog.getByText("Previous Match Found", { exact: true })).toHaveCount(3);
    await expect(dialog.getByText("Vacancy: -3", { exact: true })).toBeVisible();
    expect((await readResult()).reviewCases.every((entry: { resolution?: unknown }) => !entry.resolution)).toBe(true);
    await cases.nth(0).getByRole("button", { name: "OK, approve", exact: true }).click();
    await page.getByRole("dialog", { name: "Case saved", exact: true }).getByRole("button", { name: "Got it" }).click();
    await expect(dialog.getByText("Vacancy: -2", { exact: true })).toBeVisible();
    expect((await readResult()).reviewCases[0].resolution.caseType).toBe("Against Post");
    await page.reload();
    await page.getByRole("button", { name: "Review cases", exact: true }).click();
    await expect(dialog.getByText("Vacancy: -2", { exact: true })).toBeVisible();
    await expect(dialog.getByText("Previous Match Found", { exact: true })).toHaveCount(2);
    expect(await readFile(STRENGTH_FILE, "utf8")).toBe(templateBefore);
  } finally { await writeFile(path, before); }
});

test("configuration validates duplicate names and preserves staged changes after a failed save", async ({ page }) => {
  const before = await readFile(STRENGTH_FILE, "utf8");
  const category = JSON.parse(before).strengthStructure[0].category;
  await page.goto("/config");
  await page.getByRole("button", { name: "Add category", exact: true }).click();
  const modal = page.getByRole("dialog", { name: "Add category", exact: true });
  await modal.getByRole("textbox", { name: "Category name" }).fill(` ${category.toLowerCase()} `);
  await modal.getByRole("button", { name: "Add category", exact: true }).click();
  await expect(modal.getByRole("alert")).toHaveText("This category already exists.");
  await modal.getByRole("textbox", { name: "Category name" }).fill("TEST SAVE FAILURE");
  await modal.getByRole("button", { name: "Add category", exact: true }).click();
  await page.getByRole("dialog", { name: "Category added" }).getByRole("button", { name: "Got it" }).click();
  let release: (() => void) | undefined;
  await page.route("**/config", async (route) => {
    if (route.request().method() === "POST") {
      await new Promise<void>((resolve) => { release = resolve; });
      await route.abort();
    } else await route.continue();
  });
  await page.getByRole("button", { name: "Save changes", exact: true }).click();
  await expect(page.getByRole("button", { name: "Saving...", exact: true })).toBeDisabled();
  await expect(page.getByRole("button", { name: "Add category", exact: true })).toBeDisabled();
  await expect.poll(() => !!release).toBe(true);
  release!();
  await page.getByRole("dialog", { name: "Configuration could not be saved" }).getByRole("button", { name: "Got it" }).click();
  await expect(page.getByRole("button", { name: "Save changes", exact: true })).toBeEnabled();
  await expect(page.getByText("Unsaved changes", { exact: true })).toBeVisible();
  expect(await readFile(STRENGTH_FILE, "utf8")).toBe(before);
});

test("empty history displays immediately, background checks do not block upload, and save errors remain readable", async ({ page, request }) => {
  const path = "../strength-last-result.json";
  const before = await readFile(path, "utf8");
  try {
    await writeFile(path, JSON.stringify({ version: 1, result: null }));
    await page.addInitScript(() => {
      Object.defineProperty(window, "indexedDB", { get() { throw new Error("Browser storage must not block empty history"); } });
    });
    let release: (() => void) | undefined;
    let hold = true;
    await page.route("**/api/latest-report", async (route) => {
      if (hold) { hold = false; await new Promise<void>((resolve) => { release = resolve; }); }
      await route.continue();
    });
    await page.goto("/");
    await expect(page.getByRole("heading", { name: "Start with your turnover report" })).toBeVisible();
    await expect(page.getByText("Checking for your saved report…")).toHaveCount(0);
    await expect(page.getByRole("button", { name: "Upload report", exact: true })).toBeEnabled();
    await page.getByRole("button", { name: "Upload report", exact: true }).click();
    await expect(page.getByRole("dialog", { name: "Upload turnover report" })).toBeVisible();
    await expect.poll(() => !!release).toBe(true);
    release!();
    const empty = await request.get("/api/latest-report");
    expect((await empty.json()).result).toBeNull();
    const unchanged = await request.get("/api/latest-report", { headers: { "If-None-Match": empty.headers().etag } });
    expect(unchanged.status()).toBe(304);

    const tree = JSON.parse(await readFile(STRENGTH_FILE, "utf8"));
    const category = tree.strengthStructure[0];
    const section = category.subcategories[0];
    const designation = section.designations[0];
    const workbook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(workbook, XLSX.utils.aoa_to_sheet([[...TURNOVER_HEADERS], [section.subcategory, "KGM", "001", "Employee", designation.designation, "M", "Sunday", designation.grade, "A", designation.cadre, category.category, "2024-01-01"]]), "Employees");
    const upload = { name: "readable-failure.xlsx", mimeType: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", buffer: XLSX.write(workbook, { type: "buffer", bookType: "xlsx" }) };
    await writeFile(path, "{broken");
    await page.getByLabel("Turnover workbook", { exact: true }).setInputFiles(upload);
    await page.getByRole("button", { name: "Start scanning" }).click();
    const failure = page.getByRole("dialog", { name: "Report could not be processed" });
    await expect(failure.getByText(/strength-last-result.json could not be read/)).toBeVisible();
    await expect(failure.getByText(/Minified React/)).toHaveCount(0);
    expect(await readFile(path, "utf8")).toBe("{broken");
    const corrupt = await request.get("/api/latest-report", { headers: { "If-None-Match": empty.headers().etag } });
    expect(corrupt.status()).toBe(500);
    expect((await corrupt.json()).error).toMatch(/could not be read/);
    await failure.getByRole("button", { name: "Got it" }).click();
    await writeFile(path, JSON.stringify({ version: 1, result: null }));
    await page.getByRole("button", { name: "Upload report", exact: true }).click();
    await page.getByLabel("Turnover workbook", { exact: true }).setInputFiles(upload);
    await page.getByRole("button", { name: "Start scanning" }).click();
    await expect(page.getByRole("dialog", { name: "Report ready", exact: true })).toBeVisible();
    expect(JSON.parse(await readFile(path, "utf8")).result.fileName).toBe(upload.name);
  } finally { await writeFile(path, before); }
});

test("the existing full report can be saved through the JSON endpoint", async ({ request }) => {
  const path = "../strength-last-result.json";
  const before = await readFile(path, "utf8");
  const original = JSON.parse(before).result;
  test.skip(!original, "No existing report to exercise the full saved payload.");
  try {
    const scanned = { ...original };
    delete scanned.reportId;
    delete scanned.revision;
    delete scanned.reviewCases;
    const response = await request.post("/api/turnover", { data: { result: scanned } });
    const saved = await response.json();
    expect(saved.ok, saved.error).toBe(true);
    expect(saved.result.employees).toEqual(scanned.employees);
    expect(saved.result.rows).toEqual(scanned.rows);
    const invalid = await request.post("/api/turnover", { data: { result: { ...scanned, rows: [{ ...scanned.rows[0], onRoll: -1 }] } } });
    const failure = await invalid.json();
    expect(failure.ok).toBe(false);
    expect(failure.error).toMatch(/Invalid latest report/);
    expect(JSON.parse(await readFile(path, "utf8")).result.reportId).toBe(saved.result.reportId);
  } finally { await writeFile(path, before); }
});
