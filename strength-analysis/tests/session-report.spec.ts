import { test, expect } from "@playwright/test";
import { readFile } from "node:fs/promises";
import * as XLSX from "xlsx";
import { TURNOVER_HEADERS } from "../lib/turnover-contracts";
import { STRENGTH_FILE } from "../lib/strength-file";
import { SESSION_REPORT_KEY } from "../lib/session-report";

async function upload(page: import("@playwright/test").Page) {
  const template = JSON.parse(await readFile(STRENGTH_FILE, "utf8"));
  const category = template.strengthStructure[0];
  const section = category.subcategories[0];
  const post = section.designations[0];
  const workbook = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(workbook, XLSX.utils.aoa_to_sheet([[...TURNOVER_HEADERS], ...Array.from({ length: post.approvedStrength + 1 }, (_, i) => [section.subcategory, "KGM", String(i), `Employee ${i}`, post.designation, "M", "Sunday", post.grade, "A", post.cadre, category.category, "2024-01-01"])]), "Employees");
  await page.getByRole("button", { name: /Upload (new )?report/, exact: true }).click();
  await page.getByLabel("Turnover workbook", { exact: true }).setInputFiles({ name: "session.xlsx", mimeType: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", buffer: XLSX.write(workbook, { type: "buffer", bookType: "xlsx" }) });
  await page.getByRole("button", { name: "Start scanning" }).click();
}

test("session retains cases on reload, downloads, starts new uploads fresh, and disappears when the tab closes", async ({ page, context }) => {
  const before = await readFile(STRENGTH_FILE, "utf8");
  const requests: string[] = [];
  page.on("request", (request) => { if (request.url().includes("api/latest-report") || request.method() === "POST") requests.push(request.url()); });
  await page.goto("/");
  await upload(page);
  await page.getByRole("dialog", { name: "Report ready", exact: true }).getByRole("button", { name: "Got it" }).click();
  await page.getByRole("button", { name: "Review cases", exact: true }).click();
  let dialog = page.getByRole("dialog", { name: "Review Cases", exact: true });
  await dialog.getByRole("combobox").selectOption("__social_security__");
  await dialog.getByRole("button", { name: "Save case", exact: true }).click();
  await page.getByRole("dialog", { name: "Case saved" }).getByRole("button", { name: "Got it" }).click();
  await expect(dialog.getByRole("button", { name: "Edit", exact: true })).toBeVisible();
  await expect(dialog.getByText("Vacancy: 0", { exact: true })).toBeVisible();
  await page.reload();
  await expect(page.getByText("Restored from this tab's session.", { exact: true })).toBeVisible();
  await page.getByRole("button", { name: "Review cases", exact: true }).click();
  dialog = page.getByRole("dialog", { name: "Review Cases", exact: true });
  await expect(dialog.getByText("Vacancy: 0", { exact: true })).toBeVisible();
  await dialog.getByRole("button", { name: "Close review cases" }).click();
  const download = page.waitForEvent("download");
  await page.getByRole("button", { name: "Download Excel", exact: true }).click();
  expect((await download).suggestedFilename()).toMatch(/\.xlsx$/);
  await page.getByRole("dialog", { name: "Excel download started" }).getByRole("button", { name: "Got it" }).click();
  await upload(page);
  await page.getByRole("dialog", { name: "Report ready", exact: true }).getByRole("button", { name: "Got it" }).click();
  const saved = await page.evaluate((key) => JSON.parse(sessionStorage.getItem(key)!), SESSION_REPORT_KEY);
  expect(saved.reviewCases).toHaveLength(1);
  expect(saved.reviewCases[0].resolution).toBeUndefined();
  expect(saved.reviewCases[0].suggestion).toBeUndefined();
  expect(requests).toEqual([]);
  expect(await readFile(STRENGTH_FILE, "utf8")).toBe(before);
  await page.close();
  const reopened = await context.newPage();
  await reopened.goto("/");
  await expect(reopened.getByRole("heading", { name: "Start with your turnover report" })).toBeVisible();
  await expect(reopened.getByRole("button", { name: "Download Excel" })).toBeDisabled();
});

test("session quota failure preserves report and case input for retry", async ({ page }) => {
  await page.goto("/");
  await upload(page);
  await page.getByRole("dialog", { name: "Report ready", exact: true }).getByRole("button", { name: "Got it" }).click();
  const before = await page.evaluate((key) => sessionStorage.getItem(key), SESSION_REPORT_KEY);
  await page.evaluate(() => { Storage.prototype.setItem = () => { throw new DOMException("Full", "QuotaExceededError"); }; });
  await page.getByRole("button", { name: "Review cases", exact: true }).click();
  const dialog = page.getByRole("dialog", { name: "Review Cases", exact: true });
  await dialog.getByRole("combobox").selectOption("__other__");
  await dialog.getByRole("textbox", { name: /Other reason/ }).fill("Temporary duty");
  await dialog.getByRole("button", { name: "Save case", exact: true }).click();
  const failure = page.getByRole("dialog", { name: "Case could not be saved" });
  await expect(failure.getByText(/Storage may be full or disabled/)).toBeVisible();
  await failure.getByRole("button", { name: "Got it" }).click();
  await expect(dialog.getByRole("textbox", { name: /Other reason/ })).toHaveValue("Temporary duty");
  await expect(dialog.getByText("Vacancy: -1", { exact: true })).toBeVisible();
  expect(await page.evaluate((key) => sessionStorage.getItem(key), SESSION_REPORT_KEY)).toBe(before);
});

test("corrupt session does not block a new upload", async ({ page }) => {
  await page.addInitScript((key) => sessionStorage.setItem(key, "{broken"), SESSION_REPORT_KEY);
  await page.goto("/");
  await expect(page.getByRole("heading", { name: "Start with your turnover report" })).toBeVisible();
  await expect(page.getByRole("button", { name: "Upload report", exact: true })).toBeEnabled();
  await upload(page);
  await expect(page.getByRole("dialog", { name: "Report ready", exact: true })).toBeVisible();
});
