import test from "node:test";
import assert from "node:assert/strict";
import { IDBFactory } from "fake-indexeddb";
import { loadReport, storeReport } from "./report-storage.ts";
import { TURNOVER_HEADERS } from "./turnover-contracts.ts";

const example = () => ({ version: 1, file: new Blob(["original workbook"]), fileName: "example.xlsx", result: { version: 1, fileName: "example.xlsx", generatedAt: new Date().toISOString(), templateSignature: "{}", categories: ["RING"], rows: [{ category: "RING", subcategory: "Spinning", designation: "OPERATOR", grade: "E-03", cadre: "Worker", approvedStrength: 5, onRoll: 1, status: "Active" }], employees: [{ employee: Object.fromEntries(TURNOVER_HEADERS.map((header) => [header, header === "Emp #" ? "001" : "sample"])), sheet: "Data", row: 2, match: 0, issue: "" }], issues: [], scannedRows: 1, matched: 1, unmatched: 0, duplicates: 0, skippedRows: 0 } });

test("saved report and original workbook survive reopening the database", async () => {
  globalThis.indexedDB = new IDBFactory();
  assert.equal(await loadReport(), null);
  const source = example();
  await storeReport(source);
  const restored = await loadReport();
  assert.deepEqual(restored.result, source.result);
  assert.equal(await restored.file.text(), "original workbook");
  assert.equal((await loadReport()).result.rows[0].onRoll, 1);
});

test("invalid restored report is rejected gracefully rather than rendering corrupt data", async () => {
  globalThis.indexedDB = new IDBFactory();
  const corrupted = example();
  corrupted.result.rows[0].onRoll = -1;
  await storeReport(corrupted);
  await assert.rejects(loadReport(), /Saved report is invalid/);
});

test("denied browser storage produces a recoverable error", async () => {
  globalThis.indexedDB = undefined;
  await assert.rejects(loadReport(), /storage is unavailable/);
  await assert.rejects(storeReport(example()), /storage is unavailable/);
});
