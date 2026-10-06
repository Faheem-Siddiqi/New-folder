import { TURNOVER_HEADERS, type TurnoverResult } from "./turnover-contracts.ts";

export type SavedReport = { version: 1; file: Blob; fileName: string; result: TurnoverResult };
async function database(): Promise<IDBDatabase> {
  if (!globalThis.indexedDB) throw new Error("Browser storage is unavailable.");
  return new Promise((resolve, reject) => {
    const request = indexedDB.open("kohinoor-turnover", 1);
    request.onupgradeneeded = () => request.result.createObjectStore("reports");
    request.onsuccess = () => resolve(request.result);
    request.onerror = () => reject(request.error);
    request.onblocked = () => reject(new Error("Close other application tabs and try again."));
  });
}
export async function loadReport(): Promise<SavedReport | null> {
  const db = await database();
  try {
    return await new Promise((resolve, reject) => {
      const transaction = db.transaction("reports", "readonly");
      const request = transaction.objectStore("reports").get("latest");
      transaction.oncomplete = () => {
        const saved = request.result as SavedReport | undefined;
        if (!saved) return resolve(null);
        const result = saved.result;
        const valid = saved.version === 1 && saved.file instanceof Blob && typeof saved.fileName === "string" && result?.version === 1 && typeof result.templateSignature === "string" && typeof result.fileName === "string" && Array.isArray(result.categories) && result.categories.every((name) => typeof name === "string") && Number.isFinite(Date.parse(result.generatedAt)) &&
          [result.scannedRows, result.duplicates, result.skippedRows, result.matched, result.unmatched].every((value) => Number.isSafeInteger(value) && value >= 0) &&
          Array.isArray(result.rows) && result.rows.every((row) => row && [row.category, row.subcategory, row.designation, row.grade, row.cadre, row.status].every((value) => typeof value === "string") && [row.approvedStrength, row.onRoll].every((value) => Number.isSafeInteger(value) && value >= 0)) &&
          Array.isArray(result.employees) && result.employees.every((record) => record?.employee && TURNOVER_HEADERS.every((header) => typeof record.employee[header] === "string") && (record.match === null || (Number.isInteger(record.match) && record.match >= 0 && record.match < result.rows.length)) && typeof record.sheet === "string" && typeof record.issue === "string" && Number.isInteger(record.row)) &&
          Array.isArray(result.issues) && result.issues.every((issue) => issue && typeof issue.sheet === "string" && typeof issue.employeeId === "string" && typeof issue.message === "string" && Number.isInteger(issue.row));
        if (!valid) return reject(new Error("Saved report is invalid. Upload the original workbook again."));
        resolve(saved);
      };
      transaction.onerror = () => reject(transaction.error);
    });
  } finally { db.close(); }
}
export async function storeReport(report: SavedReport): Promise<void> {
  const db = await database();
  try {
    await new Promise<void>((resolve, reject) => {
      const transaction = db.transaction("reports", "readwrite");
      transaction.objectStore("reports").put(report, "latest");
      transaction.oncomplete = () => resolve();
      transaction.onerror = () => reject(transaction.error);
      transaction.onabort = () => reject(transaction.error ?? new Error("Storage was cancelled."));
    });
  } finally { db.close(); }
}
