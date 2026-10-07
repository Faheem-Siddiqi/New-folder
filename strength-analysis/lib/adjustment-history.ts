import { readFile, writeFile, rename, unlink } from "node:fs/promises";
import { resolve } from "node:path";
import { randomUUID } from "node:crypto";
import { validResolution, validHistoryRecord, type Resolution } from "./review-cases.ts";

export const ADJUSTMENT_FILE = resolve(process.cwd(), "..", "strength-adjustment.json");
export async function readAdjustmentHistory(path = ADJUSTMENT_FILE): Promise<Resolution[]> {
  try {
    const parsed = JSON.parse(await readFile(path, "utf8"));
    if (parsed.version !== 1 || !Array.isArray(parsed.records) || !parsed.records.every(validHistoryRecord)) throw new Error("Invalid case history structure");
    return parsed.records;
  } catch (error) {
    if ((error as NodeJS.ErrnoException).code === "ENOENT") return [];
    throw new Error("Case history could not be read. The existing history has been preserved; restore a valid strength-adjustment.json before saving cases.");
  }
}
let queue: Promise<unknown> = Promise.resolve();
export async function appendAdjustment(resolution: Resolution, path = ADJUSTMENT_FILE) {
  if (!validResolution(resolution)) throw new Error("Invalid case assignment.");
  const job = queue.catch(() => {}).then(async () => {
    const records = await readAdjustmentHistory(path);
    const existing = records.find((entry) => entry.id === resolution.id);
    if (existing) {
      if (JSON.stringify(existing) !== JSON.stringify(resolution)) throw new Error("This confirmation ID has already been used. Please retry the assignment.");
      return existing;
    }
    const temporary = `${path}.${randomUUID()}.tmp`;
    try { await writeFile(temporary, `${JSON.stringify({ version: 1, records: [...records, resolution] }, null, 2)}\n`, { encoding: "utf8", flag: "wx" }); await rename(temporary, path); }
    catch { throw new Error("Case history could not be saved. Check that strength-adjustment.json is writable, then retry. Your previous history and strength template are preserved."); }
    finally { await unlink(temporary).catch(() => {}); }
    return resolution;
  });
  queue = job;
  return job;
}
