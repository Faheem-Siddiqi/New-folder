"use server";

import { readFile, writeFile, rename, unlink } from "node:fs/promises";
import { randomUUID } from "node:crypto";
import { revalidatePath } from "next/cache";
import { normalizeData, validateStrengthData, type StrengthData } from "@/lib/strength-data";
import { STRENGTH_FILE } from "@/lib/strength-file";
import { saveLatestCase } from "@/lib/latest-report";
import { validResolution, resolutionAllowed, type Resolution } from "@/lib/review-cases";
import { strengthRows } from "@/lib/strength-rows";

async function readStrengthData(): Promise<StrengthData> {
  const contents = await readFile(STRENGTH_FILE, "utf8");
  return normalizeData(JSON.parse(contents) as StrengthData);
}

let configurationQueue: Promise<unknown> = Promise.resolve();
async function writeStrengthData(data: StrengthData, expected?: string) {
  const job = configurationQueue.catch(() => {}).then(async () => {
    if (expected !== undefined && JSON.stringify(await readStrengthData()) !== expected) throw new Error("The configuration changed in another session. Reload Configuration before saving to avoid overwriting those changes.");
    const temporary = `${STRENGTH_FILE}.${randomUUID()}.tmp`;
    try {
      await writeFile(temporary, `${JSON.stringify(data, null, 2)}\n`, { encoding: "utf8", flag: "wx" });
      await rename(temporary, STRENGTH_FILE);
    } finally { await unlink(temporary).catch(() => {}); }
  });
  configurationQueue = job;
  await job;
}

export async function saveStrengthData(data: StrengthData, expected?: string) {
  validateStrengthData(data);
  const next = normalizeData(data);
  await writeStrengthData(next, expected);
  revalidatePath("/config");
  revalidatePath("/");
  return next;
}

export async function getStrengthData() {
  return readStrengthData();
}

export async function saveCaseResolution(value: Resolution, reportId: string, revision: number) {
  try {
  if (!validResolution(value)) throw new Error("Choose a valid designation adjustment.");
  const rows = strengthRows(await readStrengthData());
  const row = rows.find((entry) => entry.category === value.category && entry.subcategory === value.subcategory && entry.designation === value.designation);
  if (!row || !resolutionAllowed(rows, row, value)) throw new Error("The designation or assignment changed in configuration. Reprocess your upload and choose another valid post in the same subcategory.");
  return { ok: true as const, result: await saveLatestCase(value, reportId, revision) };
  } catch (error) {
    return { ok: false as const, error: error instanceof Error ? error.message : "The case assignment could not be saved. Please retry." };
  }
}
