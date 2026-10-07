"use server";

import { readFile, writeFile } from "node:fs/promises";
import { revalidatePath } from "next/cache";
import { normalizeData, type StrengthData } from "@/lib/strength-data";
import { STRENGTH_FILE } from "@/lib/strength-file";
import { readAdjustmentHistory, appendAdjustment } from "@/lib/adjustment-history";
import { validResolution, resolutionAllowed, type Resolution } from "@/lib/review-cases";
import { strengthRows } from "@/lib/strength-rows";

async function readStrengthData(): Promise<StrengthData> {
  const contents = await readFile(STRENGTH_FILE, "utf8");
  return normalizeData(JSON.parse(contents) as StrengthData);
}

async function writeStrengthData(data: StrengthData) {
  await writeFile(STRENGTH_FILE, `${JSON.stringify(data, null, 2)}\n`, "utf8");
}

export async function saveStrengthData(data: StrengthData) {
  const next = normalizeData(data);
  await writeStrengthData(next);
  revalidatePath("/config");
  revalidatePath("/");
  return next;
}

export async function getStrengthData() {
  return readStrengthData();
}

export async function getCaseHistory() {
  try { return await readAdjustmentHistory(); }
  catch (error) { return { error: error instanceof Error ? error.message : "Case history could not be loaded. Your existing report and history are preserved." }; }
}

export async function saveCaseResolution(value: Resolution) {
  try {
  if (!validResolution(value)) throw new Error("Choose a valid designation adjustment.");
  const rows = strengthRows(await readStrengthData());
  const row = rows.find((entry) => entry.category === value.category && entry.subcategory === value.subcategory && entry.designation === value.designation);
  if (!row || !resolutionAllowed(rows, row, value)) throw new Error("The designation or assignment changed in configuration. Reprocess your upload and choose another valid post in the same subcategory.");
  return { ok: true as const, resolution: await appendAdjustment(value) };
  } catch (error) {
    return { ok: false as const, error: error instanceof Error ? error.message : "The case assignment could not be saved. Please retry." };
  }
}
