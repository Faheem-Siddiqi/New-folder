"use server";

import { readFile, writeFile, rename, unlink } from "node:fs/promises";
import { randomUUID } from "node:crypto";
import { revalidatePath } from "next/cache";
import { normalizeData, validateStrengthData, type StrengthData } from "@/lib/strength-data";
import { STRENGTH_FILE } from "@/lib/strength-file";

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
