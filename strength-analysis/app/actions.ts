"use server";

import { readFile, writeFile } from "node:fs/promises";
import { join } from "node:path";
import { revalidatePath } from "next/cache";
import { normalizeData, type StrengthData } from "@/lib/strength-data";

const DATA_FILE = "strength-master-seed-0.json";

async function readStrengthData(): Promise<StrengthData> {
  const file = join(process.cwd(), DATA_FILE);
  const contents = await readFile(file, "utf8");
  return normalizeData(JSON.parse(contents) as StrengthData);
}

async function writeStrengthData(data: StrengthData) {
  const file = join(process.cwd(), DATA_FILE);
  await writeFile(file, `${JSON.stringify(data, null, 2)}\n`, "utf8");
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
