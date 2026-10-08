import type { TurnoverResult } from "./turnover-contracts";

// A 304 confirms that the JSON has not changed. This never uses browser storage.
let tag: string | null = null;
let lastResult: TurnoverResult | null = null;
let hasResult = false;
let pending: Promise<TurnoverResult | null> | null = null;
export function getLatestReport(): Promise<TurnoverResult | null> {
  if (pending) return pending;
  pending = readLatestReport().finally(() => { pending = null; });
  return pending;
}
async function readLatestReport(): Promise<TurnoverResult | null> {
  const controller = new AbortController();
  const timeout = setTimeout(() => controller.abort(), 15_000);
  try {
    const response = await fetch("/api/latest-report", { cache: "no-store", signal: controller.signal, headers: tag && hasResult ? { "If-None-Match": tag } : {} });
    if (response.status === 304 && hasResult) return lastResult;
    const body = await response.json();
    if (!response.ok || !body.ok) throw new Error(body.error || "The latest saved report could not be loaded. Please retry.");
    tag = response.headers.get("ETag");
    lastResult = body.result;
    hasResult = true;
    return body.result;
  } catch (error) {
    if (error instanceof Error && error.name === "AbortError") throw new Error("The saved report check timed out. You can still upload a new workbook or retry the check.");
    throw error;
  } finally { clearTimeout(timeout); }
}
