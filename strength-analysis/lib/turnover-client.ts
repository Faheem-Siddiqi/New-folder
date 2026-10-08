import type { StrengthData } from "./strength-data";
import type { TurnoverResult } from "./turnover-contracts";

async function request(path: string, options: RequestInit, timeoutMs: number) {
  const controller = new AbortController();
  const timeout = setTimeout(() => controller.abort(), timeoutMs);
  try {
    const response = await fetch(path, { ...options, cache: "no-store", signal: controller.signal });
    const body = await response.json().catch(() => ({ ok: false, error: "The app returned an unreadable response. Refresh and retry." }));
    if (!response.ok || !body.ok) throw new Error(body.error || "The report operation failed. Please retry.");
    return body.result;
  } catch (error) {
    if (error instanceof Error && error.name === "AbortError") throw new Error("The request timed out. Refresh to check the latest saved report, then retry if needed.");
    throw error;
  } finally { clearTimeout(timeout); }
}
export async function getStrengthData(): Promise<StrengthData> {
  return request("/api/turnover", {}, 15_000);
}
export async function saveTurnoverResult(result: TurnoverResult, expectedReportId?: string): Promise<{ ok: true; result: TurnoverResult } | { ok: false; error: string }> {
  try {
    const saved = await request("/api/turnover", { method: "POST", headers: { "Content-Type": "application/json" }, body: JSON.stringify({ result, expectedReportId }) }, 60_000);
    return { ok: true, result: saved };
  } catch (error) { return { ok: false, error: error instanceof Error ? error.message : "The report could not be saved. Please retry." }; }
}
export async function refreshTurnoverResult(expectedReportId: string): Promise<TurnoverResult> {
  return request("/api/turnover", { method: "POST", headers: { "Content-Type": "application/json" }, body: JSON.stringify({ operation: "refresh", expectedReportId }) }, 60_000);
}
