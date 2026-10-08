import type { StrengthData } from "./strength-data";

async function request(path: string, options: RequestInit, timeoutMs: number) {
  const controller = new AbortController();
  const timeout = setTimeout(() => controller.abort(), timeoutMs);
  try {
    const response = await fetch(path, { ...options, cache: "no-store", signal: controller.signal });
    const body = await response.json().catch(() => ({ ok: false, error: "The app returned an unreadable response. Refresh and retry." }));
    if (!response.ok || !body.ok) throw new Error(body.error || "The report operation failed. Please retry.");
    return body.result;
  } catch (error) {
    if (error instanceof Error && error.name === "AbortError") throw new Error("The request timed out. Please retry.");
    throw error;
  } finally { clearTimeout(timeout); }
}
export async function getStrengthData(): Promise<StrengthData> {
  return request("/api/turnover", {}, 15_000);
}