import { stat } from "node:fs/promises";
import { LATEST_REPORT_FILE, readLatestReport } from "@/lib/latest-report";

export const dynamic = "force-dynamic";
export async function GET(request: Request) {
  const headers = { "Cache-Control": "no-store" };
  try {
    let tag = '"missing"';
    try {
      const file = await stat(LATEST_REPORT_FILE);
      tag = `"${file.mtimeMs}-${file.ctimeMs}-${file.size}"`;
    } catch (error) { if ((error as NodeJS.ErrnoException).code !== "ENOENT") throw error; }
    if (request.headers.get("If-None-Match") === tag) return new Response(null, { status: 304, headers: { ...headers, ETag: tag } });
    const result = await readLatestReport();
    return Response.json({ ok: true, result }, { headers: { ...headers, ETag: tag } });
  } catch (error) {
    return Response.json({ ok: false, error: error instanceof Error ? error.message : "The saved report could not be read. Please retry." }, { status: 500, headers });
  }
}
