import { getStrengthData } from "@/app/actions";
import { saveProcessedReport, refreshLatestReport } from "@/lib/latest-report";

export const dynamic = "force-dynamic";
const headers = { "Cache-Control": "no-store" };
export async function GET() {
  try { return Response.json({ ok: true, result: await getStrengthData() }, { headers }); }
  catch (error) { return Response.json({ ok: false, error: error instanceof Error ? error.message : "The strength configuration could not be read." }, { status: 500, headers }); }
}
export async function POST(request: Request) {
  try {
    const { result, expectedReportId, operation } = await request.json();
    if (expectedReportId !== undefined && typeof expectedReportId !== "string") throw new Error("Invalid report identity. Refresh and retry.");
    if (operation === "refresh") {
      if (!expectedReportId) throw new Error("Choose an existing saved report to refresh.");
      return Response.json({ ok: true, result: await refreshLatestReport(await getStrengthData(), expectedReportId) }, { headers });
    }
    if (operation !== undefined) throw new Error("Unknown report operation.");
    return Response.json({ ok: true, result: await saveProcessedReport(result, expectedReportId) }, { headers });
  } catch (error) {
    return Response.json({ ok: false, error: error instanceof Error ? error.message : "The report could not be saved. Please retry." }, { status: 400, headers });
  }
}
