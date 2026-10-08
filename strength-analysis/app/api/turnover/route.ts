import { getStrengthData } from "@/app/actions";

export const dynamic = "force-dynamic";
const headers = { "Cache-Control": "no-store" };
export async function GET() {
  try { return Response.json({ ok: true, result: await getStrengthData() }, { headers }); }
  catch (error) { return Response.json({ ok: false, error: error instanceof Error ? error.message : "The strength configuration could not be read." }, { status: 500, headers }); }
}