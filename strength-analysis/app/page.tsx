import { getStrengthData } from "@/app/actions";
import { TurnoverHome } from "@/components/turnover-home";
import { readLatestReport } from "@/lib/latest-report";

export const dynamic = "force-dynamic";

export default async function HomePage() {
  const [template, saved] = await Promise.all([
    getStrengthData(),
    readLatestReport().then((result) => ({ result, error: "" })).catch((error) => ({ result: null, error: error instanceof Error ? error.message : "The saved report could not be read." })),
  ]);
  return <TurnoverHome template={template} initialReport={saved.result} initialReportError={saved.error} />;
}
