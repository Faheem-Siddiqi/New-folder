import { getStrengthData } from "@/app/actions";
import { TurnoverHome } from "@/components/turnover-home";
export const dynamic = "force-dynamic";
export default async function HomePage() {
  return <TurnoverHome template={await getStrengthData()} />;
}
