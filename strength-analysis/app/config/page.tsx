import { ConfigEditor } from "@/components/config-editor";
import { getStrengthData } from "@/app/actions";

export const dynamic = "force-dynamic";

export default async function ConfigPage() {
  const initialData = await getStrengthData();

  return <ConfigEditor initialData={initialData} />;
}
