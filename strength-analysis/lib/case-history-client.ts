import { getCaseHistory } from "@/app/actions";

export async function loadCaseHistory() {
  const response = await getCaseHistory();
  if (!Array.isArray(response)) throw new Error(response.error);
  return response;
}
