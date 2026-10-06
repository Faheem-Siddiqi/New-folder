import { createReportWorkbook } from "./report-workbook";
import type { TurnoverResult } from "./turnover";

self.onmessage = async (event: MessageEvent<{ result: TurnoverResult; logo?: string }>) => {
  try {
    const buffer = await createReportWorkbook(event.data.result, event.data.logo, (step, percent) => self.postMessage({ type: "progress", step, percent }));
    self.postMessage({ type: "complete", buffer }, { transfer: [buffer] });
  } catch {
    self.postMessage({ type: "error", message: "The Excel workbook could not be generated. Your scanned report is still available; try downloading again." });
  }
};
