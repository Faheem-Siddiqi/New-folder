import { scanTurnover } from "./turnover";
import type { StrengthData } from "./strength-data";

self.onmessage = (event: MessageEvent<{ buffer: ArrayBuffer; fileName: string; template: StrengthData }>) => {
  try {
    const result = scanTurnover(event.data.buffer, event.data.fileName, event.data.template, (step, percent) => self.postMessage({ type: "progress", step, percent }));
    self.postMessage({ type: "complete", result });
  } catch (error) {
    self.postMessage({ type: "error", message: error instanceof Error ? error.message : "The workbook could not be processed. Save a new Excel copy and try again." });
  }
};
