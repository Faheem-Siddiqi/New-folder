import { resolve } from "node:path";

// All app commands run from strength-analysis; the template lives at workspace root.
export const STRENGTH_FILE = resolve(process.cwd(), "..", "strength.json");
