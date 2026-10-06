import { defineConfig } from "@playwright/test";

export default defineConfig({
  testDir: "./tests",
  timeout: 60_000,
  workers: 1,
  use: { baseURL: "http://127.0.0.1:3100", channel: "chrome", headless: true, viewport: { width: 1440, height: 1000 }, trace: "retain-on-failure" },
  webServer: { command: "npm run start -- --port 3100", url: "http://127.0.0.1:3100", reuseExistingServer: true, timeout: 60_000 },
});
