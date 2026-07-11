import { defineConfig, devices } from "@playwright/test";

const externalBaseURL = process.env.E2E_BASE_URL?.trim();

export default defineConfig({
  testDir: "./tests",
  timeout: 30_000,
  use: {
    baseURL: externalBaseURL || "http://127.0.0.1:5173",
    trace: "on-first-retry"
  },
  ...(externalBaseURL
    ? {}
    : {
        webServer: {
          command: "pnpm dev",
          url: "http://127.0.0.1:5173",
          reuseExistingServer: !process.env.CI
        }
      }),
  projects: [{ name: "chromium", use: { ...devices["Desktop Chrome"] } }]
});
