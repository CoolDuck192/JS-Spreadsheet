import react from "@vitejs/plugin-react";
import { configDefaults, defineConfig } from "vitest/config";

export default defineConfig({
  plugins: [react()],
  build: {
    outDir: "dist/app",
    emptyOutDir: true
  },
  test: {
    environment: "jsdom",
    globals: true,
    setupFiles: "./src/test/setup.ts",
    exclude: [
      ...configDefaults.exclude,
      "tests/**",
      "**/.worktrees/**",
      "**/*.perf.test.ts",
      "**/*.perf.test.tsx"
    ],
    css: true,
    testTimeout: 10000
  }
});
