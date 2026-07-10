import react from "@vitejs/plugin-react";
import { defineConfig } from "vitest/config";

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
    exclude: ["node_modules/**", "dist/**", "tests/**"],
    css: true,
    testTimeout: 10000
  }
});
