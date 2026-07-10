import { resolve } from "node:path";
import { fileURLToPath } from "node:url";
import react from "@vitejs/plugin-react";
import { defineConfig } from "vite";

const rootDir = fileURLToPath(new URL(".", import.meta.url));

export default defineConfig({
  plugins: [react()],
  build: {
    outDir: "dist/lib",
    emptyOutDir: true,
    sourcemap: true,
    cssCodeSplit: true,
    lib: {
      entry: {
        core: resolve(rootDir, "src/entry/core.ts"),
        react: resolve(rootDir, "src/entry/react.ts"),
        "connectors/google": resolve(rootDir, "src/entry/google.ts")
      },
      formats: ["es"]
    },
    rollupOptions: {
      input: {
        core: resolve(rootDir, "src/entry/core.ts"),
        react: resolve(rootDir, "src/entry/react.ts"),
        "connectors/google": resolve(rootDir, "src/entry/google.ts"),
        styles: resolve(rootDir, "src/entry/styles.css")
      },
      external: ["react", "react-dom", "react/jsx-runtime", "exceljs", "hyperformula"],
      output: {
        entryFileNames: (chunk) => `${chunk.name}.js`,
        chunkFileNames: "chunks/[name]-[hash].js",
        assetFileNames: (asset) => asset.name?.endsWith(".css") ? "styles.css" : "assets/[name]-[hash][extname]"
      }
    }
  }
});
