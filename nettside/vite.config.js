import { sveltekit } from "@sveltejs/kit/vite";
import { defineConfig } from "vite";
import path from "path";

export default defineConfig({
  plugins: [sveltekit()],
  resolve: {
    alias: {
      "$evidence/config": path.resolve("./src/evidence/config.js"),
      "$evidence/themes": path.resolve("./src/evidence/themes.js")
    }
  },
  optimizeDeps: {
    exclude: [
      "@evidence-dev/core-components",
      "@evidence-dev/sdk",
      "@evidence-dev/universal-sql",
      "@evidence-dev/universal-sql/client-duckdb",
      "@duckdb/duckdb-wasm"
    ],
    include: ["echarts", "echarts-stat"]
  }
});
