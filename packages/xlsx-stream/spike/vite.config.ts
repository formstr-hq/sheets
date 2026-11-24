import { fileURLToPath } from "node:url";
import { defineConfig } from "vite";

// Deliberately NO ExcelJS aliases here: the package must be usable without them.
export default defineConfig({
  root: fileURLToPath(new URL(".", import.meta.url)),
  resolve: { alias: { "@formstr/xlsx-stream": fileURLToPath(new URL("../dist/index.js", import.meta.url)) } },
  server: { port: 5199, strictPort: true },
  preview: { port: 5199, strictPort: true },
  build: { outDir: "dist-spike", emptyOutDir: true },
});
