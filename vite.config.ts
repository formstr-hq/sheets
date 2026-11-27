/// <reference types="vitest/config" />
import { fileURLToPath } from "node:url";
import { defineConfig } from "vite";
import { configDefaults } from "vitest/config";
import react from "@vitejs/plugin-react";

const here = (path: string) => fileURLToPath(new URL(path, import.meta.url));

export default defineConfig({
  plugins: [react()],
  base: "/",
  // Several suites build real encrypted sheets; under a loaded machine 5 s is too tight.
  // The long random-operation (fuzz) suites take 5-20 minutes each: they run only with FUZZ=1 (pnpm test:fuzz).
  test: {
    testTimeout: 30_000,
    exclude: process.env.FUZZ ? configDefaults.exclude : [...configDefaults.exclude, "src/lib/v2/advFuzz[ACDE].test.ts"],
  },
  resolve: {
    alias: [
      { find: "@formstr/sheets-formula", replacement: here("./packages/sheets-formula/src/index.ts") },
      { find: "@formstr/sheets-grid", replacement: here("./packages/sheets-grid/src/index.ts") },
      // --- Streaming xlsx import (src/lib/xlsxStream) ------------------------------------------
      // The import worker bundles a few ExcelJS parser files (worksheet-reader, styles-xform,
      // parse-sax, ...). They `require` Node modules they never execute on our code path, so the
      // browser build needs stand-ins. Without these Vite replaces the modules with an empty
      // "browser-external" proxy that throws on use. If these ever cause trouble for another
      // dependency, scope them to importers under /exceljs/lib/ with a resolveId plugin instead.
      { find: /^fs$/, replacement: here("./src/lib/xlsxStream/shims/empty.cjs") },
      { find: /^readable-stream$/, replacement: here("./src/lib/xlsxStream/shims/readable-stream.cjs") },
      // ExcelJS's reader extends EventEmitter. Use the npm `events` package, not the builtin stub.
      { find: /^events$/, replacement: here("./node_modules/events/events.js") },
    ],
  },
});
