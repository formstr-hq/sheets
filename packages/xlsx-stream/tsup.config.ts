import { fileURLToPath } from "node:url";
import { defineConfig } from "tsup";

const shim = (name: string) => fileURLToPath(new URL(`./src/shims/${name}`, import.meta.url));

export default defineConfig([
  // Library entry: small, dependencies stay external (ESM + CJS, like the other @formstr packages).
  {
    entry: { index: "src/index.ts" },
    format: ["esm", "cjs"],
    dts: true,
    clean: true,
    sourcemap: true,
    target: "es2022",
  },
  // The worker is one self-contained file. ExcelJS's parser files, `events` and fflate are bundled
  // into it, and the Node modules ExcelJS requires but never runs here (`fs`, `readable-stream`)
  // are replaced with stubs, so a consumer needs no bundler aliases at all.
  {
    entry: { worker: "src/worker.ts" },
    format: ["esm"],
    platform: "browser",
    target: "es2022",
    clean: false,
    sourcemap: false,
    minify: true,
    noExternal: [/.*/],
    esbuildOptions(options) {
      options.alias = {
        ...options.alias,
        fs: shim("empty.cjs"),
        "readable-stream": shim("readable-stream.cjs"),
        events: fileURLToPath(new URL("./node_modules/events/events.js", import.meta.url)),
      };
    },
  },
]);
