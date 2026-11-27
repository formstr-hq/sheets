// Runs the long random-operation suites (5-20 minutes each), which the normal `pnpm test` leaves out.
import { spawnSync } from "node:child_process";

const run = spawnSync("pnpm", ["exec", "vitest", "run", "src/lib/v2/advFuzz"], {
  stdio: "inherit",
  shell: true,
  env: { ...process.env, FUZZ: "1" },
});
process.exit(run.status ?? 1);
