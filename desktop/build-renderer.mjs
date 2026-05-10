// Copy the built web assets into desktop/renderer/ so Electron can load them
// directly from disk (no localhost server required).
import { mkdirSync, cpSync, existsSync } from "node:fs";
import { resolve, dirname } from "node:path";
import { fileURLToPath } from "node:url";

const here = dirname(fileURLToPath(import.meta.url));
const repo = resolve(here, "..");
const src = resolve(repo, "web");
const dst = resolve(here, "renderer");

if (!existsSync(resolve(src, "lib", "core.js"))) {
  console.error("Web bundle missing at web/lib/core.js. Run `npm run build:web` first.");
  process.exit(1);
}

mkdirSync(dst, { recursive: true });
cpSync(src, dst, { recursive: true });
console.log(`Copied web/ → ${dst}`);
