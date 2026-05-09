// Copy the built web bundle into mobile/www so `npx cap sync` ships it into
// the Android and iOS native projects.
import { mkdirSync, cpSync, existsSync, rmSync } from "node:fs";
import { resolve, dirname } from "node:path";
import { fileURLToPath } from "node:url";

const here = dirname(fileURLToPath(import.meta.url));
const repo = resolve(here, "..");
const src = resolve(repo, "web");
const dst = resolve(here, "www");

if (!existsSync(resolve(src, "lib", "core.js"))) {
  console.error("Web bundle missing at web/lib/core.js. Run `npm run build:web` first.");
  process.exit(1);
}

if (existsSync(dst)) rmSync(dst, { recursive: true, force: true });
mkdirSync(dst, { recursive: true });
cpSync(src, dst, { recursive: true });
console.log(`Copied web/ → ${dst}`);
