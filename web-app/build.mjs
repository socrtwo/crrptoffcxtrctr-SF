// Bundle the TypeScript core into a single ES module
// the browser can <script type="module">. Output goes to web/lib/core.js
// and is mirrored into dist/web/ for the release zip.
import { build } from "esbuild";
import { mkdirSync, copyFileSync, readdirSync, statSync, cpSync } from "node:fs";
import { resolve, dirname, join } from "node:path";
import { fileURLToPath } from "node:url";

const here = dirname(fileURLToPath(import.meta.url));
const repo = resolve(here, "..");

const webDir = resolve(repo, "web");
const distDir = resolve(repo, "dist", "web");
mkdirSync(resolve(webDir, "lib"), { recursive: true });
mkdirSync(distDir, { recursive: true });

await build({
  entryPoints: [resolve(repo, "core", "src", "index.ts")],
  bundle: true,
  format: "esm",
  target: ["es2020"],
  outfile: resolve(webDir, "lib", "core.js"),
  minify: true,
  sourcemap: true,
  logLevel: "info",
});

// Mirror everything in web/ to dist/web/
function mirror(src, dst) {
  mkdirSync(dst, { recursive: true });
  for (const ent of readdirSync(src)) {
    const s = join(src, ent);
    const d = join(dst, ent);
    if (statSync(s).isDirectory()) mirror(s, d);
    else copyFileSync(s, d);
  }
}
mirror(webDir, distDir);
console.log(`Web bundle written to ${webDir}/lib/core.js and mirrored to ${distDir}`);
