// In-sandbox web verification: import the *exact* web/lib/core.js bundle the
// browser loads, and run each fixture through it. If a system Chromium isn't
// available for the Playwright test, this still proves the shipped bundle
// extracts text correctly.
import { readFileSync } from "node:fs";
import { resolve, dirname } from "node:path";
import { fileURLToPath, pathToFileURL } from "node:url";

const here = dirname(fileURLToPath(import.meta.url));
const repo = resolve(here, "..");
const bundle = pathToFileURL(resolve(repo, "web", "lib", "core.js")).href;

const { extractFromBuffer } = await import(bundle);

const fixtures = [
  { file: "truncated.docx", marker: "HELLO_FIXTURE_1" },
  { file: "bad-central-dir.xlsx", marker: "HELLO_FIXTURE_2" },
  { file: "crc-mismatch.pptx", marker: "HELLO_FIXTURE_3" },
  { file: "mixed-garbage.docx", marker: "HELLO_FIXTURE_4" },
];

let failed = 0;
for (const f of fixtures) {
  const path = resolve(repo, "tests", "fixtures", f.file);
  const buf = readFileSync(path);
  const r = extractFromBuffer(new Uint8Array(buf));
  const ok = r.text.includes(f.marker);
  console.log(`  [web-bundle] ${ok ? "PASS" : "FAIL"}  ${f.file.padEnd(24)} kind=${r.kind}`);
  if (!ok) failed++;
}
process.exit(failed === 0 ? 0 : 1);
