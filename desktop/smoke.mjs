// Headless smoke test: simulates what the Electron renderer does end-to-end
// (read the file from disk, run it through the bundled core, check the marker)
// without actually launching the windowed Electron app — useful for CI on
// Linux runners where there's no display.
import { readFileSync } from "node:fs";
import { resolve, dirname } from "node:path";
import { fileURLToPath } from "node:url";
import { extractFromBuffer } from "@crrptoffcxtrctr/core";

const here = dirname(fileURLToPath(import.meta.url));
const repo = resolve(here, "..");
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
  console.log(`  [desktop] ${ok ? "PASS" : "FAIL"}  ${f.file.padEnd(24)} kind=${r.kind} entries=${r.scan.entries.length}`);
  if (!ok) failed++;
}
process.exit(failed === 0 ? 0 : 1);
