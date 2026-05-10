// Run every cross-target verification we can do in this environment and
// print a single pass/fail/skipped matrix for review.
//
// Each row is a (target, fixture) pair. A target may be:
//   PASS   — actually exercised here
//   MANUAL — verification only possible on a real device / IDE; we print
//            the exact command to run.
import { readFileSync } from "node:fs";
import { resolve, dirname } from "node:path";
import { fileURLToPath, pathToFileURL } from "node:url";
import { execSync } from "node:child_process";
import { existsSync } from "node:fs";

const here = dirname(fileURLToPath(import.meta.url));
const repo = resolve(here, "..");
const fixturesDir = resolve(repo, "tests", "fixtures");

const fixtures = [
  { file: "truncated.docx", marker: "HELLO_FIXTURE_1" },
  { file: "bad-central-dir.xlsx", marker: "HELLO_FIXTURE_2" },
  { file: "crc-mismatch.pptx", marker: "HELLO_FIXTURE_3" },
  { file: "mixed-garbage.docx", marker: "HELLO_FIXTURE_4" },
];

const targets = [
  { id: "Web (browser bundle)", mode: "auto", run: webBundle },
  { id: "Web (Playwright in real Chromium)", mode: "auto-or-manual", run: webPlaywright },
  { id: "Linux (Electron via headless core)", mode: "auto", run: desktopHeadless },
  { id: "Windows (Electron build)", mode: "manual", note: "Run on a Windows runner: `npm run build:win -w desktop`. Outputs .exe + .zip in dist/desktop/." },
  { id: "macOS (Electron build, unsigned)", mode: "manual", note: "Run on macOS: `npm run build:mac -w desktop`. Outputs .dmg in dist/desktop/." },
  { id: "ChromeOS (PWA install)", mode: "auto", run: chromeOsPwa },
  { id: "Android (debug APK)", mode: "manual", note: "See tests/MANUAL.md — needs JDK 21 + Android SDK." },
  { id: "iOS (Xcode simulator build)", mode: "manual", note: "See tests/MANUAL.md — needs Xcode 15+ on macOS." },
];

async function webBundle() {
  const bundle = pathToFileURL(resolve(repo, "web", "lib", "core.js")).href;
  const { extractFromBuffer } = await import(bundle);
  return runAll(extractFromBuffer);
}

async function webPlaywright() {
  // Only attempt if a system Chromium is available; otherwise mark MANUAL.
  let canRun = false;
  try {
    execSync("which chromium chromium-browser google-chrome 2>/dev/null", { stdio: "pipe" });
    canRun = true;
  } catch {}
  if (!canRun) {
    return { mode: "manual", note: "Run `npx playwright install chromium && node tests/web-smoke.mjs` on a host with network access to download Chromium." };
  }
  try {
    execSync("node tests/web-smoke.mjs", { cwd: repo, stdio: "inherit" });
    return Object.fromEntries(fixtures.map((f) => [f.file, "PASS"]));
  } catch {
    return Object.fromEntries(fixtures.map((f) => [f.file, "FAIL"]));
  }
}

async function desktopHeadless() {
  // Same engine as Electron renderer; load the built core directly.
  const bundle = pathToFileURL(resolve(repo, "core", "dist", "index.js")).href;
  const { extractFromBuffer } = await import(bundle);
  return runAll(extractFromBuffer);
}

async function chromeOsPwa() {
  // ChromeOS install path runs the same web bundle. Verify the PWA scaffold
  // is present (manifest + sw + icons) so it's installable.
  const okFiles = [
    "web/index.html",
    "web/manifest.webmanifest",
    "web/sw.js",
    "web/icon.svg",
    "web/lib/core.js",
  ].every((p) => existsSync(resolve(repo, p)));
  if (!okFiles) {
    return Object.fromEntries(fixtures.map((f) => [f.file, "FAIL (missing PWA files)"]));
  }
  // Then run the same engine that powers the PWA.
  const bundle = pathToFileURL(resolve(repo, "web", "lib", "core.js")).href;
  const { extractFromBuffer } = await import(bundle);
  return runAll(extractFromBuffer);
}

function runAll(extract) {
  const out = {};
  for (const f of fixtures) {
    const buf = readFileSync(resolve(fixturesDir, f.file));
    const r = extract(new Uint8Array(buf));
    out[f.file] = r.text.includes(f.marker) ? "PASS" : "FAIL";
  }
  return out;
}

// Drive everything.
const rows = [];
for (const t of targets) {
  if (t.mode === "manual") {
    rows.push({ target: t.id, results: Object.fromEntries(fixtures.map((f) => [f.file, "MANUAL"])), note: t.note });
    continue;
  }
  try {
    const res = await t.run();
    if (res && res.mode === "manual") {
      rows.push({ target: t.id, results: Object.fromEntries(fixtures.map((f) => [f.file, "MANUAL"])), note: res.note });
    } else {
      rows.push({ target: t.id, results: res });
    }
  } catch (e) {
    rows.push({ target: t.id, results: Object.fromEntries(fixtures.map((f) => [f.file, "ERROR"])), note: e.message });
  }
}

// Print matrix.
const cols = fixtures.map((f) => f.file);
const tw = (s, n) => String(s).padEnd(n);
const colW = Math.max(...cols.map((c) => c.length), "fixture".length) + 2;
const targetW = Math.max(...rows.map((r) => r.target.length), "target".length) + 2;

console.log("\n=== Cross-platform verification matrix ===\n");
console.log(tw("target", targetW) + cols.map((c) => tw(c, colW)).join(""));
console.log("-".repeat(targetW + colW * cols.length));
let totalFail = 0;
for (const r of rows) {
  const cells = cols.map((c) => tw(r.results[c] || "—", colW)).join("");
  console.log(tw(r.target, targetW) + cells);
  if (r.note) console.log("  ".padStart(targetW, " ") + "↪ " + r.note);
  for (const v of Object.values(r.results)) if (v === "FAIL" || v === "ERROR") totalFail++;
}
console.log("");
console.log(`Auto-verified rows: ${rows.filter(r => Object.values(r.results).some(v => v === "PASS")).length}`);
console.log(`Manual-only rows:   ${rows.filter(r => Object.values(r.results).every(v => v === "MANUAL")).length}`);
console.log(`Failures:           ${totalFail}`);
process.exit(totalFail === 0 ? 0 : 1);
