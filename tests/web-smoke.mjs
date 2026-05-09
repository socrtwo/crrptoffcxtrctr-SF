// Headless browser smoke test for the static web build.
// Serves dist/web/ on a local port, drives Chromium with Playwright, and
// asserts the recovered text panel contains the fixture marker.
import { chromium } from "playwright";
import { createServer } from "node:http";
import { readFileSync, statSync } from "node:fs";
import { resolve, dirname, join, extname } from "node:path";
import { fileURLToPath } from "node:url";

const here = dirname(fileURLToPath(import.meta.url));
const repo = resolve(here, "..");
const root = resolve(repo, "dist", "web");
const fixturesDir = resolve(repo, "tests", "fixtures");

const MIME = {
  ".html": "text/html; charset=utf-8",
  ".js": "application/javascript; charset=utf-8",
  ".mjs": "application/javascript; charset=utf-8",
  ".webmanifest": "application/manifest+json; charset=utf-8",
  ".svg": "image/svg+xml",
  ".png": "image/png",
  ".map": "application/json; charset=utf-8",
};

const server = createServer((req, res) => {
  let urlPath = decodeURIComponent(req.url.split("?")[0]);
  if (urlPath === "/") urlPath = "/index.html";
  const fp = join(root, urlPath);
  if (!fp.startsWith(root)) {
    res.writeHead(403);
    return res.end();
  }
  try {
    const data = readFileSync(fp);
    res.writeHead(200, { "content-type": MIME[extname(fp)] || "application/octet-stream" });
    res.end(data);
  } catch {
    res.writeHead(404);
    res.end("not found");
  }
});

await new Promise((r) => server.listen(0, "127.0.0.1", r));
const port = server.address().port;
const url = `http://127.0.0.1:${port}/`;

const fixtures = [
  { file: "truncated.docx", marker: "HELLO_FIXTURE_1" },
  { file: "bad-central-dir.xlsx", marker: "HELLO_FIXTURE_2" },
  { file: "crc-mismatch.pptx", marker: "HELLO_FIXTURE_3" },
  { file: "mixed-garbage.docx", marker: "HELLO_FIXTURE_4" },
];

const browser = await chromium.launch();
const ctx = await browser.newContext();
const page = await ctx.newPage();

let failed = 0;
for (const f of fixtures) {
  await page.goto(url);
  // Service worker registers async; the page should be interactive immediately.
  await page.waitForSelector("#file");
  await page.setInputFiles("#file", resolve(fixturesDir, f.file));
  await page.waitForFunction(
    (sel) => document.querySelector(sel) && document.querySelector(sel).textContent.length > 0,
    "#text",
    { timeout: 5000 }
  );
  const text = await page.locator("#text").textContent();
  const ok = text.includes(f.marker);
  console.log(`  [web] ${ok ? "PASS" : "FAIL"}  ${f.file.padEnd(24)} marker=${f.marker}`);
  if (!ok) failed++;
}

await browser.close();
server.close();
process.exit(failed === 0 ? 0 : 1);
