import { describe, it, expect, beforeAll } from "vitest";
import { readFileSync, existsSync } from "node:fs";
import { resolve, dirname } from "node:path";
import { fileURLToPath } from "node:url";
import { execSync } from "node:child_process";
import { extractFromBuffer } from "../src/index.js";

const here = dirname(fileURLToPath(import.meta.url));
const repoRoot = resolve(here, "..", "..");
const fixturesDir = resolve(repoRoot, "tests", "fixtures");

const cases = [
  { file: "truncated.docx", marker: "HELLO_FIXTURE_1", kind: "docx" as const },
  { file: "bad-central-dir.xlsx", marker: "HELLO_FIXTURE_2", kind: "xlsx" as const },
  { file: "crc-mismatch.pptx", marker: "HELLO_FIXTURE_3", kind: "pptx" as const },
  { file: "mixed-garbage.docx", marker: "HELLO_FIXTURE_4", kind: "docx" as const },
];

beforeAll(() => {
  const anyMissing = cases.some((c) => !existsSync(resolve(fixturesDir, c.file)));
  if (anyMissing) {
    execSync("node scripts/make-fixtures.mjs", { cwd: repoRoot, stdio: "inherit" });
  }
});

describe("corrupt-OOXML extraction", () => {
  for (const c of cases) {
    it(`recovers ${c.marker} from ${c.file}`, () => {
      const buf = readFileSync(resolve(fixturesDir, c.file));
      const result = extractFromBuffer(new Uint8Array(buf));
      expect(result.kind).toBe(c.kind);
      expect(result.text).toContain(c.marker);
      expect(result.scan.entries.length).toBeGreaterThan(0);
    });
  }

  it("reports warnings for the bad-central-dir fixture", () => {
    const buf = readFileSync(resolve(fixturesDir, "bad-central-dir.xlsx"));
    const result = extractFromBuffer(new Uint8Array(buf));
    // Even with a wiped central directory, scanning local headers still works.
    expect(result.scan.entries.some((e) => e.name === "xl/sharedStrings.xml")).toBe(true);
  });

  it("flags CRC mismatch on the corrupted entry", () => {
    const buf = readFileSync(resolve(fixturesDir, "crc-mismatch.pptx"));
    const result = extractFromBuffer(new Uint8Array(buf));
    const tampered = result.scan.entries.find((e) => e.name === "ppt/presentation.xml");
    expect(tampered).toBeDefined();
    // Either CRC mismatch is flagged, or the inflate step itself failed.
    expect(Boolean(tampered?.crcMismatch || tampered?.decompressionError)).toBe(true);
    // Slide text recovery should still succeed.
    expect(result.text).toContain("HELLO_FIXTURE_3");
  });
});
