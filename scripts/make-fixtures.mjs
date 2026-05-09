#!/usr/bin/env node
// Generate corrupt DOCX/XLSX/PPTX fixtures for testing the recovery core.
// Each file embeds a recoverable marker "HELLO_FIXTURE_<n>" so tests can
// assert that recovery actually pulled real content out of broken bytes.

import { zipSync, strToU8 } from "fflate";
import { mkdirSync, writeFileSync } from "node:fs";
import { fileURLToPath } from "node:url";
import { dirname, resolve } from "node:path";
import { randomBytes } from "node:crypto";

const here = dirname(fileURLToPath(import.meta.url));
const outDir = resolve(here, "..", "tests", "fixtures");
mkdirSync(outDir, { recursive: true });

const CONTENT_TYPES = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
<Default Extension="xml" ContentType="application/xml"/>
<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
<Override PartName="/xl/sharedStrings.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"/>
<Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
<Override PartName="/ppt/presentation.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"/>
<Override PartName="/ppt/slides/slide1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/>
</Types>`;

const ROOT_RELS = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>`;

function docxBytes(marker, opts = {}) {
  // Add filler at the end so the file is large enough that truncating the
  // tail only damages the central directory, not the marker run near the top.
  // Use a deterministic-but-varied PRNG so deflate can't trivially compress.
  let seed = 0xc0ffee;
  const rand = () => {
    seed = (seed * 1664525 + 1013904223) >>> 0;
    return seed;
  };
  const fillerParas = Array.from({ length: opts.fillerParas ?? 0 }, (_, i) => {
    let chars = "";
    for (let k = 0; k < 80; k++) {
      const c = 0x61 + (rand() % 26); // a-z
      chars += String.fromCharCode(c);
    }
    return `<w:p><w:r><w:t>${i}-${chars}</w:t></w:r></w:p>`;
  }).join("\n");
  const documentXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
<w:body>
<w:p><w:r><w:t>${marker}</w:t></w:r></w:p>
<w:p><w:r><w:t>This is a corrupt-fixture sample document.</w:t></w:r></w:p>
${fillerParas}
</w:body>
</w:document>`;
  return zipSync({
    "[Content_Types].xml": strToU8(CONTENT_TYPES),
    "_rels/.rels": strToU8(ROOT_RELS),
    "word/document.xml": strToU8(documentXml),
  });
}

function xlsxBytes(marker) {
  const workbookXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
  xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
<sheets><sheet name="Sheet1" sheetId="1" r:id="rId1"/></sheets>
</workbook>`;
  const sharedStringsXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="2" uniqueCount="2">
<si><t>${marker}</t></si>
<si><t>spreadsheet sample</t></si>
</sst>`;
  const sheet1Xml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
<sheetData>
<row r="1"><c r="A1" t="s"><v>0</v></c></row>
<row r="2"><c r="A2" t="s"><v>1</v></c></row>
</sheetData>
</worksheet>`;
  return zipSync({
    "[Content_Types].xml": strToU8(CONTENT_TYPES),
    "_rels/.rels": strToU8(ROOT_RELS),
    "xl/workbook.xml": strToU8(workbookXml),
    "xl/sharedStrings.xml": strToU8(sharedStringsXml),
    "xl/worksheets/sheet1.xml": strToU8(sheet1Xml),
  });
}

function pptxBytes(marker) {
  const presentationXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:presentation xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"/>`;
  const slide1Xml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:sld xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
  xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
<p:cSld><p:spTree>
<p:sp><p:txBody>
<a:p><a:r><a:t>${marker}</a:t></a:r></a:p>
<a:p><a:r><a:t>presentation sample</a:t></a:r></a:p>
</p:txBody></p:sp>
</p:spTree></p:cSld>
</p:sld>`;
  return zipSync({
    "[Content_Types].xml": strToU8(CONTENT_TYPES),
    "_rels/.rels": strToU8(ROOT_RELS),
    "ppt/presentation.xml": strToU8(presentationXml),
    "ppt/slides/slide1.xml": strToU8(slide1Xml),
  });
}

// --- Corruption strategies ---

// 1) Truncate the last 2 KB.
function truncate(bytes, n = 2048) {
  if (bytes.length <= n) throw new Error("File too small to truncate");
  return bytes.slice(0, bytes.length - n);
}

// 2) Overwrite the End-of-Central-Directory record (and the central
//    directory entries before it) with zeros, leaving local headers intact.
function wipeCentralDirectory(bytes) {
  const out = new Uint8Array(bytes);
  // Find EOCD signature 0x06054b50 from the end.
  let eocd = -1;
  for (let i = out.length - 22; i >= 0 && i >= out.length - 65557; i--) {
    if (
      out[i] === 0x50 &&
      out[i + 1] === 0x4b &&
      out[i + 2] === 0x05 &&
      out[i + 3] === 0x06
    ) {
      eocd = i;
      break;
    }
  }
  if (eocd < 0) throw new Error("EOCD not found");
  // Find first central directory header signature.
  let cd = -1;
  for (let i = 0; i < out.length - 4; i++) {
    if (
      out[i] === 0x50 &&
      out[i + 1] === 0x4b &&
      out[i + 2] === 0x01 &&
      out[i + 3] === 0x02
    ) {
      cd = i;
      break;
    }
  }
  if (cd < 0) throw new Error("Central directory not found");
  // Zero from start of CD through end of file (covers CD entries + EOCD).
  for (let i = cd; i < out.length; i++) out[i] = 0;
  return out;
}

// 3) Flip a single byte in the compressed data of one entry so the CRC
//    check fails. We pick an entry by name, find its local header, then
//    flip a byte within the data region.
function corruptEntryCrc(bytes, entryName) {
  const out = new Uint8Array(bytes);
  const target = new TextEncoder().encode(entryName);
  // Find local file header signature followed (at +30) by the matching name.
  for (let i = 0; i < out.length - 30 - target.length; i++) {
    if (
      out[i] === 0x50 &&
      out[i + 1] === 0x4b &&
      out[i + 2] === 0x03 &&
      out[i + 3] === 0x04
    ) {
      const nameLen = out[i + 26] | (out[i + 27] << 8);
      const extraLen = out[i + 28] | (out[i + 29] << 8);
      const compSize =
        (out[i + 18] |
          (out[i + 19] << 8) |
          (out[i + 20] << 16) |
          (out[i + 21] << 24)) >>>
        0;
      if (nameLen !== target.length) continue;
      let match = true;
      for (let k = 0; k < nameLen; k++) {
        if (out[i + 30 + k] !== target[k]) {
          match = false;
          break;
        }
      }
      if (!match) continue;
      const dataStart = i + 30 + nameLen + extraLen;
      if (compSize < 4) throw new Error("Entry too small to corrupt");
      // Flip a byte in the middle (avoid the deflate header bytes at index 0/1).
      const flipAt = dataStart + Math.floor(compSize / 2);
      out[flipAt] ^= 0xff;
      return out;
    }
  }
  throw new Error(`Local header not found for entry ${entryName}`);
}

// 4) Prepend N bytes of random garbage. A valid OOXML file should still be
//    recoverable because we scan for LFH signatures.
function prependGarbage(bytes, n = 1024) {
  const garbage = randomBytes(n);
  // Make sure the first 4 bytes are NOT a PK\03\04 signature.
  garbage[0] = 0x00;
  garbage[1] = 0x00;
  garbage[2] = 0x00;
  garbage[3] = 0x00;
  const out = new Uint8Array(garbage.length + bytes.length);
  out.set(garbage, 0);
  out.set(bytes, garbage.length);
  return out;
}

// --- Build all fixtures ---

const fixtures = [
  // Pad the docx so it's larger than the 2 KB we chop off the tail.
  { file: "truncated.docx", marker: "HELLO_FIXTURE_1", build: () => truncate(docxBytes("HELLO_FIXTURE_1", { fillerParas: 200 })) },
  { file: "bad-central-dir.xlsx", marker: "HELLO_FIXTURE_2", build: () => wipeCentralDirectory(xlsxBytes("HELLO_FIXTURE_2")) },
  { file: "crc-mismatch.pptx", marker: "HELLO_FIXTURE_3", build: () => corruptEntryCrc(pptxBytes("HELLO_FIXTURE_3"), "ppt/presentation.xml") },
  { file: "mixed-garbage.docx", marker: "HELLO_FIXTURE_4", build: () => prependGarbage(docxBytes("HELLO_FIXTURE_4")) },
];

const summary = [];
for (const f of fixtures) {
  const bytes = f.build();
  const path = resolve(outDir, f.file);
  writeFileSync(path, bytes);
  summary.push({ file: f.file, marker: f.marker, size: bytes.length });
}

// Also drop a manifest so other tools can discover the fixtures.
writeFileSync(
  resolve(outDir, "manifest.json"),
  JSON.stringify({ generated: new Date().toISOString(), fixtures: summary }, null, 2)
);

console.log("Generated fixtures:");
for (const s of summary) console.log(`  ${s.file.padEnd(24)} ${s.marker.padEnd(20)} ${s.size} bytes`);
