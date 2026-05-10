import { scanZip, type ZipEntry, type ScanResult } from "./zip-scan.js";

export type DocumentKind = "docx" | "xlsx" | "pptx" | "unknown";

export interface ExtractedImage {
  name: string;
  data: Uint8Array;
}

export interface ExtractionResult {
  kind: DocumentKind;
  text: string;
  images: ExtractedImage[];
  rawXml: Record<string, string>;
  scan: ScanResult;
}

const TEXT_DECODER = new TextDecoder("utf-8", { fatal: false });

function decodeXml(data: Uint8Array): string {
  // Strip BOM if present.
  if (data.length >= 3 && data[0] === 0xef && data[1] === 0xbb && data[2] === 0xbf) {
    data = data.subarray(3);
  }
  return TEXT_DECODER.decode(data);
}

const W_T_RE = /<w:t(?:\s[^>]*)?>([^<]*)<\/w:t>/g;
const A_T_RE = /<a:t(?:\s[^>]*)?>([^<]*)<\/a:t>/g;
const SI_T_RE = /<t(?:\s[^>]*)?>([^<]*)<\/t>/g; // shared strings <si><t>
const PARA_BREAK_RE = /<\/w:p>/g;
const ROW_BREAK_RE = /<\/row>/g;
const SLIDE_BREAK_RE = /<\/p:sld>/g;

function unescapeXml(s: string): string {
  return s
    .replace(/&lt;/g, "<")
    .replace(/&gt;/g, ">")
    .replace(/&quot;/g, '"')
    .replace(/&apos;/g, "'")
    .replace(/&amp;/g, "&");
}

function extractWithRegex(xml: string, re: RegExp): string[] {
  const out: string[] = [];
  let m: RegExpExecArray | null;
  re.lastIndex = 0;
  while ((m = re.exec(xml)) !== null) out.push(unescapeXml(m[1]));
  return out;
}

function detectKind(entries: ZipEntry[]): DocumentKind {
  const names = new Set(entries.map((e) => e.name.toLowerCase()));
  if (names.has("word/document.xml")) return "docx";
  if (names.has("xl/workbook.xml") || names.has("xl/sharedstrings.xml")) return "xlsx";
  for (const n of names) if (n.startsWith("ppt/slides/slide")) return "pptx";
  // Fall back: look for substrings (some entries may be partially named).
  for (const n of names) {
    if (n.includes("word/document")) return "docx";
    if (n.includes("xl/workbook") || n.includes("sharedstrings")) return "xlsx";
    if (n.includes("ppt/slides/slide")) return "pptx";
  }
  return "unknown";
}

/**
 * Pull text from a (possibly corrupt) DOCX/XLSX/PPTX file.
 * Returns extracted text, embedded images, and the raw XML files we found.
 */
export function extractFromBuffer(buf: Uint8Array): ExtractionResult {
  const scan = scanZip(buf);
  const kind = detectKind(scan.entries);
  const rawXml: Record<string, string> = {};
  const images: ExtractedImage[] = [];

  for (const e of scan.entries) {
    if (!e.data) continue;
    const lname = e.name.toLowerCase();
    if (lname.endsWith(".xml") || lname.endsWith(".rels")) {
      rawXml[e.name] = decodeXml(e.data);
    }
    if (
      lname.startsWith("word/media/") ||
      lname.startsWith("xl/media/") ||
      lname.startsWith("ppt/media/")
    ) {
      images.push({ name: e.name, data: e.data });
    }
  }

  let text = "";
  if (kind === "docx") {
    const doc = rawXml["word/document.xml"];
    if (doc) {
      // Split paragraphs first, then extract <w:t> per paragraph.
      const paras = doc.split(PARA_BREAK_RE);
      const lines: string[] = [];
      for (const p of paras) {
        const runs = extractWithRegex(p, W_T_RE);
        if (runs.length) lines.push(runs.join(""));
      }
      text = lines.join("\n");
    }
  } else if (kind === "xlsx") {
    const ss = rawXml["xl/sharedStrings.xml"] || rawXml["xl/sharedstrings.xml"];
    const lines: string[] = [];
    if (ss) {
      // Each <si>…<t>X</t>…</si> is a single shared string.
      const siRe = /<si[\s>][\s\S]*?<\/si>/g;
      let m: RegExpExecArray | null;
      while ((m = siRe.exec(ss)) !== null) {
        const tParts = extractWithRegex(m[0], SI_T_RE);
        lines.push(tParts.join(""));
      }
    }
    // Also pull inline strings from sheet XMLs.
    for (const [name, xml] of Object.entries(rawXml)) {
      if (/^xl\/worksheets\/sheet\d+\.xml$/i.test(name)) {
        const inline = extractWithRegex(xml, /<is>[\s\S]*?<t(?:\s[^>]*)?>([^<]*)<\/t>[\s\S]*?<\/is>/g);
        lines.push(...inline);
      }
    }
    text = lines.filter(Boolean).join("\n");
  } else if (kind === "pptx") {
    const lines: string[] = [];
    const slideNames = Object.keys(rawXml)
      .filter((n) => /^ppt\/slides\/slide\d+\.xml$/i.test(n))
      .sort();
    for (const n of slideNames) {
      const runs = extractWithRegex(rawXml[n], A_T_RE);
      if (runs.length) lines.push(runs.join(" "));
    }
    text = lines.join("\n");
  } else {
    // Unknown kind — last-ditch: regex over all decoded XML for any <*:t> runs.
    const lines: string[] = [];
    for (const xml of Object.values(rawXml)) {
      lines.push(...extractWithRegex(xml, /<(?:w|a):t(?:\s[^>]*)?>([^<]*)<\/(?:w|a):t>/g));
    }
    text = lines.join("\n");
  }

  return { kind, text, images, rawXml, scan };
}
