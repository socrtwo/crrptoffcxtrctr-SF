import { immortalInflate } from "./immortal-inflate.js";

export interface ZipEntry {
  name: string;
  compressedSize: number;
  uncompressedSize: number;
  crc32: number;
  compressionMethod: number;
  localHeaderOffset: number;
  dataOffset: number;
  flags: number;
  data?: Uint8Array;
  decompressionError?: string;
  crcMismatch?: boolean;
}

export interface ScanResult {
  entries: ZipEntry[];
  warnings: string[];
}

const LFH_SIG = 0x04034b50;
const CD_SIG = 0x02014b50;
const EOCD_SIG = 0x06054b50;
const DD_SIG = 0x08074b50;

function readU16(buf: Uint8Array, off: number): number {
  return buf[off] | (buf[off + 1] << 8);
}
function readU32(buf: Uint8Array, off: number): number {
  return (
    (buf[off] |
      (buf[off + 1] << 8) |
      (buf[off + 2] << 16) |
      (buf[off + 3] << 24)) >>>
    0
  );
}

function findSignature(buf: Uint8Array, sig: number, start = 0): number {
  const b0 = sig & 0xff;
  const b1 = (sig >>> 8) & 0xff;
  const b2 = (sig >>> 16) & 0xff;
  const b3 = (sig >>> 24) & 0xff;
  for (let i = start; i <= buf.length - 4; i++) {
    if (buf[i] === b0 && buf[i + 1] === b1 && buf[i + 2] === b2 && buf[i + 3] === b3) {
      return i;
    }
  }
  return -1;
}

const CRC_TABLE = (() => {
  const t = new Uint32Array(256);
  for (let n = 0; n < 256; n++) {
    let c = n;
    for (let k = 0; k < 8; k++) c = c & 1 ? 0xedb88320 ^ (c >>> 1) : c >>> 1;
    t[n] = c >>> 0;
  }
  return t;
})();

function crc32(buf: Uint8Array): number {
  let c = 0xffffffff;
  for (let i = 0; i < buf.length; i++) c = CRC_TABLE[(c ^ buf[i]) & 0xff] ^ (c >>> 8);
  return (c ^ 0xffffffff) >>> 0;
}

function decodeName(raw: Uint8Array, utf8: boolean): string {
  if (utf8) return new TextDecoder("utf-8", { fatal: false }).decode(raw);
  // CP437-ish fallback — good enough for ASCII filenames found in OOXML.
  return new TextDecoder("utf-8", { fatal: false }).decode(raw);
}

/**
 * Scan a (possibly corrupt) zip by walking Local File Headers directly,
 * regardless of whether the central directory is present or valid.
 *
 * Mirrors the strategy of Angus Johnson's TZipFix: don't trust the EOCD/CD,
 * scan the body for 0x04034b50 signatures and reconstruct entries.
 */
export function scanZip(buf: Uint8Array, opts: { decompress?: boolean } = {}): ScanResult {
  const decompress = opts.decompress !== false;
  const entries: ZipEntry[] = [];
  const warnings: string[] = [];

  let pos = 0;
  while (pos <= buf.length - 4) {
    const at = findSignature(buf, LFH_SIG, pos);
    if (at < 0) break;
    if (at + 30 > buf.length) {
      warnings.push(`Truncated local file header at offset ${at}`);
      break;
    }
    const flags = readU16(buf, at + 6);
    const method = readU16(buf, at + 8);
    const crc = readU32(buf, at + 14);
    let compressedSize = readU32(buf, at + 18);
    let uncompressedSize = readU32(buf, at + 22);
    const nameLen = readU16(buf, at + 26);
    const extraLen = readU16(buf, at + 28);
    const nameStart = at + 30;
    const dataOffset = nameStart + nameLen + extraLen;
    if (dataOffset > buf.length) {
      warnings.push(`Local header at ${at} extends past EOF`);
      break;
    }
    const utf8 = (flags & 0x800) !== 0;
    const name = decodeName(buf.subarray(nameStart, nameStart + nameLen), utf8);

    // Bit 3 of flags means sizes/CRC are zero in the header and live in a
    // data descriptor following the data. We have to find the next LFH/CD
    // signature or DD_SIG to determine the compressed size.
    const hasDataDescriptor = (flags & 0x08) !== 0;
    if (hasDataDescriptor && (compressedSize === 0 || uncompressedSize === 0)) {
      const nextLfh = findSignature(buf, LFH_SIG, dataOffset);
      const nextCd = findSignature(buf, CD_SIG, dataOffset);
      const candidates = [nextLfh, nextCd].filter((x) => x > 0);
      const next = candidates.length ? Math.min(...candidates) : buf.length;
      // Try to find a data descriptor signature just before "next".
      // The DD is 12 bytes (no sig) or 16 bytes (with sig). We search for the
      // signature variant first.
      let ddPos = -1;
      for (let i = next - 16; i >= dataOffset; i--) {
        if (i + 4 <= buf.length && readU32(buf, i) === DD_SIG) {
          ddPos = i;
          break;
        }
      }
      if (ddPos >= 0) {
        compressedSize = ddPos - dataOffset;
        // CRC at ddPos+4, compSize at ddPos+8, uncompSize at ddPos+12.
        uncompressedSize = readU32(buf, ddPos + 12);
      } else {
        compressedSize = next - dataOffset;
      }
    }

    const dataEnd = Math.min(dataOffset + compressedSize, buf.length);
    if (dataEnd < dataOffset + compressedSize) {
      warnings.push(`Entry "${name}" data truncated at EOF (wanted ${compressedSize} bytes, got ${dataEnd - dataOffset})`);
    }

    const entry: ZipEntry = {
      name,
      compressedSize: dataEnd - dataOffset,
      uncompressedSize,
      crc32: crc,
      compressionMethod: method,
      localHeaderOffset: at,
      dataOffset,
      flags,
    };

    if (decompress) {
      const raw = buf.subarray(dataOffset, dataEnd);
      if (method === 0) {
        entry.data = raw.slice();
      } else if (method === 8) {
        // Fault-tolerant inflate: never throws, returns whatever it could
        // decode plus an `isCorrupt` flag for truncated/damaged DEFLATE.
        const { data: inflated, isCorrupt } = immortalInflate(raw);
        if (inflated.length > 0) {
          entry.data = inflated;
          if (isCorrupt) {
            entry.decompressionError = `Partial recovery (${inflated.length} bytes): corrupt or truncated DEFLATE stream`;
          }
        } else {
          entry.decompressionError = isCorrupt
            ? "Corrupt or truncated DEFLATE stream (no bytes recovered)"
            : "Empty DEFLATE stream";
        }
      } else {
        entry.decompressionError = `Unsupported compression method ${method}`;
      }
      if (entry.data && crc !== 0) {
        const actual = crc32(entry.data);
        if (actual !== crc) entry.crcMismatch = true;
      }
    }

    entries.push(entry);
    pos = dataEnd > 0 ? dataEnd : at + 4;
  }

  if (!entries.length) warnings.push("No local file headers found");
  return { entries, warnings };
}
