// immortal-inflate.ts
//
// Fault-tolerant DEFLATE (RFC 1951) decompressor — a TypeScript ES-module port
// of the "Immortal Inflater" from socrtwo/Universal-File-Repair-Tool.
//
// Unlike an ordinary inflater (e.g. fflate), `immortalInflate` NEVER throws on
// corrupt or truncated input: it decodes as much as it can and returns whatever
// bytes it produced together with an `isCorrupt` flag. This is exactly what the
// corrupt-Office extractor needs when recovering damaged .docx/.xlsx/.pptx ZIP
// members, where the DEFLATE stream may be cut off or contain bad blocks.
//
// The algorithm (BitStream + fixed/dynamic Huffman + fault tolerance) is kept
// identical to the frozen reference implementation; only types and ES-module
// packaging were added.

/** Result of a fault-tolerant inflate: best-effort bytes plus a corruption flag. */
export interface ImmortalInflateResult {
  data: Uint8Array;
  isCorrupt: boolean;
}

// ---- Bit stream ----
class BitStream {
  buf: Uint8Array;
  pos: number;
  bit: number;
  len: number;

  constructor(u8: Uint8Array) {
    this.buf = u8;
    this.pos = 0;
    this.bit = 0;
    this.len = u8.length;
  }

  /** Read `n` bits LSB-first. Returns -1 if the stream is exhausted. */
  read(n: number): number {
    let v = 0;
    for (let i = 0; i < n; i++) {
      if (this.pos >= this.len) return -1;
      v |= ((this.buf[this.pos] >>> this.bit) & 1) << i;
      this.bit++;
      if (this.bit === 8) {
        this.bit = 0;
        this.pos++;
      }
    }
    return v;
  }

  align(): void {
    if (this.bit !== 0) {
      this.bit = 0;
      this.pos++;
    }
  }
}

// ---- Fixed Huffman tables ----
const FIXED_LIT = new Uint8Array(288);
for (let i = 0; i < 144; i++) FIXED_LIT[i] = 8;
for (let i = 144; i < 256; i++) FIXED_LIT[i] = 9;
for (let i = 256; i < 280; i++) FIXED_LIT[i] = 7;
for (let i = 280; i < 288; i++) FIXED_LIT[i] = 8;
const FIXED_DIST = new Uint8Array(32);
for (let i = 0; i < 32; i++) FIXED_DIST[i] = 5;

const CLEN_ORDER = [16, 17, 18, 0, 8, 7, 9, 6, 10, 5, 11, 4, 12, 3, 13, 2, 14, 1, 15];
const LEN_BASE = [
  3, 4, 5, 6, 7, 8, 9, 10, 11, 13, 15, 17, 19, 23, 27, 31, 35, 43, 51, 59, 67, 83, 99, 115, 131,
  163, 195, 227, 258,
];
const LEN_EXTRA = [
  0, 0, 0, 0, 0, 0, 0, 0, 1, 1, 1, 1, 2, 2, 2, 2, 3, 3, 3, 3, 4, 4, 4, 4, 5, 5, 5, 5, 0,
];
const DIST_BASE = [
  1, 2, 3, 4, 5, 7, 9, 13, 17, 25, 33, 49, 65, 97, 129, 193, 257, 385, 513, 769, 1025, 1537, 2049,
  3073, 4097, 6145, 8193, 12289, 16385, 24577,
];
const DIST_EXTRA = [
  0, 0, 0, 0, 1, 1, 2, 2, 3, 3, 4, 4, 5, 5, 6, 6, 7, 7, 8, 8, 9, 9, 10, 10, 11, 11, 12, 12, 13, 13,
];

interface HuffTree {
  map: Record<number, number>;
  maxLen: number;
}

function buildTree(lengths: ArrayLike<number>): HuffTree | null {
  const counts = new Int32Array(16);
  const nextCode = new Int32Array(16);
  let maxLen = 0;
  for (let i = 0; i < lengths.length; i++) {
    counts[lengths[i]]++;
    if (lengths[i] > maxLen) maxLen = lengths[i];
  }
  if (maxLen === 0) return null;
  let code = 0;
  counts[0] = 0;
  for (let i = 1; i <= 15; i++) {
    code = (code + counts[i - 1]) << 1;
    nextCode[i] = code;
  }
  const map: Record<number, number> = {};
  for (let i = 0; i < lengths.length; i++) {
    const len = lengths[i];
    if (len !== 0) {
      map[(len << 16) | nextCode[len]] = i;
      nextCode[len]++;
    }
  }
  return { map, maxLen };
}

function decodeSym(s: BitStream, t: HuffTree): number {
  let c = 0;
  for (let l = 1; l <= t.maxLen; l++) {
    const b = s.read(1);
    if (b === -1) return -1;
    c = (c << 1) | b;
    const k = (l << 16) | c;
    if (t.map[k] !== undefined) return t.map[k];
  }
  return -2;
}

/**
 * Fault-tolerant raw-DEFLATE inflate. Decodes as much as possible and returns
 * the produced bytes; `isCorrupt` is true if the stream was truncated or
 * contained an undecodable block. Never throws.
 */
export function immortalInflate(data: Uint8Array): ImmortalInflateResult {
  const s = new BitStream(data);
  const out: number[] = [];
  let bfinal = 0;
  let corrupted = false;
  try {
    while (!bfinal) {
      bfinal = s.read(1);
      const btype = s.read(2);
      if (bfinal === -1 || btype === -1) {
        corrupted = true;
        break;
      }
      if (btype === 0) {
        s.align();
        const len = s.read(16);
        s.read(16); // nlen (unused; tolerate corruption)
        if (len === -1) {
          corrupted = true;
          break;
        }
        for (let i = 0; i < len; i++) out.push(s.buf[s.pos++] || 0);
      } else if (btype === 1 || btype === 2) {
        let lt: HuffTree | null;
        let dt: HuffTree | null;
        if (btype === 1) {
          lt = buildTree(FIXED_LIT);
          dt = buildTree(FIXED_DIST);
        } else {
          const hl = s.read(5) + 257;
          const hd = s.read(5) + 1;
          const hc = s.read(4) + 4;
          if (hl < 257) {
            corrupted = true;
            break;
          }
          const cl = new Uint8Array(19);
          for (let i = 0; i < hc; i++) cl[CLEN_ORDER[i]] = s.read(3);
          const ct = buildTree(cl);
          if (!ct) {
            corrupted = true;
            break;
          }
          const unpack = (count: number): Uint8Array | null => {
            const r: number[] = [];
            while (r.length < count) {
              const sy = decodeSym(s, ct);
              if (sy < 0 || sy > 18) return null;
              if (sy < 16) r.push(sy);
              else if (sy === 16) {
                let c = 3 + s.read(2);
                const p = r[r.length - 1];
                while (c--) r.push(p);
              } else if (sy === 17) {
                let z = 3 + s.read(3);
                while (z--) r.push(0);
              } else if (sy === 18) {
                let z = 11 + s.read(7);
                while (z--) r.push(0);
              }
            }
            return new Uint8Array(r);
          };
          const ll = unpack(hl);
          const dl = unpack(hd);
          if (!ll || !dl) {
            corrupted = true;
            break;
          }
          lt = buildTree(ll);
          dt = buildTree(dl);
        }
        if (!lt || !dt) {
          corrupted = true;
          break;
        }
        while (true) {
          const sym = decodeSym(s, lt);
          if (sym === -1 || sym === -2) {
            corrupted = true;
            break;
          }
          if (sym === 256) break;
          if (sym < 256) out.push(sym);
          else {
            const lc = sym - 257;
            if (lc > 28) {
              corrupted = true;
              break;
            }
            const length = LEN_BASE[lc] + s.read(LEN_EXTRA[lc]);
            const dc = decodeSym(s, dt);
            if (dc < 0) {
              corrupted = true;
              break;
            }
            const dist = DIST_BASE[dc] + s.read(DIST_EXTRA[dc]);
            if (dist > out.length) {
              corrupted = true;
              bfinal = 1;
              break;
            }
            let ptr = out.length - dist;
            for (let i = 0; i < length; i++) out.push(out[ptr++]);
          }
        }
      } else {
        corrupted = true;
        break;
      }
    }
  } catch {
    corrupted = true;
  }
  return { data: new Uint8Array(out), isCorrupt: corrupted };
}
