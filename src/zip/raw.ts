/**
 * Raw ZIP container layer.
 *
 * Hand-rolled parsing of local file headers, the central directory and the
 * end-of-central-directory record (EOCD), so that a write can reuse each
 * untouched entry's ORIGINAL compressed bytes + CRC verbatim instead of
 * re-compressing the whole archive. All container logic (headers, CRC32,
 * DOS timestamps) lives here and is dependency-free; only the DEFLATE codec
 * itself comes from fflate (which this package already depends on).
 *
 * Supported on read and write: method 0 (stored) and method 8 (deflate), and
 * data descriptors (general purpose flag bit 3 — the central directory carries
 * the real CRC/sizes, and the descriptor is dropped when re-serializing).
 * Deliberately rejected with a clear `SheetLiteError` code instead of
 * crashing: zip64, encryption (bit 0) and multi-disk archives.
 */
import { deflateSync, inflateSync } from 'fflate';
import { NotAZipError, SheetLiteError, ZipSecurityError, ZipUnsupportedError } from './errors';
import { isSafeEntryName, normalizeName } from './names';

export { isSafeEntryName, normalizeName };

/** Maximum number of entries allowed in the archive. */
export const DEFAULT_MAX_ENTRIES = 10_000;
/** Maximum total decompressed size across all entries (bytes). */
export const DEFAULT_MAX_TOTAL_UNCOMPRESSED = 512 * 1024 * 1024; // 512 MiB
/** Maximum decompressed size for a single entry (bytes). */
export const DEFAULT_MAX_ENTRY_UNCOMPRESSED = 256 * 1024 * 1024; // 256 MiB
/** Maximum tolerated overall compression ratio (decompressed / compressed). */
export const DEFAULT_MAX_COMPRESSION_RATIO = 1000;

export interface ReadLimits {
  maxEntries: number;
  maxTotalUncompressed: number;
  maxEntryUncompressed: number;
  maxCompressionRatio: number;
}

export const DEFAULT_LIMITS: ReadLimits = {
  maxEntries: DEFAULT_MAX_ENTRIES,
  maxTotalUncompressed: DEFAULT_MAX_TOTAL_UNCOMPRESSED,
  maxEntryUncompressed: DEFAULT_MAX_ENTRY_UNCOMPRESSED,
  maxCompressionRatio: DEFAULT_MAX_COMPRESSION_RATIO,
};

// Signatures ("PK\x03\x04", "PK\x01\x02", "PK\x05\x06", "PK\x06\x07").
const LOCAL_SIG = 0x0403_4b50;
const CENTRAL_SIG = 0x0201_4b50;
const EOCD_SIG = 0x0605_4b50;
const ZIP64_LOCATOR_SIG = 0x0706_4b50;

const FLAG_ENCRYPTED = 0x0001;
const FLAG_DATA_DESCRIPTOR = 0x0008;

/**
 * OLE2 / CFB compound-file magic (the OLE2 signature is stored little-endian).
 * Legacy `.xls` workbooks and password-protected/encrypted OOXML packages are
 * this container, not a ZIP.
 */
const OLE2_MAGIC = [0xd0, 0xcf, 0x11, 0xe0, 0xa1, 0xb1, 0x1a, 0xe1];

/**
 * Reject a legacy OLE2/CFB container with an honest "wrong format" error
 * instead of letting it die inside the ZIP parser as a corrupt (and, by
 * implication, suspicious) archive. Only the first 8 bytes are inspected, and
 * `bytes` may be a short prefix of a file, so this is safe to call with any
 * leading chunk.
 */
export function assertNotCompoundFile(bytes: Uint8Array): void {
  if (bytes.length < OLE2_MAGIC.length) return;
  for (let i = 0; i < OLE2_MAGIC.length; i += 1) {
    if (bytes[i] !== OLE2_MAGIC[i]) return;
  }
  throw new SheetLiteError(
    'Not an XLSX package: OLE2 compound file (legacy .xls or encrypted OOXML); sheet-lite reads .xlsx/.xlsm only',
    'NOT_XLSX',
  );
}

/** Fixed part sizes, in bytes. */
export const LOCAL_FIXED = 30; // local file header
export const CENTRAL_FIXED = 46; // central directory record, fixed part
const EOCD_FIXED = 22;

const EMPTY = new Uint8Array(0);

function toUint8(input: Uint8Array | ArrayBuffer | ArrayBufferView): Uint8Array {
  if (input instanceof Uint8Array) return input;
  if (ArrayBuffer.isView(input)) {
    return new Uint8Array(input.buffer, input.byteOffset, input.byteLength);
  }
  return new Uint8Array(input);
}

// ---------------------------------------------------------------------------
// CRC32 (IEEE 802.3, reflected, as used by ZIP)
// ---------------------------------------------------------------------------

const CRC_TABLE = /* @__PURE__ */ (() => {
  const t0 = new Uint32Array(256);
  for (let i = 0; i < 256; i += 1) {
    let c = i;
    for (let k = 0; k < 8; k += 1) {
      c = c & 1 ? 0xedb8_8320 ^ (c >>> 1) : c >>> 1;
    }
    t0[i] = c >>> 0;
  }
  // Slicing-by-8 tables, flattened: slot k*256 + x, where
  // Tk[i] = (Tk-1[i] >>> 8) ^ T0[Tk-1[i] & 0xff].
  const flat = new Uint32Array(8 * 256);
  flat.set(t0, 0);
  for (let k = 1; k < 8; k += 1) {
    const base = k * 256;
    const previous = k - 1 === 0 ? t0 : flat.subarray((k - 1) * 256, k * 256);
    for (let i = 0; i < 256; i += 1) {
      const v = previous[i]!;
      flat[base + i] = (v >>> 8) ^ t0[v & 0xff]!;
    }
  }
  return flat;
})();

/**
 * CRC-32 checksum of a buffer (the checksum ZIP stores per entry).
 *
 * Streams can compute it incrementally: `crc32(b, crc32(a))` equals
 * `crc32(concat(a, b))` — pass the previous result as the `previous` argument
 * (0 for the first chunk).
 */
export function crc32(data: Uint8Array, previous = 0): number {
  const t = CRC_TABLE;
  let c = (previous ^ 0xffff_ffff) >>> 0;
  const len = data.length;
  let i = 0;
  while (i + 8 <= len) {
    const one = (c ^ (data[i]! | (data[i + 1]! << 8) | (data[i + 2]! << 16) | (data[i + 3]! << 24))) >>> 0;
    const two = (data[i + 4]! | (data[i + 5]! << 8) | (data[i + 6]! << 16) | (data[i + 7]! << 24)) >>> 0;
    c = t[7 * 256 + (one & 0xff)]! ^
      t[6 * 256 + ((one >>> 8) & 0xff)]! ^
      t[5 * 256 + ((one >>> 16) & 0xff)]! ^
      t[4 * 256 + ((one >>> 24) & 0xff)]! ^
      t[3 * 256 + (two & 0xff)]! ^
      t[2 * 256 + ((two >>> 8) & 0xff)]! ^
      t[1 * 256 + ((two >>> 16) & 0xff)]! ^
      t[(two >>> 24) & 0xff]!;
    i += 8;
  }
  while (i < len) {
    c = t[(c ^ data[i]!) & 0xff]! ^ (c >>> 8);
    i += 1;
  }
  return (c ^ 0xffff_ffff) >>> 0;
}

// ---------------------------------------------------------------------------
// Parsing
// ---------------------------------------------------------------------------

/** The original, raw representation of an entry, kept for verbatim reuse. */
export interface RawEntrySource {
  /** Decompressed content (the same buffer the entry's `data` points at). */
  data: Uint8Array;
  /** Compressed bytes, as a view into the original file buffer. */
  compressed: Uint8Array;
  /** CRC-32 of the decompressed content, as stored in the archive. */
  crc32: number;
  /** Compression method: 0 (stored) or 8 (deflate). */
  method: 0 | 8;
  /** Local header + name + extra field + compressed data, byte-for-byte. */
  localBytes: Uint8Array;
  /** Full central-directory record, byte-for-byte. */
  centralBytes: Uint8Array;
}

export interface RawParsedEntry {
  name: string;
  data: Uint8Array;
  source: RawEntrySource;
}

export interface RawZipFile {
  /** Entries that passed validation, in original archive order. */
  entries: RawParsedEntry[];
  /** Bytes before the first local header (e.g. a self-extractor prefix). */
  prefix: Uint8Array;
  /** EOCD comment bytes, verbatim. */
  comment: Uint8Array;
  /** The original file buffer. */
  raw: Uint8Array;
}

/**
 * Locate and validate the end-of-central-directory record.
 * Scans backwards over the maximum possible ZIP comment (64 KiB).
 */
function findEocd(raw: Uint8Array, dv: DataView): number {
  const minPos = Math.max(0, raw.length - EOCD_FIXED - 0xffff);
  for (let pos = raw.length - EOCD_FIXED; pos >= minPos; pos -= 1) {
    if (dv.getUint32(pos, true) !== EOCD_SIG) continue;
    const commentLen = dv.getUint16(pos + 20, true);
    if (pos + EOCD_FIXED + commentLen === raw.length) return pos;
  }
  throw new NotAZipError('Not a ZIP archive: end of central directory record not found');
}

/** End-of-central-directory information, validated (multi-disk / zip64 already rejected). */
export interface EocdInfo {
  /** Index of the EOCD record within the provided buffer. */
  eocdPos: number;
  /** Absolute offset of the central directory in the archive. */
  centralDirOffset: number;
  centralDirSize: number;
  totalEntries: number;
  comment: Uint8Array;
}

/**
 * Read and validate the end-of-central-directory record. `raw` must end at the
 * end of the archive; `base` is the absolute archive offset of `raw[0]` (0 for a
 * whole-file buffer, `size - raw.length` when only the tail was read).
 *
 * Rejects multi-disk archives and every zip64 marker with a `ZipUnsupportedError`
 * before any entry is touched. Shared by the eager (`parseRawZip`) and the
 * streaming (`zip/stream.ts`) readers so the two can never disagree.
 */
export function readEocd(raw: Uint8Array, base = 0): EocdInfo {
  const dv = new DataView(raw.buffer, raw.byteOffset, raw.byteLength);
  const eocdPos = findEocd(raw, dv);

  const diskNumber = dv.getUint16(eocdPos + 4, true);
  const diskWithCentralDir = dv.getUint16(eocdPos + 6, true);
  const entriesOnDisk = dv.getUint16(eocdPos + 8, true);
  const totalEntries = dv.getUint16(eocdPos + 10, true);
  const centralDirSize = dv.getUint32(eocdPos + 12, true);
  const centralDirOffset = dv.getUint32(eocdPos + 16, true);
  const comment = raw.subarray(eocdPos + EOCD_FIXED, raw.length);

  if (diskNumber !== 0 || diskWithCentralDir !== 0) {
    throw new ZipUnsupportedError('ZIP feature not supported: multi-disk archives are not supported');
  }
  if (
    entriesOnDisk === 0xffff ||
    totalEntries === 0xffff ||
    centralDirSize === 0xffff_ffff ||
    centralDirOffset === 0xffff_ffff
  ) {
    throw new ZipUnsupportedError('ZIP feature not supported: zip64 archives are not supported (end-of-central-directory fields exceed the 32-bit limits)');
  }
  // The zip64 EOCD locator (when present) sits 20 bytes before the EOCD.
  if (eocdPos >= 20 && dv.getUint32(eocdPos - 20, true) === ZIP64_LOCATOR_SIG) {
    throw new ZipUnsupportedError('ZIP feature not supported: zip64 archives are not supported (zip64 end-of-central-directory locator found)');
  }
  const centralDirRel = centralDirOffset - base;
  if (centralDirRel < 0 || centralDirRel + centralDirSize > eocdPos) {
    throw new ZipSecurityError('Unable to read ZIP archive: central directory out of bounds');
  }

  return { eocdPos, centralDirOffset, centralDirSize, totalEntries, comment };
}

/**
 * A validated central-directory entry: metadata only, no inflated bytes.
 * This is the common currency between the eager reader (which additionally
 * attaches the raw byte ranges for verbatim reuse) and the streaming reader.
 */
export interface ZipDirEntry {
  name: string;
  method: 0 | 8;
  dosTime: number;
  dosDate: number;
  crc32: number;
  compressedSize: number;
  uncompressedSize: number;
  /** Absolute offset of the entry's local file header. */
  localOffset: number;
  /** Byte range of the central-directory record within the central-directory buffer. */
  centralStart: number;
  centralEnd: number;
  /** True when the local header uses a trailing data descriptor (flag bit 3). */
  dataDescriptor?: boolean;
  // Populated by the eager path only (verbatim reuse + bounds checks).
  compressed?: Uint8Array;
  localBytes?: Uint8Array;
  centralBytes?: Uint8Array;
}

/**
 * Parse and validate central-directory records from `cd` (the raw bytes of the
 * central directory). Enforces the same structural, traversal, entry-count,
 * per-entry-size and total-size limits as the eager reader BEFORE any inflation.
 * Names that fail the traversal check are collected in `rejected`; the callers
 * turn a non-empty list into a `ZipSecurityError`.
 */
export function parseCentralDirectory(
  cd: Uint8Array,
  totalEntries: number,
  lim: ReadLimits,
): { entries: ZipDirEntry[]; rejected: string[] } {
  const dv = new DataView(cd.buffer, cd.byteOffset, cd.byteLength);
  const decoder = new TextDecoder('utf-8');
  const entries: ZipDirEntry[] = [];
  const rejected: string[] = [];
  let recPos = 0;
  let declaredTotal = 0;

  for (let index = 0; index < totalEntries; index += 1) {
    if (recPos + CENTRAL_FIXED > cd.length || dv.getUint32(recPos, true) !== CENTRAL_SIG) {
      throw new ZipSecurityError('Unable to read ZIP archive: corrupt central directory');
    }
    const recordStart = recPos;
    const flags = dv.getUint16(recPos + 8, true);
    const method = dv.getUint16(recPos + 10, true);
    const dosTime = dv.getUint16(recPos + 12, true);
    const dosDate = dv.getUint16(recPos + 14, true);
    const crc = dv.getUint32(recPos + 16, true);
    const compressedSize = dv.getUint32(recPos + 20, true);
    const uncompressedSize = dv.getUint32(recPos + 24, true);
    const nameLen = dv.getUint16(recPos + 28, true);
    const extraLen = dv.getUint16(recPos + 30, true);
    const commentLen = dv.getUint16(recPos + 32, true);
    const diskStart = dv.getUint16(recPos + 34, true);
    const localOffset = dv.getUint32(recPos + 42, true);
    recPos += CENTRAL_FIXED + nameLen + extraLen + commentLen;

    if (recPos > cd.length) {
      throw new ZipSecurityError('Unable to read ZIP archive: central directory record out of bounds');
    }
    const name = decoder.decode(cd.subarray(recordStart + CENTRAL_FIXED, recordStart + CENTRAL_FIXED + nameLen));

    if (name.endsWith('/')) continue; // directory marker: not a part, skipped
    if (diskStart === 0xffff || compressedSize === 0xffff_ffff || uncompressedSize === 0xffff_ffff || localOffset === 0xffff_ffff) {
      throw new ZipUnsupportedError(`ZIP feature not supported: zip64 archives are not supported (entry ${JSON.stringify(name)} uses 64-bit fields)`);
    }
    if (flags & FLAG_ENCRYPTED) {
      throw new ZipUnsupportedError(`ZIP feature not supported: entry ${JSON.stringify(name)} is encrypted`);
    }
    // Data descriptors (general purpose flag bit 3) are supported: the central
    // directory always carries the real CRC and sizes, and the descriptor bytes
    // are dropped when the entry is re-serialized (its local header is normalized
    // in parseStructure). Streaming ZIP writers emit them routinely.
    const dataDescriptor = (flags & FLAG_DATA_DESCRIPTOR) !== 0;
    if (method !== 0 && method !== 8) {
      throw new ZipUnsupportedError(`ZIP feature not supported: entry ${JSON.stringify(name)} uses compression method ${method} (only stored and deflate are supported)`);
    }
    if (!isSafeEntryName(name)) {
      rejected.push(name);
      continue;
    }
    // Anti zip-bomb: enforce the per-entry limit on the declared size BEFORE inflating.
    if (uncompressedSize > lim.maxEntryUncompressed) {
      rejected.push(name);
      continue;
    }
    declaredTotal += uncompressedSize;
    if (declaredTotal > lim.maxTotalUncompressed) {
      throw new ZipSecurityError(`Total decompressed size exceeds limit (${declaredTotal} > ${lim.maxTotalUncompressed})`);
    }
    if (entries.length >= lim.maxEntries) {
      throw new ZipSecurityError(`Archive has too many entries (> ${lim.maxEntries})`);
    }

    entries.push({
      name: normalizeName(name),
      method: method as 0 | 8,
      dosTime,
      dosDate,
      crc32: crc,
      compressedSize,
      uncompressedSize,
      localOffset,
      centralStart: recordStart,
      centralEnd: recPos,
      dataDescriptor,
    });
  }

  return { entries, rejected };
}

/**
 * Validate an entry's local file header (the 30 fixed bytes read at
 * `localOffset`) and return the absolute offset of its compressed data.
 * Shared by the eager and streaming readers.
 */
export function resolveEntryDataOffset(
  localHeader: Uint8Array,
  localOffset: number,
  archiveSize: number,
  name: string,
  compressedSize: number,
): number {
  if (localHeader.length < LOCAL_FIXED) {
    throw new ZipSecurityError(`Unable to read ZIP archive: corrupt local header for ${JSON.stringify(name)}`);
  }
  const dv = new DataView(localHeader.buffer, localHeader.byteOffset, localHeader.byteLength);
  if (dv.getUint32(0, true) !== LOCAL_SIG) {
    throw new ZipSecurityError(`Unable to read ZIP archive: corrupt local header for ${JSON.stringify(name)}`);
  }
  const localNameLen = dv.getUint16(26, true);
  const localExtraLen = dv.getUint16(28, true);
  const dataStart = localOffset + LOCAL_FIXED + localNameLen + localExtraLen;
  const dataEnd = dataStart + compressedSize;
  if (dataStart > archiveSize || dataEnd > archiveSize) {
    throw new ZipSecurityError(`Unable to read ZIP archive: data of ${JSON.stringify(name)} out of bounds`);
  }
  return dataStart;
}

/**
 * Parse a ZIP file's structure without inflating anything yet:
 * EOCD -> central directory records -> local header bounds.
 * Enforces entry-count and per-entry/total size limits BEFORE decompression.
 */
function parseStructure(
  raw: Uint8Array,
  lim: ReadLimits,
): { records: ZipDirEntry[]; prefix: Uint8Array; comment: Uint8Array } {
  const { centralDirOffset, centralDirSize, totalEntries, comment } = readEocd(raw, 0);
  const cd = raw.subarray(centralDirOffset, centralDirOffset + centralDirSize);
  const { entries, rejected } = parseCentralDirectory(cd, totalEntries, lim);

  if (rejected.length > 0) {
    throw new ZipSecurityError(`Unsafe entry rejected: ${JSON.stringify(rejected[0])}`);
  }

  // Locate each entry's data via its local file header (still no inflation).
  let minLocalOffset = raw.length;
  for (const entry of entries) {
    const at = entry.localOffset;
    const dataStart = resolveEntryDataOffset(
      raw.subarray(at, at + LOCAL_FIXED),
      at,
      raw.length,
      entry.name,
      entry.compressedSize,
    );
    const dataEnd = dataStart + entry.compressedSize;
    entry.compressed = raw.subarray(dataStart, dataEnd);
    if (entry.dataDescriptor) {
      // Normalize a bit-3 entry: drop the trailing data descriptor and write the
      // real CRC/sizes into the local header, clearing flag bit 3 so the rebuilt
      // archive is self-consistent (readers that trust the local header work).
      const local = raw.slice(at, dataEnd);
      const ldv = new DataView(local.buffer, local.byteOffset, local.byteLength);
      ldv.setUint16(6, ldv.getUint16(6, true) & ~FLAG_DATA_DESCRIPTOR, true);
      ldv.setUint32(14, entry.crc32, true);
      ldv.setUint32(18, entry.compressedSize, true);
      ldv.setUint32(22, entry.uncompressedSize, true);
      entry.localBytes = local;

      const central = cd.slice(entry.centralStart, entry.centralEnd);
      const cdv = new DataView(central.buffer, central.byteOffset, central.byteLength);
      cdv.setUint16(8, cdv.getUint16(8, true) & ~FLAG_DATA_DESCRIPTOR, true);
      entry.centralBytes = central;
    } else {
      entry.localBytes = raw.subarray(at, dataEnd);
      entry.centralBytes = cd.subarray(entry.centralStart, entry.centralEnd);
    }
    if (at < minLocalOffset) minLocalOffset = at;
  }

  return { records: entries, prefix: raw.subarray(0, minLocalOffset), comment };
}

/**
 * Parse a ZIP with the full hardening limits, keeping each entry's original
 * compressed bytes around for verbatim reuse on write. Inflation happens
 * only after every structural and limit check has passed (the pre-inflation
 * filter), and each entry is verified against its stored CRC-32.
 */
export function parseRawZip(
  input: Uint8Array | ArrayBuffer | ArrayBufferView,
  limits: Partial<ReadLimits> = {},
): RawZipFile {
  const raw = toUint8(input);
  if (raw.length < EOCD_FIXED) {
    throw new NotAZipError('Not a ZIP archive: file is too small to be a ZIP');
  }
  const lim: ReadLimits = { ...DEFAULT_LIMITS, ...limits };
  const { records, prefix, comment } = parseStructure(raw, lim);

  let total = 0;
  const entries: RawParsedEntry[] = [];
  for (const record of records) {
    // All limits were already enforced on the declared sizes; inflate and
    // verify integrity (CRC-32) before the entry is accepted. The output
    // buffer is preallocated to the DECLARED size, so even an archive with
    // lying headers can never allocate more memory than was validated.
    const compressed = record.compressed!;
    let data: Uint8Array;
    if (record.method === 0) {
      data = compressed.slice();
    } else if (record.uncompressedSize === 0) {
      data = EMPTY;
    } else {
      const out = new Uint8Array(record.uncompressedSize);
      try {
        data = inflateSync(compressed, { out });
      } catch (error) {
        throw new ZipSecurityError(`Unable to read ZIP archive: cannot inflate ${JSON.stringify(record.name)}: ${(error as Error).message}`);
      }
    }
    if (data.length !== record.uncompressedSize) {
      throw new ZipSecurityError(`Unable to read ZIP archive: size mismatch in ${JSON.stringify(record.name)} (${data.length} != ${record.uncompressedSize})`);
    }
    if (crc32(data) !== record.crc32) {
      throw new ZipSecurityError(`Unable to read ZIP archive: CRC mismatch in ${JSON.stringify(record.name)}`);
    }
    if (data.length > lim.maxEntryUncompressed) {
      throw new ZipSecurityError(`Entry exceeds size limit (${record.name}: ${data.length} > ${lim.maxEntryUncompressed})`);
    }
    total += data.length;
    if (total > lim.maxTotalUncompressed) {
      throw new ZipSecurityError(`Total decompressed size exceeds limit (${total} > ${lim.maxTotalUncompressed})`);
    }
    entries.push({
      name: record.name,
      data,
      source: {
        data,
        compressed,
        crc32: record.crc32,
        method: record.method,
        localBytes: record.localBytes!,
        centralBytes: record.centralBytes!,
      },
    });
  }

  // Whole-archive decompression-bomb guard.
  if (raw.length > 0 && total / raw.length > lim.maxCompressionRatio) {
    throw new ZipSecurityError(
      `Suspicious compression ratio (${Math.round(total / raw.length)}x > ${lim.maxCompressionRatio}x)`,
    );
  }

  return { entries, prefix, comment, raw };
}

// ---------------------------------------------------------------------------
// Serialization
// ---------------------------------------------------------------------------

export interface RawWriteEntry {
  name: string;
  /** Decompressed content. */
  data: Uint8Array;
  /** Last modification time (fresh entries only; verbatim reuse ignores it). */
  mtime?: Date;
  /** Original raw entry: set for untouched entries read from a file. */
  source?: RawEntrySource;
}

export interface RawWriteOptions {
  /** Compression level 0-9 for re-serialized entries (0 = stored). */
  level?: 0 | 1 | 2 | 3 | 4 | 5 | 6 | 7 | 8 | 9;
  /** Bytes before the first local header, preserved from the source file. */
  prefix?: Uint8Array;
  /** EOCD comment, preserved from the source file. */
  comment?: Uint8Array;
}

/** Encode a Date as ZIP (MS-DOS) local time. Clamped to the 1980-2107 range. */
function dosDateTime(date: Date): { time: number; date: number } {
  const year = Math.min(2107, Math.max(1980, date.getFullYear()));
  return {
    time: (date.getHours() << 11) | (date.getMinutes() << 5) | (date.getSeconds() >> 1),
    date: ((year - 1980) << 9) | ((date.getMonth() + 1) << 5) | date.getDate(),
  };
}

/** Default timestamp for fresh entries without an mtime (1980-01-01 00:00). */
const DOS_EPOCH = { time: 0, date: (0 << 9) | (1 << 5) | 1 };

const encoder = new TextEncoder();

interface FreshEntry {
  kind: 'fresh';
  name: string;
  nameBytes: Uint8Array;
  payload: Uint8Array;
  uncompressedSize: number;
  method: 0 | 8;
  crc: number;
  time: number;
  date: number;
}

interface VerbatimEntry {
  kind: 'verbatim';
  source: RawEntrySource;
}

/**
 * Serialize entries back into a ZIP. Untouched entries (those whose
 * `source` is set and whose `data` buffer is still the original one) are
 * copied byte-for-byte — original local header, name, extra field, compressed
 * data, CRC and central-directory record (with only the relocated offset
 * patched) — so they remain byte-identical to the source archive.
 */
export function serializeRawZip(entries: RawWriteEntry[], options: RawWriteOptions = {}): Uint8Array {
  const level = options.level ?? 6;
  const prefix = options.prefix ?? EMPTY;
  const comment = options.comment ?? EMPTY;

  // Decide, per entry, whether it is reused verbatim or re-serialized.
  const out: Array<FreshEntry | VerbatimEntry> = [];
  for (const entry of entries) {
    const source = entry.source;
    if (source && entry.data === source.data) {
      out.push({ kind: 'verbatim', source });
      continue;
    }
    if (!isSafeEntryName(entry.name)) {
      throw new ZipSecurityError(`Unsafe entry name rejected: ${JSON.stringify(entry.name)}`);
    }
    const method = level === 0 ? 0 : 8;
    const payload = method === 0 ? entry.data : deflateSync(entry.data, { level });
    const { time, date } = entry.mtime ? dosDateTime(entry.mtime) : DOS_EPOCH;
    out.push({
      kind: 'fresh',
      name: entry.name,
      nameBytes: encoder.encode(entry.name),
      payload,
      uncompressedSize: entry.data.length,
      method,
      crc: crc32(entry.data),
      time,
      date,
    });
  }

  // Layout & 32-bit overflow guards (zip64 is out of scope by design).
  if (out.length > 0xffff) {
    throw new ZipUnsupportedError('ZIP feature not supported: more than 65535 entries would require zip64 (not supported)');
  }
  let pos = prefix.length;
  const localOffsets: number[] = [];
  for (const entry of out) {
    const size = entry.kind === 'verbatim'
      ? entry.source.localBytes.length
      : LOCAL_FIXED + entry.nameBytes.length + entry.payload.length;
    if (pos + size > 0xffff_fffe) {
      throw new ZipUnsupportedError('ZIP feature not supported: output exceeds the 32-bit ZIP size limit (zip64 is not supported)');
    }
    localOffsets.push(pos);
    pos += size;
  }
  const centralOffset = pos;
  for (const entry of out) {
    const size = entry.kind === 'verbatim' ? entry.source.centralBytes.length : CENTRAL_FIXED + entry.nameBytes.length;
    if (pos + size > 0xffff_fffe) {
      throw new ZipUnsupportedError('ZIP feature not supported: output exceeds the 32-bit ZIP size limit (zip64 is not supported)');
    }
    pos += size;
  }
  const centralSize = pos - centralOffset;
  const totalLength = pos + EOCD_FIXED + comment.length;

  const output = new Uint8Array(totalLength);
  const dv = new DataView(output.buffer);
  output.set(prefix, 0);

  // Local entries.
  out.forEach((entry, index) => {
    const at = localOffsets[index]!;
    if (entry.kind === 'verbatim') {
      output.set(entry.source.localBytes, at);
      return;
    }
    dv.setUint32(at, LOCAL_SIG, true);
    dv.setUint16(at + 4, entry.method === 0 ? 10 : 20, true); // version needed
    dv.setUint16(at + 6, 0, true); // general purpose flags
    dv.setUint16(at + 8, entry.method, true);
    dv.setUint16(at + 10, entry.time, true);
    dv.setUint16(at + 12, entry.date, true);
    dv.setUint32(at + 14, entry.crc, true);
    dv.setUint32(at + 18, entry.payload.length, true); // compressed size
    dv.setUint32(at + 22, entry.uncompressedSize, true);
    dv.setUint16(at + 26, entry.nameBytes.length, true);
    dv.setUint16(at + 28, 0, true); // no extra field
    output.set(entry.nameBytes, at + LOCAL_FIXED);
    output.set(entry.payload, at + LOCAL_FIXED + entry.nameBytes.length);
  });

  // Central directory.
  pos = centralOffset;
  out.forEach((entry, index) => {
    const at = pos;
    const localOffset = localOffsets[index]!;
    if (entry.kind === 'verbatim') {
      output.set(entry.source.centralBytes, at);
      dv.setUint32(at + 42, localOffset, true); // relocated local header offset
      pos += entry.source.centralBytes.length;
      return;
    }
    dv.setUint32(at, CENTRAL_SIG, true);
    dv.setUint16(at + 4, 0x0014, true); // version made by (MS-DOS, 2.0)
    dv.setUint16(at + 6, entry.method === 0 ? 10 : 20, true); // version needed
    dv.setUint16(at + 8, 0, true); // general purpose flags
    dv.setUint16(at + 10, entry.method, true);
    dv.setUint16(at + 12, entry.time, true);
    dv.setUint16(at + 14, entry.date, true);
    dv.setUint32(at + 16, entry.crc, true);
    dv.setUint32(at + 20, entry.payload.length, true); // compressed size
    dv.setUint32(at + 24, entry.uncompressedSize, true);
    dv.setUint16(at + 28, entry.nameBytes.length, true);
    dv.setUint16(at + 30, 0, true); // no extra field
    dv.setUint16(at + 32, 0, true); // no comment
    dv.setUint16(at + 34, 0, true); // disk start
    dv.setUint16(at + 36, 0, true); // internal attrs
    dv.setUint32(at + 38, 0, true); // external attrs
    dv.setUint32(at + 42, localOffset, true);
    output.set(entry.nameBytes, at + CENTRAL_FIXED);
    pos += CENTRAL_FIXED + entry.nameBytes.length;
  });

  // End of central directory.
  const eocdAt = centralOffset + centralSize;
  dv.setUint32(eocdAt, EOCD_SIG, true);
  dv.setUint16(eocdAt + 4, 0, true); // disk number
  dv.setUint16(eocdAt + 6, 0, true); // disk with central directory
  dv.setUint16(eocdAt + 8, out.length, true);
  dv.setUint16(eocdAt + 10, out.length, true);
  dv.setUint32(eocdAt + 12, centralSize, true);
  dv.setUint32(eocdAt + 16, centralOffset, true);
  dv.setUint16(eocdAt + 20, comment.length, true);
  output.set(comment, eocdAt + EOCD_FIXED);

  return output;
}
