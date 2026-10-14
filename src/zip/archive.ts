import { unzipSync, zipSync, type Zippable, type ZipOptions } from 'fflate';
import { ZipSecurityError } from './errors';

export interface ZipEntry {
  /** Normalized entry path inside the archive (forward slashes). */
  name: string;
  /** Decompressed entry content. */
  data: Uint8Array;
  /** Last modification time, when known. */
  mtime?: Date;
}

export interface ReadLimits {
  /** Maximum number of entries allowed in the archive. */
  maxEntries: number;
  /** Maximum total decompressed size across all entries (bytes). */
  maxTotalUncompressed: number;
  /** Maximum decompressed size for a single entry (bytes). */
  maxEntryUncompressed: number;
  /** Maximum tolerated overall compression ratio (decompressed / compressed). */
  maxCompressionRatio: number;
}

export const DEFAULT_LIMITS: ReadLimits = {
  maxEntries: 10_000,
  maxTotalUncompressed: 512 * 1024 * 1024, // 512 MiB
  maxEntryUncompressed: 256 * 1024 * 1024, // 256 MiB
  maxCompressionRatio: 1000,
};

export interface ZipArchive {
  /** Entries in their original archive order. */
  entries: ZipEntry[];
  /** Fast lookup by entry name. */
  byName: Map<string, ZipEntry>;
  /** The original bytes, kept for zero-change pass-through writes. */
  raw: Uint8Array;
}

const UNSAFE = /(^\/)|(^[a-zA-Z]:)|(^|\/)\.\.(\/|$)|(\0)/;

function normalizeName(name: string): string {
  return name.replace(/\\/g, '/');
}

export function isSafeEntryName(name: string): boolean {
  if (name.length === 0) return false;
  const normalized = normalizeName(name);
  if (normalized.endsWith('/')) return false; // directory markers, not needed
  return !UNSAFE.test(normalized);
}

function toUint8(input: Uint8Array | ArrayBuffer | ArrayBufferView): Uint8Array {
  if (input instanceof Uint8Array) return input;
  if (ArrayBuffer.isView(input)) {
    return new Uint8Array(input.buffer, input.byteOffset, input.byteLength);
  }
  return new Uint8Array(input);
}

/**
 * Read a ZIP (OPC package) into memory with defense-in-depth limits.
 * Rejects path traversal, excessive entry counts and decompression bombs.
 */
export function readZip(
  input: Uint8Array | ArrayBuffer | ArrayBufferView,
  limits: Partial<ReadLimits> = {},
): ZipArchive {
  const raw = toUint8(input);
  const lim: ReadLimits = { ...DEFAULT_LIMITS, ...limits };

  const rejected: string[] = [];
  let rawMap: Record<string, Uint8Array>;
  try {
    rawMap = unzipSync(raw, {
      filter: (file) => {
        if (file.name.endsWith('/')) return false;
        if (!isSafeEntryName(file.name)) {
          rejected.push(file.name);
          return false;
        }
        // Reject oversized entries BEFORE inflating them (anti zip-bomb).
        if (file.originalSize > lim.maxEntryUncompressed) {
          rejected.push(file.name);
          return false;
        }
        return true;
      },
    });
  } catch (error) {
    throw new ZipSecurityError(`Unable to read ZIP archive: ${(error as Error).message}`);
  }
  if (rejected.length > 0) {
    throw new ZipSecurityError(`Unsafe entry rejected: ${JSON.stringify(rejected[0])}`);
  }

  const names = Object.keys(rawMap);
  if (names.length > lim.maxEntries) {
    throw new ZipSecurityError(
      `Archive has too many entries (${names.length} > ${lim.maxEntries})`,
    );
  }

  const entries: ZipEntry[] = [];
  const byName = new Map<string, ZipEntry>();
  let total = 0;

  for (const originalName of names) {
    if (!isSafeEntryName(originalName)) {
      throw new ZipSecurityError(`Unsafe entry name rejected: ${JSON.stringify(originalName)}`);
    }
    const name = normalizeName(originalName);
    const data = rawMap[originalName]!;
    if (data.length > lim.maxEntryUncompressed) {
      throw new ZipSecurityError(
        `Entry exceeds size limit (${name}: ${data.length} > ${lim.maxEntryUncompressed})`,
      );
    }
    total += data.length;
    if (total > lim.maxTotalUncompressed) {
      throw new ZipSecurityError(
        `Total decompressed size exceeds limit (${total} > ${lim.maxTotalUncompressed})`,
      );
    }
    const entry: ZipEntry = { name, data };
    entries.push(entry);
    byName.set(name, entry);
  }

  // Whole-archive decompression-bomb guard.
  if (raw.length > 0 && total / raw.length > lim.maxCompressionRatio) {
    throw new ZipSecurityError(
      `Suspicious compression ratio (${Math.round(total / raw.length)}x > ${lim.maxCompressionRatio}x)`,
    );
  }

  return { entries, byName, raw };
}

export interface WriteOptions {
  /** Compression level 0-9 (0 = stored). Defaults to 6. */
  level?: 0 | 1 | 2 | 3 | 4 | 5 | 6 | 7 | 8 | 9;
}

/**
 * Serialize archive entries back into a ZIP. Entry order is preserved.
 * (v0.2 will reuse untouched entries' raw deflate stream to be faster and
 * to guarantee whole-file byte identity.)
 */
export function writeZip(archive: ZipArchive, options: WriteOptions = {}): Uint8Array {
  const level = options.level ?? 6;
  const zippable: Zippable = {};
  for (const entry of archive.entries) {
    const entryOptions: ZipOptions = { level };
    if (entry.mtime) entryOptions.mtime = entry.mtime;
    zippable[entry.name] = [entry.data, entryOptions];
  }
  return zipSync(zippable, { level });
}
