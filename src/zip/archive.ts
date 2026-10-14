import {
  parseRawZip,
  serializeRawZip,
  isSafeEntryName,
  DEFAULT_LIMITS,
  type RawEntrySource,
  type RawWriteEntry,
  type ReadLimits,
} from './raw';

// Re-exported for backward compatibility: `archive` was the only ZIP module
// before the raw layer was introduced (the public API exports these from here).
export { isSafeEntryName, DEFAULT_LIMITS };
export type { ReadLimits, RawEntrySource };

export interface ZipEntry {
  /** Normalized entry path inside the archive (forward slashes). */
  name: string;
  /** Decompressed entry content. */
  data: Uint8Array;
  /** Last modification time, when known. */
  mtime?: Date;
  /**
   * Original raw representation (compressed bytes + CRC + headers), when the
   * entry was read from a file. While `data` still is the original buffer,
   * `writeZip` reuses the original compressed bytes verbatim instead of
   * re-compressing the entry.
   */
  source?: RawEntrySource;
}

export interface ZipArchive {
  /** Entries in their original archive order. */
  entries: ZipEntry[];
  /** Fast lookup by entry name. */
  byName: Map<string, ZipEntry>;
  /** The original bytes, kept for zero-change pass-through writes. */
  raw: Uint8Array;
  /** Names explicitly marked as changed (dirty); these are re-compressed. */
  dirtyNames?: ReadonlySet<string>;
  /** Bytes before the first local header, preserved on write. */
  prefixBytes?: Uint8Array;
  /** EOCD comment, preserved on write. */
  commentBytes?: Uint8Array;
}

/**
 * Read a ZIP (OPC package) into memory with defense-in-depth limits.
 * Rejects path traversal, excessive entry counts and decompression bombs —
 * all enforced on the ZIP structure before/while inflating (see `zip/raw.ts`).
 */
export function readZip(
  input: Uint8Array | ArrayBuffer | ArrayBufferView,
  limits: Partial<ReadLimits> = {},
): ZipArchive {
  const parsed = parseRawZip(input, limits);

  const entries: ZipEntry[] = parsed.entries.map((entry) => ({
    name: entry.name,
    data: entry.data,
    source: entry.source,
  }));

  return {
    entries,
    byName: new Map(entries.map((entry) => [entry.name, entry])),
    raw: parsed.raw,
    prefixBytes: parsed.prefix,
    commentBytes: parsed.comment,
  };
}

export interface WriteOptions {
  /** Compression level 0-9 (0 = stored). Defaults to 6. */
  level?: 0 | 1 | 2 | 3 | 4 | 5 | 6 | 7 | 8 | 9;
}

/**
 * Serialize archive entries back into a ZIP. Entry order is preserved.
 * Untouched entries are copied byte-for-byte (original compressed bytes +
 * CRC + headers); only changed or new entries are re-compressed.
 */
export function writeZip(archive: ZipArchive, options: WriteOptions = {}): Uint8Array {
  const dirty = archive.dirtyNames;
  const entries: RawWriteEntry[] = archive.entries.map((entry) => {
    const source = entry.source;
    if (source && entry.data === source.data && !(dirty && dirty.has(entry.name))) {
      // Untouched: reuse the original compressed bytes + CRC verbatim.
      return { name: entry.name, data: source.data, source };
    }
    return { name: entry.name, data: entry.data, mtime: entry.mtime };
  });
  return serializeRawZip(entries, {
    level: options.level ?? 6,
    prefix: archive.prefixBytes,
    comment: archive.commentBytes,
  });
}
