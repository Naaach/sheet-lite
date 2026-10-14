/**
 * Streaming ZIP entry access.
 *
 * The eager reader (`zip/raw.ts`) inflates the whole archive into memory. This
 * module adds a low-memory counterpart built on the SAME structural parser
 * (`readEocd` / `parseCentralDirectory` / `resolveEntryDataOffset` in raw.ts),
 * so a streaming open enforces exactly the same hardening rules:
 *
 * - zip64, encryption and multi-disk archives are
 *   rejected with the same stable error codes,
 * - path-traversal names, entry count, per-entry size and total-size limits
 *   are enforced on the central directory BEFORE any inflation,
 * - the compression-ratio limit is enforced per entry as it is streamed,
 * - every streamed byte is verified against the entry's stored CRC-32, and a
 *   stream that produces more bytes than the header declares is aborted.
 *
 * Memory stays bounded because entries are read from a random-access source in
 * fixed-size chunks and inflated with fflate's push-based `Inflate`: only one
 * chunk of input and a few chunks of output are ever alive at a time.
 */
import { Inflate } from 'fflate';
import {
  crc32,
  readEocd,
  parseCentralDirectory,
  resolveEntryDataOffset,
  DEFAULT_LIMITS,
  LOCAL_FIXED,
  type ReadLimits,
  type ZipDirEntry,
} from './raw';
import { ZipSecurityError } from './errors';

/** Compressed bytes are pulled from the source in chunks of this size. */
const READ_CHUNK = 64 * 1024;

/** EOCD record + maximum ZIP comment length (the tail region an EOCD can live in). */
const EOCD_TAIL = 22 + 0xffff;

/**
 * Random-access byte source. `read` may return fewer bytes than requested only
 * at the end of the source (short reads elsewhere are treated as corruption).
 */
export interface ByteSource {
  readonly size: number;
  read(offset: number, length: number): Uint8Array | Promise<Uint8Array>;
}

/**
 * Minimal interface of a Node `FileHandle` (kept structural so this module
 * stays runtime-agnostic and free of static `node:` imports).
 */
export interface RandomAccessFile {
  read(buffer: Uint8Array, offset: number, length: number, position: number): Promise<{ bytesRead: number }>;
  stat(): Promise<{ size: number }>;
  close(): Promise<void>;
}

/** Random-access source over an open file handle (Node `fs.promises.open`). */
export class FileByteSource implements ByteSource {
  readonly size: number;

  private constructor(
    private readonly file: RandomAccessFile,
    size: number,
  ) {
    this.size = size;
  }

  /** Open a file for random access and return its size. Node-only (dynamic import by the caller). */
  static async open(path: string): Promise<FileByteSource> {
    const fs = await import('node:fs/promises');
    const handle = await fs.open(path, 'r');
    try {
      const size = (await handle.stat()).size;
      return new FileByteSource(handle, size);
    } catch (error) {
      await handle.close();
      throw error;
    }
  }

  async read(offset: number, length: number): Promise<Uint8Array> {
    const buffer = new Uint8Array(length);
    let total = 0;
    while (total < length) {
      const { bytesRead } = await this.file.read(buffer, total, length - total, offset + total);
      if (bytesRead <= 0) break;
      total += bytesRead;
    }
    return total === length ? buffer : buffer.subarray(0, total);
  }

  async close(): Promise<void> {
    await this.file.close();
  }
}

/** Random-access source over an in-memory buffer (small/medium files, tests, browser). */
export class BufferByteSource implements ByteSource {
  readonly size: number;

  constructor(private readonly bytes: Uint8Array) {
    this.size = bytes.length;
  }

  read(offset: number, length: number): Uint8Array {
    if (offset > this.size) return new Uint8Array(0);
    return this.bytes.subarray(offset, Math.min(this.size, offset + length));
  }
}

export interface ZipDirectory {
  /** Validated entries, in archive order. */
  entries: ZipDirEntry[];
  /** Fast lookup by normalized entry name. */
  byName: Map<string, ZipDirEntry>;
  /** EOCD comment bytes, verbatim. */
  comment: Uint8Array;
}

/**
 * Parse and validate a ZIP's structure with targeted reads (EOCD tail +
 * central directory) instead of loading the whole file. Everything the eager
 * reader validates before inflation is validated here too, so a streaming
 * `open()` rejects a malformed or hostile archive before reading any entry.
 */
export async function readZipDirectory(
  source: ByteSource,
  limits: Partial<ReadLimits> = {},
): Promise<ZipDirectory> {
  const lim: ReadLimits = { ...DEFAULT_LIMITS, ...limits };
  if (source.size < 22) {
    throw new ZipSecurityError('Unable to read ZIP archive: file is too small to be a ZIP');
  }
  const tailLen = Math.min(source.size, EOCD_TAIL);
  const tail = await source.read(source.size - tailLen, tailLen);
  const { centralDirOffset, centralDirSize, totalEntries, comment } = readEocd(tail, source.size - tail.length);

  const cd = await source.read(centralDirOffset, centralDirSize);
  if (cd.length < centralDirSize) {
    throw new ZipSecurityError('Unable to read ZIP archive: corrupt central directory');
  }
  const { entries, rejected } = parseCentralDirectory(cd, totalEntries, lim);
  if (rejected.length > 0) {
    throw new ZipSecurityError(`Unsafe entry rejected: ${JSON.stringify(rejected[0])}`);
  }
  return {
    entries,
    byName: new Map(entries.map((entry) => [entry.name, entry])),
    comment,
  };
}

/**
 * Incrementally decompress one entry, yielding chunks of its content.
 *
 * The entry's declared sizes were already validated by `readZipDirectory`;
 * here we additionally validate its local file header, enforce the
 * compression-ratio limit on the declared sizes, stream CRC-32 verification
 * and abort if the stream produces more bytes than the header declares (so
 * lying headers can never over-allocate).
 */
export async function* streamEntryChunks(
  source: ByteSource,
  entry: ZipDirEntry,
  limits: Partial<ReadLimits> = {},
): AsyncGenerator<Uint8Array> {
  const lim: ReadLimits = { ...DEFAULT_LIMITS, ...limits };

  const localHeader = await source.read(entry.localOffset, LOCAL_FIXED);
  const dataStart = resolveEntryDataOffset(
    localHeader,
    entry.localOffset,
    source.size,
    entry.name,
    entry.compressedSize,
  );

  // Whole-archive bombs are caught by the declared sizes in the central
  // directory; an individually highly-compressible entry is caught here.
  if (entry.uncompressedSize / Math.max(1, entry.compressedSize) > lim.maxCompressionRatio) {
    throw new ZipSecurityError(
      `Suspicious compression ratio in ${JSON.stringify(entry.name)} (${Math.round(
        entry.uncompressedSize / Math.max(1, entry.compressedSize),
      )}x > ${lim.maxCompressionRatio}x)`,
    );
  }

  let produced = 0;
  let crc = 0;
  const fail = (message: string): never => {
    throw new ZipSecurityError(`Unable to read ZIP archive: ${message}`);
  };

  if (entry.method === 0) {
    // Stored: the compressed bytes ARE the content; slice them out directly.
    let offset = dataStart;
    let remaining = entry.compressedSize;
    while (remaining > 0) {
      const chunk = await source.read(offset, Math.min(READ_CHUNK, remaining));
      if (chunk.length === 0) fail(`cannot read ${JSON.stringify(entry.name)}: unexpected end of archive`);
      offset += chunk.length;
      remaining -= chunk.length;
      produced += chunk.length;
      if (produced > entry.uncompressedSize) {
        fail(`size mismatch in ${JSON.stringify(entry.name)} (more than the declared ${entry.uncompressedSize} bytes)`);
      }
      crc = crc32(chunk, crc);
      yield chunk;
    }
  } else {
    // Deflate: feed compressed chunks to a push-based inflater and yield the
    // decompressed output it produces for each one.
    let pending: Uint8Array[] = [];
    const inflate = new Inflate({}, (chunk) => {
      pending.push(chunk);
    });
    let offset = dataStart;
    let remaining = entry.compressedSize;
    while (remaining > 0) {
      const chunk = await source.read(offset, Math.min(READ_CHUNK, remaining));
      if (chunk.length === 0) fail(`cannot inflate ${JSON.stringify(entry.name)}: unexpected end of archive`);
      offset += chunk.length;
      remaining -= chunk.length;
      pending = [];
      try {
        inflate.push(chunk, remaining === 0);
      } catch (error) {
        fail(`cannot inflate ${JSON.stringify(entry.name)}: ${(error as Error).message}`);
      }
      for (const out of pending) {
        produced += out.length;
        if (produced > entry.uncompressedSize) {
          fail(`size mismatch in ${JSON.stringify(entry.name)} (more than the declared ${entry.uncompressedSize} bytes)`);
        }
        crc = crc32(out, crc);
        yield out;
      }
      pending = [];
    }
  }

  if (produced !== entry.uncompressedSize) {
    fail(`size mismatch in ${JSON.stringify(entry.name)} (${produced} != ${entry.uncompressedSize})`);
  }
  if (entry.uncompressedSize > 0 && crc !== entry.crc32) {
    fail(`CRC mismatch in ${JSON.stringify(entry.name)}`);
  }
}

/**
 * Fully read one entry into a preallocated buffer of the DECLARED size
 * (used for the small parts: workbook, rels, shared strings, styles — never
 * for the big worksheet parts).
 */
export async function readEntry(
  source: ByteSource,
  entry: ZipDirEntry,
  limits: Partial<ReadLimits> = {},
): Promise<Uint8Array> {
  const out = new Uint8Array(entry.uncompressedSize);
  let at = 0;
  for await (const chunk of streamEntryChunks(source, entry, limits)) {
    out.set(chunk, at);
    at += chunk.length;
  }
  return out;
}
