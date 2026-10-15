/**
 * Streaming ZIP container writer — the write-side counterpart of
 * `zip/stream.ts` (bounded-memory reads) and of the in-memory
 * `serializeRawZip` in `raw.ts`.
 *
 * Entries are emitted as they are produced: a local header with the
 * general-purpose data-descriptor flag (bit 3) is written first, the content is
 * pushed through fflate's incremental deflate (`ZipDeflate`, or
 * `ZipPassThrough` at level 0), and the real CRC-32 plus both sizes follow in a
 * data descriptor. That is what makes a *streaming* ZIP possible: the sizes
 * are simply not known until the entry has been fed, so they cannot be written
 * in the local header. Both the eager reader (`zip/raw.ts`) and the streaming
 * reader (`zip/stream.ts`) already accept bit-3 entries, reading the real
 * CRC/sizes from the central directory.
 *
 * The central directory is written at the end (`finish()`); only bounded
 * metadata (name, CRC, sizes, offset) is held per entry. Zip64 is deliberately
 * NOT supported, exactly as on the read side: exceeding a 32-bit field throws a
 * stable `ZipUnsupportedError` instead of writing a corrupt archive.
 */
import { ZipDeflate, ZipPassThrough } from 'fflate';
import { isSafeEntryName } from './names';
import { SheetLiteError, ZipSecurityError, ZipUnsupportedError } from './errors';

// Signatures: local header, central directory, data descriptor, EOCD.
const LOCAL_SIG = 0x0403_4b50;
const CENTRAL_SIG = 0x0201_4b50;
const DESCRIPTOR_SIG = 0x0807_4b50;
const EOCD_SIG = 0x0605_4b50;

const LOCAL_FIXED_BYTES = 30;
const CENTRAL_FIXED_BYTES = 46;
const DESCRIPTOR_BYTES = 16;
const EOCD_BYTES = 22;

/** General purpose bit 3: CRC-32 + sizes follow the data, not the header. */
const FLAG_DATA_DESCRIPTOR = 0x0008;

/** Largest value a 32-bit ZIP size/offset field can hold before zip64 is needed. */
const MAX_32 = 0xffff_fffe;

const EMPTY = new Uint8Array(0);

/** Compression level accepted by the writer (0 = stored). */
export type ZipLevel = 0 | 1 | 2 | 3 | 4 | 5 | 6 | 7 | 8 | 9;

/**
 * Where the archive bytes go. Structural on purpose: a Node file descriptor,
 * a WHATWG writer wrapper or anything else with a synchronous `write`.
 */
export interface ZipSink {
  write(chunk: Uint8Array): void;
}

export interface ZipWriteOptions {
  /**
   * DEFLATE level 0-9 for every entry (0 = stored, no compression). Default 6.
   *
   * @throws {SheetLiteError} `BAD_LEVEL` at construction for anything else.
   */
  level?: ZipLevel;
  /**
   * DOS timestamp written on every entry (defaults to the fixed DOS epoch).
   * Read in UTC and clamped to 1980-2107, so the archive is reproducible.
   *
   * @throws {SheetLiteError} `BAD_MTIME` at construction for an invalid Date.
   */
  mtime?: Date;
}

interface CentralRecord {
  nameBytes: Uint8Array;
  crc: number;
  compressedSize: number;
  uncompressedSize: number;
  localOffset: number;
  method: 0 | 8;
  time: number;
  date: number;
}

interface OpenEntry {
  name: string;
  stream: ZipDeflate | ZipPassThrough;
  localOffset: number;
  method: 0 | 8;
  time: number;
  date: number;
  compressedSize: number;
}

/** MS-DOS timestamps cover 1980-01-01 .. 2107-12-31; anything else is clamped. */
const DOS_MIN = Date.UTC(1980, 0, 1);
const DOS_MAX = Date.UTC(2107, 11, 31, 23, 59, 59);

/**
 * Encode a Date as ZIP (MS-DOS) local time, clamped to the 1980-2107 range.
 * UTC on purpose: the archive must not depend on the machine's timezone.
 */
function dosDateTime(value: Date): { time: number; date: number } {
  const at = new Date(Math.min(DOS_MAX, Math.max(DOS_MIN, value.getTime())));
  return {
    time: (at.getUTCHours() << 11) | (at.getUTCMinutes() << 5) | (at.getUTCSeconds() >> 1),
    date: ((at.getUTCFullYear() - 1980) << 9) | ((at.getUTCMonth() + 1) << 5) | at.getUTCDate(),
  };
}

/** Fixed timestamp for entries without an mtime (1980-01-01 00:00). */
const DOS_EPOCH = { time: 0, date: (0 << 9) | (1 << 5) | 1 };

/**
 * Write a ZIP archive incrementally. Entries are strictly sequential (a ZIP
 * local header must immediately precede its data), so `begin()` closes the
 * previous entry first.
 *
 * ```ts
 * const zip = new StreamingZipWriter(sink, { level: 6 });
 * zip.begin('xl/worksheets/sheet1.xml');
 * zip.write(encoder.encode('<row r="1">…'));
 * zip.end();
 * zip.finish(); // central directory + EOCD
 * ```
 */
export class StreamingZipWriter {
  private readonly encoder = new TextEncoder();
  private readonly records: CentralRecord[] = [];
  private offset = 0;
  private open: OpenEntry | null = null;
  private finished = false;

  constructor(
    private readonly sink: ZipSink,
    private readonly options: ZipWriteOptions = {},
  ) {
    // Checked once, here: an option that cannot produce a valid archive must
    // not fail halfway through a multi-gigabyte write.
    const level = options.level;
    if (level !== undefined && (!Number.isInteger(level) || level < 0 || level > 9)) {
      throw new SheetLiteError(
        `Invalid compression level ${level}: must be an integer 0-9`,
        'BAD_LEVEL',
      );
    }
    const mtime = options.mtime;
    if (mtime !== undefined && !(mtime instanceof Date && Number.isFinite(mtime.getTime()))) {
      throw new SheetLiteError('Invalid mtime: must be a valid Date', 'BAD_MTIME');
    }
  }

  /** Entries written so far (local headers emitted, central records pending). */
  get entryCount(): number {
    return this.records.length;
  }

  /** Bytes of the archive produced so far (local headers + data so far). */
  get bytesWritten(): number {
    return this.offset;
  }

  /** True once `finish()` has written the central directory and the EOCD. */
  get done(): boolean {
    return this.finished;
  }

  /** Start a new entry, closing the previous one if it is still open. */
  begin(name: string): void {
    if (this.finished) {
      throw new SheetLiteError(
        'The archive is already finished; no further entries can be added',
        'BAD_STATE',
      );
    }
    if (!isSafeEntryName(name)) {
      throw new ZipSecurityError(`Unsafe entry name rejected: ${JSON.stringify(name)}`);
    }
    if (this.open) this.end();
    if (this.records.length >= 0xffff) {
      throw new ZipUnsupportedError(
        'ZIP feature not supported: more than 65535 entries would require zip64 (not supported)',
      );
    }

    const { time, date } = this.options.mtime ? dosDateTime(this.options.mtime) : DOS_EPOCH;
    const nameBytes = this.encoder.encode(name);
    const method: 0 | 8 = (this.options.level ?? 6) === 0 ? 0 : 8;
    const localOffset = this.offset;

    const header = new Uint8Array(LOCAL_FIXED_BYTES + nameBytes.length);
    const dv = new DataView(header.buffer);
    dv.setUint32(0, LOCAL_SIG, true);
    dv.setUint16(4, method === 0 ? 10 : 20, true); // version needed to extract
    dv.setUint16(6, FLAG_DATA_DESCRIPTOR, true); // general purpose flags
    dv.setUint16(8, method, true);
    dv.setUint16(10, time, true);
    dv.setUint16(12, date, true);
    // CRC-32 and both sizes are zero in the local header: they follow the data.
    dv.setUint16(26, nameBytes.length, true);
    dv.setUint16(28, 0, true); // no extra field
    header.set(nameBytes, LOCAL_FIXED_BYTES);
    this.emit(header);

    const entry: OpenEntry = {
      name,
      localOffset,
      method,
      time,
      date,
      compressedSize: 0,
      // Replaced immediately below by the real compressor.
      stream: EMPTY as unknown as ZipDeflate,
    };
    if (method === 0) {
      entry.stream = new ZipPassThrough(name);
    } else {
      entry.stream = new ZipDeflate(name, { level: this.options.level ?? 6 });
    }
    entry.stream.ondata = (error, chunk) => {
      if (error) {
        throw new SheetLiteError(
          `Unable to compress ${JSON.stringify(entry.name)}: ${error.message}`,
          'BAD_STATE',
        );
      }
      if (chunk && chunk.length > 0) {
        entry.compressedSize += chunk.length;
        this.emit(chunk);
      }
    };
    this.open = entry;
  }

  /** Append (uncompressed) bytes to the current entry. */
  write(bytes: Uint8Array): void {
    if (bytes.length === 0) return;
    if (this.finished) {
      throw new SheetLiteError(
        'The archive is already finished; no further bytes can be written',
        'BAD_STATE',
      );
    }
    if (!this.open) {
      throw new SheetLiteError('No open ZIP entry: call begin(name) before write()', 'BAD_STATE');
    }
    this.open.stream.push(bytes, false);
  }

  /** Close the current entry: flush the deflate stream and add its data descriptor. */
  end(): void {
    const entry = this.open;
    if (!entry) return;
    this.open = null;
    entry.stream.push(EMPTY, true);

    const descriptor = new Uint8Array(DESCRIPTOR_BYTES);
    const dv = new DataView(descriptor.buffer);
    dv.setUint32(0, DESCRIPTOR_SIG, true);
    dv.setUint32(4, entry.stream.crc >>> 0, true);
    dv.setUint32(8, entry.compressedSize, true);
    dv.setUint32(12, entry.stream.size, true);
    this.emit(descriptor);

    this.records.push({
      nameBytes: this.encoder.encode(entry.name),
      crc: entry.stream.crc >>> 0,
      compressedSize: entry.compressedSize,
      uncompressedSize: entry.stream.size,
      localOffset: entry.localOffset,
      method: entry.method,
      time: entry.time,
      date: entry.date,
    });
  }

  /**
   * Write the central directory and the end-of-central-directory record. The
   * archive is complete after this call; no further entries can be added.
   */
  finish(): void {
    if (this.finished) return;
    if (this.open) this.end();
    const centralOffset = this.offset;
    let centralSize = 0;
    for (const record of this.records) {
      const size = CENTRAL_FIXED_BYTES + record.nameBytes.length;
      this.check32(centralOffset + centralSize + size);
      const bytes = new Uint8Array(size);
      const dv = new DataView(bytes.buffer);
      dv.setUint32(0, CENTRAL_SIG, true);
      dv.setUint16(4, 0x0014, true); // version made by (MS-DOS, 2.0)
      dv.setUint16(6, record.method === 0 ? 10 : 20, true); // version needed
      dv.setUint16(8, FLAG_DATA_DESCRIPTOR, true);
      dv.setUint16(10, record.method, true);
      dv.setUint16(12, record.time, true);
      dv.setUint16(14, record.date, true);
      dv.setUint32(16, record.crc, true);
      dv.setUint32(20, record.compressedSize, true);
      dv.setUint32(24, record.uncompressedSize, true);
      dv.setUint16(28, record.nameBytes.length, true);
      dv.setUint16(30, 0, true); // extra field length
      dv.setUint16(32, 0, true); // comment length
      dv.setUint16(34, 0, true); // disk number start
      dv.setUint16(36, 0, true); // internal attributes
      dv.setUint32(38, 0, true); // external attributes
      dv.setUint32(42, record.localOffset, true);
      bytes.set(record.nameBytes, CENTRAL_FIXED_BYTES);
      this.emit(bytes);
      centralSize += size;
    }

    const eocd = new Uint8Array(EOCD_BYTES);
    const dv = new DataView(eocd.buffer);
    dv.setUint32(0, EOCD_SIG, true);
    dv.setUint16(4, 0, true); // this disk
    dv.setUint16(6, 0, true); // disk with the central directory
    dv.setUint16(8, this.records.length, true);
    dv.setUint16(10, this.records.length, true);
    dv.setUint32(12, centralSize, true);
    dv.setUint32(16, centralOffset, true);
    dv.setUint16(20, 0, true); // no archive comment
    this.emit(eocd);
    this.finished = true;
  }

  private emit(chunk: Uint8Array): void {
    this.check32(this.offset + chunk.length);
    this.sink.write(chunk);
    this.offset += chunk.length;
  }

  private check32(end: number): void {
    if (end > MAX_32) {
      throw new ZipUnsupportedError(
        'ZIP feature not supported: output exceeds the 32-bit ZIP size limit (zip64 is not supported)',
      );
    }
  }
}