/**
 * Streaming (read-only) workbook API.
 *
 * `StreamingWorkbook` reads the small package parts (workbook.xml, its
 * relationships, sharedStrings, styles) and then streams worksheet rows one
 * at a time, inflating the sheet part incrementally through the hardened
 * streaming ZIP layer (`zip/stream.ts`). Peak memory is bounded by the chunk
 * size plus one row — a multi-gigabyte file reads with a few MB of heap.
 *
 * Deliberately READ-ONLY: streaming mode has no patch/write path. Editing
 * requires the eager `Workbook`, which keeps every untouched part in memory
 * so it can be re-serialized byte-for-byte.
 */
import type { ReadLimits } from '../zip/raw';
import { assertNotCompoundFile } from '../zip/raw';
import {
  BufferByteSource,
  FileByteSource,
  readZipDirectory,
  readEntry,
  streamEntryChunks,
  type ByteSource,
  type ZipDirectory,
} from '../zip/stream';
import { SheetLiteError } from '../zip/errors';
import { readPackageMeta, type PackageMeta, type SheetRef } from '../reader/package-reader';
import { SheetRowReader, type StreamRow, type StreamCell } from '../reader/stream-sheet';

export type { StreamRow, StreamCell };

const toPartKey = (partName: string): string => partName.replace(/^\//, '');

/** Parts read eagerly by `open()`; everything else stays on disk until streamed. */
const SMALL_PARTS = [
  'xl/workbook.xml',
  // Only present in `.xlsb` packages; reading it (when it exists at all) lets
  // `readPackageMeta` name the binary format instead of a missing part, so
  // the streaming reader refuses `.xlsb` exactly like the eager one does.
  'xl/workbook.bin',
  'xl/_rels/workbook.xml.rels',
  'xl/sharedStrings.xml',
  'xl/styles.xml',
] as const;

export type StreamInput = string | Uint8Array | ArrayBuffer | ArrayBufferView;

/** One sheet of a streaming workbook: list it, or iterate its rows. */
export class StreamingSheet {
  constructor(
    readonly name: string,
    readonly id: number,
    /** Store key of the sheet part (no leading slash), e.g. "xl/worksheets/sheet1.xml". */
    readonly partName: string,
    private readonly workbook: StreamingWorkbook,
  ) {}

  /**
   * Iterate the sheet's rows lazily: each call starts a fresh pass over the
   * sheet part (rows are re-read, never buffered). The yielded values use the
   * same `CellValue` model as the eager reader.
   */
  readRows(): AsyncGenerator<StreamRow> {
    return this.workbook.readRows(this.name);
  }
}

export class StreamingWorkbook {
  private constructor(
    private readonly source: ByteSource,
    private readonly dir: ZipDirectory,
    private readonly meta: PackageMeta,
    private readonly limits: Partial<ReadLimits>,
    private readonly fileSource: FileByteSource | null,
  ) {}

  /**
   * Open a workbook for streaming reads.
   *
   * - `path` (string): a Node file path — the archive is never loaded into
   *   memory, only the central directory and the parts/rows actually read.
   * - a `Uint8Array`/`ArrayBuffer`/view: works in any runtime; the input
   *   buffer itself is caller-owned, entries are still inflated incrementally.
   *
   * A `ReadableStream` is intentionally NOT accepted: ZIP streams have no
   * seeking and the EOCD lives at the end, so a stream would have to be fully
   * buffered anyway — pass a file path instead.
   */
  static async open(input: StreamInput, limits: Partial<ReadLimits> = {}): Promise<StreamingWorkbook> {
    if (typeof input === 'string') {
      // Node-only: dynamic import keeps `node:` out of the static dependency graph.
      const fileSource = await FileByteSource.open(input);
      try {
        // A legacy `.xls` / encrypted OOXML container is not a broken archive.
        assertNotCompoundFile(await fileSource.read(0, 8));
        return await StreamingWorkbook.fromSource(fileSource, limits, fileSource);
      } catch (error) {
        await fileSource.close();
        throw error;
      }
    }
    const bytes =
      input instanceof Uint8Array
        ? input
        : ArrayBuffer.isView(input)
          ? new Uint8Array(input.buffer, input.byteOffset, input.byteLength)
          : new Uint8Array(input);
    assertNotCompoundFile(bytes);
    return StreamingWorkbook.fromSource(new BufferByteSource(bytes), limits, null);
  }

  private static async fromSource(
    source: ByteSource,
    limits: Partial<ReadLimits>,
    fileSource: FileByteSource | null,
  ): Promise<StreamingWorkbook> {
    const dir = await readZipDirectory(source, limits);
    const parts = new Map<string, Uint8Array>();
    for (const name of SMALL_PARTS) {
      const entry = dir.byName.get(name);
      if (entry) parts.set(name, await readEntry(source, entry, limits));
    }
    const meta = readPackageMeta({ get: (name) => parts.get(name) });
    return new StreamingWorkbook(source, dir, meta, limits, fileSource);
  }

  /** Sheet list (name / id / part name), like `Workbook.metadata.sheets`. */
  get sheets(): SheetRef[] {
    return this.meta.sheets;
  }

  get sheetNames(): string[] {
    return this.meta.sheets.map((sheet) => sheet.name);
  }

  getSheet(nameOrId: string | number): StreamingSheet | undefined {
    const ref =
      typeof nameOrId === 'number'
        ? this.meta.sheets.find((sheet) => Number(sheet.sheetId) === nameOrId)
        : this.meta.sheets.find((sheet) => sheet.name === nameOrId);
    if (!ref) return undefined;
    return new StreamingSheet(ref.name, Number(ref.sheetId), toPartKey(ref.partName), this);
  }

  /**
   * Iterate a sheet's rows one at a time. Each call re-reads the sheet part
   * from the source; rows are produced lazily as the decompressed bytes flow.
   *
   * Throws `SHEET_NOT_FOUND` for an unknown sheet, `NOT_XLSX` if the sheet's
   * part is missing from the package.
   */
  readRows(sheet: string | number): AsyncGenerator<StreamRow> {
    const selected =
      typeof sheet === 'number'
        ? this.meta.sheets.find((s) => Number(s.sheetId) === sheet)
        : this.meta.sheets.find((s) => s.name === sheet);
    if (!selected) {
      throw new SheetLiteError(`Unknown sheet: ${typeof sheet === 'number' ? sheet : JSON.stringify(sheet)}`, 'SHEET_NOT_FOUND');
    }
    const key = toPartKey(selected.partName);
    return this.iteratePart(key);
  }

  private async *iteratePart(key: string): AsyncGenerator<StreamRow> {
    const entry = this.dir.byName.get(key);
    if (!entry) {
      throw new SheetLiteError(`Not an XLSX package: missing ${key}`, 'NOT_XLSX');
    }
    const reader = new SheetRowReader(this.meta);
    for await (const chunk of streamEntryChunks(this.source, entry, this.limits)) {
      reader.write(chunk);
      for (;;) {
        const row = reader.next();
        if (row === null) break;
        yield row;
      }
    }
    reader.end();
    for (;;) {
      const row = reader.next();
      if (row === null) break;
      yield row;
    }
  }

  /** Close the underlying file handle (file inputs only; buffer inputs are closed implicitly). */
  async close(): Promise<void> {
    if (this.fileSource) await this.fileSource.close();
  }
}