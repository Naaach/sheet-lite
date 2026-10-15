/**
 * Incremental workbook creation (the `sheet-lite/write` entry point).
 *
 * `WorkbookWriter` builds an .xlsx package from scratch and streams it out
 * without ever holding the worksheet XML or the archive in memory: each row is
 * serialized, pushed through fflate's incremental deflate and appended to the
 * target as a ZIP entry, and only bounded metadata (row counts, the style
 * registry, one row) is kept alive.
 *
 * This is the WRITE counterpart of `StreamingWorkbook` (`sheet-lite/stream`):
 * both sit on the same hardened ZIP layer (`zip/raw.ts` for the container rules,
 * `zip/stream.ts` for incremental IO). It does NOT replace the eager `Workbook`
 * write path — editing an existing file keeps every untouched part
 * byte-identical, which a from-scratch writer cannot promise.
 *
 * ```ts
 * const wb = await WorkbookWriter.create({ file: './huge.xlsx' });
 * const ws = wb.addWorksheet('Data', { views: [{ state: 'frozen', ySplit: 1 }] });
 * ws.getColumn('A').width = 20;
 * for (let i = 1; i <= 100_000; i++) ws.addRow([i, `row ${i}`, i / 2]);
 * await wb.commit();   // …or: const bytes = await WorkbookWriter.create().toBuffer()
 * ```
 *
 * A `{ file }` target is written ATOMICALLY: the bytes go to a sibling
 * temporary file that is renamed over the destination only once the archive is
 * complete, so a failure in the middle of a row loop leaves the previous file
 * exactly as it was instead of a truncated .xlsx. See `WorkbookWriterOptions`.
 */
import { StylesBuilder } from '../writer/style-builder';
import { SheetLiteError } from '../zip/errors';
import { validateSheetName } from '../opc/sheet-name';
import {
  CONTENT_TYPES_XMLNS,
  MAIN_CONTENT_TYPE,
  OFFICE_DOCUMENT_REL,
  OFFICE_XMLNS,
  RELS_CONTENT_TYPE,
  SHEET_CONTENT_TYPE,
  SHEET_NS,
  STYLES_CONTENT_TYPE,
  WORKSHEET_REL,
  XML_CONTENT_TYPE,
  XML_DECL,
} from '../writer/opc-constants';
import type { CellStyleOverrides } from '../model/style';
import type { SheetViewProps } from '../model/view';
import { escapeXml } from '../xml/escape';
import { StreamingZipWriter, type ZipLevel, type ZipSink } from '../zip/zip-stream';
import { WriteWorksheet, type WriteContext } from './write-worksheet';

const OFFICE_DOCUMENT_REL_TYPE = OFFICE_DOCUMENT_REL;

export interface WorkbookWriterOptions {
  /**
   * Node file path: the archive is streamed straight to this file as rows are
   * added, so memory stays bounded whatever the sheet size.
   *
   * The write is ATOMIC. Bytes are appended to a sibling temporary file
   * (`<file>.<pid>.tmp`) that is opened on the FIRST byte written — not in
   * `create()` — and renamed over `file` only after the archive has been
   * finished. Until that rename the destination is untouched, so a failure
   * anywhere (a bad value, a full disk, `abort()`) leaves the previous file
   * byte-for-byte intact and removes the scratch file. The destination is
   * therefore either the old workbook or a complete new one, never a
   * half-written one.
   *
   * A `{ sink }` target cannot offer that guarantee: bytes handed to
   * `write(chunk)` are the caller's to deal with — a socket, a pipe or an
   * upload cannot be recalled — so the archive is only complete once the sink
   * has received everything, and `abort()` cannot undo what it already took.
   */
  file?: string;
  /**
   * Custom output target — anything with a synchronous
   * `write(chunk: Uint8Array): void` (a WHATWG `WritableStreamDefaultWriter`, a
   * socket wrapper, a test spy…). Takes precedence over `file`.
   */
  sink?: ZipSink;
  /** DEFLATE level 0-9 (0 = stored, no compression). Default 6. */
  level?: ZipLevel;
  /** Write serial dates on the 1904 date system (default: 1900, like Excel). */
  date1904?: boolean;
  /** DOS timestamp of every ZIP entry. Defaults to the fixed DOS epoch. */
  mtime?: Date;
}

/** Collects the archive in memory (the default target; bounded by file size). */
class MemorySink implements ZipSink {
  readonly chunks: Uint8Array[] = [];
  length = 0;

  write(chunk: Uint8Array): void {
    this.chunks.push(chunk);
    this.length += chunk.length;
  }

  concat(): Uint8Array {
    const out = new Uint8Array(this.length);
    let at = 0;
    for (const chunk of this.chunks) {
      out.set(chunk, at);
      at += chunk.length;
    }
    return out;
  }

  /** Drop everything collected so far (an aborted writer holds nothing). */
  clear(): void {
    this.chunks.length = 0;
    this.length = 0;
  }
}

/**
 * The `{ file }` target: a lazily-opened descriptor on a sibling temporary
 * file, renamed over the destination only when the archive is complete.
 *
 * Nothing reaches `file` before `commit()`, so an exception in the middle of a
 * row loop can never leave a truncated workbook where the user's file was —
 * the usual "never break your file" promise, on a from-scratch writer.
 */
class FileTarget implements ZipSink {
  private fd = -1;
  /** Scratch file, or null once it has been renamed away (or removed). */
  private temp: string | null;

  constructor(
    private readonly fs: typeof import('node:fs'),
    private readonly finalPath: string,
  ) {
    // Same directory as the target, so the rename is a same-volume rename.
    // The pid keeps two processes writing the same path out of each other's way.
    this.temp = `${finalPath}.${process.pid}.tmp`;
  }

  /** Append archive bytes, opening the scratch file on the first one. */
  write(chunk: Uint8Array): void {
    // Opening here, not in create(): a workbook that is built and then never
    // committed (or aborted) must not have touched the destination at all.
    const temp = this.temp;
    if (temp === null) return; // rolled back: discard anything still in flight
    if (this.fd < 0) this.fd = this.fs.openSync(temp, 'w');
    this.fs.writeSync(this.fd, chunk, 0, chunk.length, null);
  }

  /** Close the descriptor and move the finished archive into place. */
  commit(): void {
    const temp = this.temp;
    if (temp === null) return;
    this.closeFd();
    // Only after the rename succeeds is there nothing left to roll back, so a
    // failure here must leave `temp` set for abort() to delete.
    this.fs.renameSync(temp, this.finalPath);
    this.temp = null;
  }

  /** Close the descriptor and delete the scratch file. No-op once committed. */
  abort(): void {
    const temp = this.temp;
    this.temp = null;
    this.closeFd();
    if (temp === null) return;
    try {
      this.fs.rmSync(temp, { force: true });
    } catch {
      // Best effort: never let cleanup mask (or replace) the error that caused
      // the rollback in the first place.
    }
  }

  private closeFd(): void {
    if (this.fd < 0) return;
    this.fs.closeSync(this.fd);
    this.fd = -1;
  }
}

/**
 * `Symbol.asyncDispose` where explicit resource management exists, the
 * polyfill's well-known symbol where it does not — read once, so `await using`
 * works on every runtime that can run it.
 */
const ASYNC_DISPOSE: typeof Symbol.asyncDispose =
  Symbol.asyncDispose ?? (Symbol.for('Symbol.asyncDispose') as typeof Symbol.asyncDispose);

const hasStyleOverrides = (patch: CellStyleOverrides): boolean =>
  patch.numFmt !== undefined ||
  patch.font !== undefined ||
  patch.fill !== undefined ||
  patch.border !== undefined ||
  patch.alignment !== undefined;

/** One `<sheet>` registration in `xl/workbook.xml`. */
interface SheetEntry {
  name: string;
  sheetId: number;
  rId: string;
  /** Part key, e.g. "xl/worksheets/sheet1.xml". */
  partName: string;
}

/** Bridges a worksheet to its workbook (style registry + entry sequencing). */
class WriterContext implements WriteContext {
  constructor(private readonly workbook: WorkbookWriter) {}

  activate(sheet: WriteWorksheet): void {
    this.workbook.activateSheet(sheet);
  }

  resolveStyle(patch: CellStyleOverrides): number {
    return this.workbook.resolveStyleIndex(patch);
  }

  get date1904(): boolean {
    return this.workbook.date1904;
  }

  get sealed(): boolean {
    return this.workbook.sealed;
  }
}

/**
 * A workbook that is CREATED and written incrementally. Use the async factory:
 * a file target is resolved (but not opened) there, and the destination is only
 * replaced by `commit()`. `Workbook.create()` on the main entry is a lazy alias
 * for `WorkbookWriter.create()`.
 */
export class WorkbookWriter {
  private readonly zip: StreamingZipWriter;
  private readonly styles = new StylesBuilder();
  private readonly sheetList: WriteWorksheet[] = [];
  /** The `{ file }` target; null for a custom sink or a buffered writer. */
  private readonly file: FileTarget | null;
  private readonly memory: MemorySink | null;
  private readonly context: WriteContext;
  private readonly encoder = new TextEncoder();
  /** Serial dates are written on this date system. */
  readonly date1904: boolean;
  private activeSheet: WriteWorksheet | null = null;
  private committed = false;

  private constructor(sink: ZipSink, file: FileTarget | null, options: WorkbookWriterOptions) {
    this.file = file;
    this.memory = sink instanceof MemorySink ? sink : null;
    this.date1904 = options.date1904 ?? false;
    this.zip = new StreamingZipWriter(sink, { level: options.level, mtime: options.mtime });
    this.context = new WriterContext(this);
  }

  /**
   * Create a writer.
   *
   * - `{ file: path }` — streams straight to a Node file (constant memory),
   *   atomically: see `WorkbookWriterOptions.file`;
   * - `{ sink }` — streams to any `{ write(chunk) }` target;
   * - no target — buffers, so `toBuffer()` / `writeFile()` can return the file.
   *
   * A file target touches nothing on disk here; the scratch file appears on the
   * first row written and becomes the destination on `commit()`.
   */
  static async create(options: WorkbookWriterOptions = {}): Promise<WorkbookWriter> {
    if (options.sink) return new WorkbookWriter(options.sink, null, options);
    if (options.file) {
      // Node-only, and dynamic: keeps `node:` out of the static dependency graph.
      const fs = await import('node:fs');
      const file = new FileTarget(fs, options.file);
      return new WorkbookWriter(file, file, options);
    }
    const memory = new MemorySink();
    return new WorkbookWriter(memory, null, options);
  }

  /** Sheets added so far, in creation order. */
  get worksheets(): WriteWorksheet[] {
    return this.sheetList;
  }

  get sheetNames(): string[] {
    return this.sheetList.map((sheet) => sheet.name);
  }

  getWorksheet(nameOrId: string | number): WriteWorksheet | undefined {
    if (typeof nameOrId === 'number') {
      return this.sheetList.find((sheet) => sheet.id === nameOrId);
    }
    return this.sheetList.find((sheet) => sheet.name === nameOrId);
  }

  /** True once `commit()` (or `abort()`) has finished the archive. */
  get sealed(): boolean {
    return this.committed;
  }

  /**
   * Add a worksheet (ExcelJS-style). Its part is streamed as rows arrive; the
   * registration in `xl/workbook.xml`, the workbook relationships and
   * `[Content_Types].xml` are emitted at commit time.
   *
   * Accepts `options.views` (frozen panes), like `Workbook.addWorksheet`.
   * Throws `SHEET_EXISTS` for a duplicate name and `BAD_SHEET_NAME` for names
   * Excel would refuse.
   *
   * Rows must be appended to one sheet at a time: a ZIP entry's bytes cannot be
   * interleaved, so returning to a sheet that already has rows throws
   * `SHEET_INTERLEAVED`.
   */
  addWorksheet(name: string, options?: { views?: SheetViewProps[] }): WriteWorksheet {
    validateSheetName(name);
    if (this.committed) {
      throw new SheetLiteError(
        'Cannot add a worksheet: the workbook has already been committed',
        'SHEET_FINALIZED',
      );
    }
    if (this.sheetList.some((sheet) => sheet.name === name)) {
      throw new SheetLiteError(`A worksheet named "${name}" already exists`, 'SHEET_EXISTS');
    }
    const id = this.sheetList.length + 1;
    const partName = `xl/worksheets/sheet${id}.xml`;
    const sheet = new WriteWorksheet(
      name,
      id,
      partName,
      this.zip,
      this.date1904,
      this.context,
      options?.views,
    );
    this.sheetList.push(sheet);
    return sheet;
  }

  /**
   * Finish the workbook: the pending row of every sheet, the sheet parts, the
   * style sheet and the OPC registration parts are written and the target is
   * closed. Idempotent; a workbook needs at least one sheet (`NO_SHEETS`).
   *
   * Sheet parts are closed in ZIP-entry order (see `closeSheetParts`); for a
   * `{ file }` target the destination is replaced only once the whole archive
   * is on disk. After this, `addRow()` / `addWorksheet()` throw
   * `SHEET_FINALIZED`.
   */
  async commit(): Promise<void> {
    if (this.committed) return;
    if (this.sheetList.length === 0) {
      throw new SheetLiteError(
        'Cannot write a workbook without worksheets: add at least one with addWorksheet(name)',
        'NO_SHEETS',
      );
    }
    try {
      this.closeSheetParts();

      this.writeEntry('xl/styles.xml', this.styles.toXml());
      const entries: SheetEntry[] = this.sheetList.map((sheet, index) => ({
        name: sheet.name,
        sheetId: index + 1,
        rId: `rId${index + 1}`,
        partName: sheet.partName,
      }));
      this.writeEntry('xl/workbook.xml', workbookXml(entries, this.date1904));
      this.writeEntry('xl/_rels/workbook.xml.rels', workbookRelsXml(entries));
      this.writeEntry('[Content_Types].xml', contentTypesXml(entries));
      this.writeEntry('_rels/.rels', rootRelsXml());
      this.zip.finish();
      // Last step, and the only one that touches the destination: everything
      // above only ever reached the scratch file.
      this.file?.commit();
    } catch (error) {
      // The archive is unusable: seal the writer and roll the target back, so
      // a failure is neither silent nor destructive.
      this.committed = true;
      this.rollbackTarget();
      throw error;
    }
    this.committed = true;
  }

  /**
   * Give up on the workbook without writing it.
   *
   * For a `{ file }` target this closes the descriptor and deletes the scratch
   * file, so the destination keeps the bytes it had — which is what an
   * exception thrown by the caller's own row loop needs: the file they were
   * about to overwrite is still there, and no `.tmp` is left next to it. A
   * buffered target simply drops what it collected; a `{ sink }` target cannot
   * be recalled (see `WorkbookWriterOptions.sink`).
   *
   * Idempotent, seals the writer (further writes throw `SHEET_FINALIZED`),
   * and a no-op once `commit()` has completed.
   */
  abort(): void {
    if (this.committed) return;
    this.committed = true;
    this.rollbackTarget();
  }

  /**
   * `await using` support: leaving the scope without a `commit()` aborts, so a
   * forgotten commit can never leave a scratch file behind.
   */
  [ASYNC_DISPOSE](): Promise<void> {
    this.abort();
    return Promise.resolve();
  }

  /**
   * Commit and return the workbook bytes. Buffer target only: with a file or a
   * custom sink the archive is already on its way out (`BAD_TARGET`).
   */
  async toBuffer(): Promise<Uint8Array> {
    if (!this.memory) {
      throw new SheetLiteError(
        'toBuffer() needs a buffered workbook: this writer streams to a file or a sink, use commit() instead',
        'BAD_TARGET',
      );
    }
    await this.commit();
    return this.memory.concat();
  }

  /** Commit and write the workbook to a Node file. Buffer target only. */
  async writeFile(path: string): Promise<void> {
    const bytes = await this.toBuffer();
    const fs = await import('node:fs/promises');
    await fs.writeFile(path, bytes);
  }

  /**
   * @internal Make `sheet` the sheet currently being written.
   *
   * Finalising the previous one here is what keeps the "at most one open ZIP
   * entry" invariant that `closeSheetParts` relies on: a sheet that has rows
   * can no longer become the active one, so the only sheet that can hold an
   * open entry at `commit()` time is the active one.
   */
  activateSheet(sheet: WriteWorksheet): void {
    if (this.activeSheet === sheet) return;
    if (sheet.rowCount > 0) {
      throw new SheetLiteError(
        `Cannot append more rows to worksheet "${sheet.name}": its rows must be written in a single pass (a ZIP entry's bytes cannot be interleaved)`,
        'SHEET_INTERLEAVED',
      );
    }
    this.activeSheet?.finalize();
    this.activeSheet = sheet;
  }

  /**
   * Close every sheet part, **in ZIP-entry order**.
   *
   * A ZIP entry's bytes must be contiguous: the moment a sheet flushes its
   * first row, `begin()` has emitted its local header, and `StreamingZipWriter`
   * silently closes that entry if another one starts. So the rule is simply
   * "whoever is holding the open entry closes first", and:
   *
   * - at most one sheet can hold an open entry — `activateSheet` finalises the
   *   previous sheet, and `finalize()` is idempotent — so that sheet is the
   *   ACTIVE one (or has no entry open yet at all, because a single row is
   *   still buffered);
   * - every remaining sheet then opens and closes its own entry in creation
   *   order, because `SheetStreamWriter.finish()` emits the prologue when the
   *   part has not started yet, so `<sheetViews>` / `<cols>` still precede
   *   `<sheetData>` (and an untouched sheet is emitted with an empty
   *   `<sheetData/>`).
   *
   * Entry order need not match sheet order: `xl/workbook.xml` refers to sheets
   * through relationship ids, never through their position in the archive.
   */
  private closeSheetParts(): void {
    this.activeSheet?.finalize();
    this.activeSheet = null;
    for (const sheet of this.sheetList) sheet.finalize();
  }

  /** Drop everything the current target has accumulated so far. */
  private rollbackTarget(): void {
    this.file?.abort();
    this.memory?.clear();
  }

  /** @internal Resolve a cell's style patch against the workbook style sheet. */
  resolveStyleIndex(patch: CellStyleOverrides): number {
    return hasStyleOverrides(patch) ? this.styles.resolveIndex(patch) : 0;
  }

  /** Write one small part as a complete ZIP entry. */
  private writeEntry(name: string, xml: string): void {
    this.zip.begin(name);
    this.zip.write(this.encoder.encode(xml));
    this.zip.end();
  }
}

// --- package parts ---------------------------------------------------------

function workbookXml(sheets: SheetEntry[], date1904: boolean): string {
  const entries = sheets
    .map((sheet) => `<sheet name="${escapeXml(sheet.name)}" sheetId="${sheet.sheetId}" r:id="${sheet.rId}"/>`)
    .join('');
  const pr = date1904 ? '<workbookPr date1904="1"/>' : '';
  return `${XML_DECL}<workbook ${SHEET_NS}>${pr}<sheets>${entries}</sheets></workbook>`;
}

function workbookRelsXml(sheets: SheetEntry[]): string {
  const entries = sheets
    .map(
      (sheet) =>
        `<Relationship Id="${sheet.rId}" Type="${WORKSHEET_REL}" Target="${escapeXml(
          sheet.partName.replace(/^xl\//, ''),
        )}"/>`,
    )
    .join('');
  return `${XML_DECL}<Relationships ${OFFICE_XMLNS}>${entries}</Relationships>`;
}

function contentTypesXml(sheets: SheetEntry[]): string {
  const overrides = sheets
    .map((sheet) => `<Override PartName="/${sheet.partName}" ContentType="${SHEET_CONTENT_TYPE}"/>`)
    .join('');
  return (
    `${XML_DECL}<Types ${CONTENT_TYPES_XMLNS}>` +
    `<Default Extension="rels" ContentType="${RELS_CONTENT_TYPE}"/>` +
    `<Default Extension="xml" ContentType="${XML_CONTENT_TYPE}"/>` +
    `<Override PartName="/xl/workbook.xml" ContentType="${MAIN_CONTENT_TYPE}"/>` +
    `<Override PartName="/xl/styles.xml" ContentType="${STYLES_CONTENT_TYPE}"/>` +
    overrides +
    '</Types>'
  );
}

function rootRelsXml(): string {
  return (
    `${XML_DECL}<Relationships ${OFFICE_XMLNS}>` +
    `<Relationship Id="rId1" Type="${OFFICE_DOCUMENT_REL_TYPE}" Target="xl/workbook.xml"/>` +
    '</Relationships>'
  );
}
