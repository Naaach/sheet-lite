/**
 * `sheet-lite/write` — incremental, bounded-memory workbook creation.
 *
 * Shipped separately from the main entry (like `sheet-lite/stream`, and for the
 * same reason: the main entry has a hard gzip budget) so a file that is only
 * READ never pays for the writer:
 *
 * ```ts
 * import { WorkbookWriter } from 'sheet-lite/write';
 *
 * const wb = await WorkbookWriter.create({ file: './huge.xlsx' });
 * const ws = wb.addWorksheet('Data');
 * for (let i = 0; i < 1_000_000; i++) ws.addRow([i, `row ${i}`]);
 * await wb.commit();
 * ```
 *
 * `Workbook.create(...)` on the main entry is a lazy alias for
 * `WorkbookWriter.create(...)` here.
 *
 * Strings are written INLINE (`t="inlineStr"`): a shared-string table grows
 * with the number of distinct strings, so it would make the memory bound
 * unbounded by definition — the exact trade-off this writer refuses to make.
 */
export { WorkbookWriter, type WorkbookWriterOptions } from './api/write-workbook';
export { WriteWorksheet, WriteRow, type WriteContext } from './api/write-worksheet';
export { WriteCell } from './api/write-cell';
export { WriteColumn } from './api/write-column';
export { StylesBuilder } from './writer/style-builder';
export {
  StreamingZipWriter,
  type ZipSink,
  type ZipWriteOptions,
  type ZipLevel,
} from './zip/zip-stream';
export {
  MAX_ROWS,
  MAX_COLUMNS,
  SheetStreamWriter,
  assertWritableValue,
  streamCellXml,
} from './writer/sheet-stream-writer';