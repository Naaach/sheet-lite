/**
 * `sheet-lite/stream` — streaming (read-only) entry point.
 *
 * Shipped separately from the main entry so the eager API bundle keeps its
 * size budget; the two share the same hardened ZIP/XML core:
 *
 * ```ts
 * import { StreamingWorkbook } from 'sheet-lite/stream';
 *
 * const wb = await StreamingWorkbook.open('./big.xlsx');   // Node file path
 * for (const sheet of wb.sheets) console.log(sheet.name, sheet.sheetId);
 * for await (const row of wb.getSheet('Data')!.readRows()) {
 *   console.log(row.number, row.cells.map((c) => c.value));
 * }
 * await wb.close();
 * ```
 *
 * `Workbook.openStream(...)` (main entry) is a lazy alias for
 * `StreamingWorkbook.open` here.
 */
export {
  StreamingWorkbook,
  StreamingSheet,
  type StreamInput,
  type StreamRow,
  type StreamCell,
} from './api/streaming';
export {
  readZipDirectory,
  streamEntryChunks,
  readEntry,
  FileByteSource,
  BufferByteSource,
  type ByteSource,
  type ZipDirectory,
  type RandomAccessFile,
} from './zip/stream';
export type { EocdInfo, ZipDirEntry, ReadLimits } from './zip/raw';
export { SheetRowReader } from './reader/stream-sheet';
export { XmlPullParser, type XmlEvent } from './xml/pull-parser';