import { describe, it, expect } from 'vitest';
import ExcelJS from 'exceljs';
import { zipSync, strToU8 } from 'fflate';
import { Workbook, ZipSecurityError, ZipUnsupportedError, type CellValue } from '../src/index';
import {
  StreamingWorkbook,
  BufferByteSource,
  readZipDirectory,
  streamEntryChunks,
  XmlPullParser,
  type ByteSource,
  type StreamRow,
} from '../src/stream';
import { makeMinimalXlsx, makeXlsxWithSheet } from './fixtures/make';

// ---------------------------------------------------------------------------
// helpers
// ---------------------------------------------------------------------------

/** Canonical form of a CellValue so eager and streaming reads compare exactly. */
function normalize(value: CellValue): unknown {
  if (value instanceof Date) return `date:${value.toISOString()}`;
  if (value === null || typeof value !== 'object') return value;
  if ('formula' in value) {
    return {
      formula: value.formula,
      result: value.result === undefined ? undefined : normalize(value.result),
    };
  }
  if ('error' in value) return { error: value.error };
  return value;
}

interface CellSnapshot {
  type: string;
  value: unknown;
}

function collectEager(file: Uint8Array, sheet: string): Map<string, CellSnapshot> {
  const ws = Workbook.read(file).getWorksheet(sheet)!;
  const out = new Map<string, CellSnapshot>();
  ws.eachRow((row) => {
    row.eachCell((cell) => out.set(cell.ref, { type: cell.type, value: normalize(cell.value) }));
  });
  return out;
}

async function collectStreaming(file: Uint8Array, sheet: string): Promise<Map<string, CellSnapshot>> {
  const wb = await StreamingWorkbook.open(file);
  try {
    const out = new Map<string, CellSnapshot>();
    for await (const row of wb.readRows(sheet)) {
      for (const cell of row.cells) {
        out.set(cell.ref, { type: cell.type, value: normalize(cell.value) });
      }
    }
    return out;
  } finally {
    await wb.close();
  }
}

function expectSameValues(actual: Map<string, CellSnapshot>, expected: Map<string, CellSnapshot>): void {
  expect(actual.size).toBe(expected.size);
  for (const [ref, snapshot] of expected) {
    const got = actual.get(ref);
    expect(got, `missing streamed cell ${ref}`).toBeDefined();
    expect(got!.value, `value of ${ref}`).toEqual(snapshot.value);
    expect(got!.type, `type of ${ref}`).toBe(snapshot.type);
  }
}

async function makeExcelJsFile(rows: number, cols: number): Promise<Buffer> {
  const wb = new ExcelJS.Workbook();
  const ws = wb.addWorksheet('Data');
  for (let r = 1; r <= rows; r += 1) {
    ws.getCell(r, 1).value = `row-${r}`;
    ws.getCell(r, 2).value = r;
    ws.getCell(r, 3).value = r * 1.5;
    ws.getCell(r, 4).value = new Date(Date.UTC(2020, 0, 1) + r * 86_400_000);
    ws.getCell(r, 4).numFmt = 'yyyy-mm-dd';
    ws.getCell(r, 5).value = r % 2 === 0;
    ws.getCell(r, 6).value = r * 0.01;
    ws.getCell(r, 7).value = `text ${r}`;
    ws.getCell(r, 8).value = r;
    if (cols > 8) ws.getCell(r, 9).value = 'tail';
  }
  // One formula with a cached result, cross-checked below.
  ws.getCell(2, 2).value = { formula: 'B1+B3', result: 4 };
  return Buffer.from(await wb.xlsx.writeBuffer());
}

/** Patch helpers reused from raw-zip.test.ts (independent of the code under test). */
function findEocd(zip: Uint8Array): number {
  const dv = new DataView(zip.buffer);
  let eocd = zip.length - 22;
  while (dv.getUint32(eocd, true) !== 0x06054b50) eocd -= 1;
  return eocd;
}

async function openError(zip: Uint8Array): Promise<unknown> {
  try {
    const wb = await StreamingWorkbook.open(zip);
    await wb.close();
    return null;
  } catch (error) {
    return error;
  }
}

async function drainRows(zip: Uint8Array): Promise<void> {
  const wb = await StreamingWorkbook.open(zip);
  try {
    for await (const row of wb.readRows('Sheet1')) {
      expect(row.cells.length).toBeGreaterThanOrEqual(0);
    }
  } finally {
    await wb.close();
  }
}

// ---------------------------------------------------------------------------
// streaming values
// ---------------------------------------------------------------------------

describe('stream — value parity with the eager reader', () => {
  it('streams the minimal fixture with identical values and types', async () => {
    const file = makeMinimalXlsx();
    expectSameValues(await collectStreaming(file, 'Sheet1'), collectEager(file, 'Sheet1'));
  });

  it('streams positional cells, empty <v>, booleans, errors and formulas like the eager reader', async () => {
    const xml =
      '<?xml version="1.0"?><worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">' +
      '<sheetData>' +
      '<row r="1">' +
      '<c r="A1" t="inlineStr"><is><t>a &amp; b</t></is></c>' +
      '<c r="B1"><v></v></c>' +
      '<c r="C1" t="b"><v>1</v></c>' +
      '<c r="D1" t="e"><v>#DIV/0!</v></c>' +
      '<c r="E1" t="str"><v>plain</v></c>' +
      '</row>' +
      '<row r="3"/>' +
      '<row><c><v>1</v></c><c><v>2</v></c></row>' +
      '<row r="5"><c r="B5"><f>SUM(A1:B1)</f><v>3</v></c></row>' +
      '</sheetData></worksheet>';
    const file = makeXlsxWithSheet(xml);
    expectSameValues(await collectStreaming(file, 'Sheet1'), collectEager(file, 'Sheet1'));

    const streamed = await collectStreaming(file, 'Sheet1');
    expect(streamed.get('A1')).toEqual({ type: 'String', value: 'a & b' });
    expect(streamed.get('B1')).toEqual({ type: 'Null', value: null });
    expect(streamed.get('C1')).toEqual({ type: 'Boolean', value: true });
    expect(streamed.get('D1')).toEqual({ type: 'Error', value: { error: '#DIV/0!' } });
    expect(streamed.get('E1')).toEqual({ type: 'String', value: 'plain' });
    expect(streamed.get('B5')).toEqual({ type: 'Formula', value: { formula: 'SUM(A1:B1)', result: 3 } });
  });

  it('streams a real ExcelJS file cell-for-cell (strings, numbers, dates, boolean, formula)', async () => {
    const file = await makeExcelJsFile(200, 8);
    expectSameValues(await collectStreaming(file, 'Data'), collectEager(file, 'Data'));

    const streamed = await collectStreaming(file, 'Data');
    expect(streamed.get('B2')!.type).toBe('Formula');
    expect(streamed.get('B2')!.value).toEqual({ formula: 'B1+B3', result: 4 });
    expect(String(streamed.get('D5')!.value)).toMatch(/^date:/);
    expect(streamed.get('E4')!.value).toBe(true);
  });

  it('parses rich-text shared strings identically to the eager reader', async () => {
    const wb = new ExcelJS.Workbook();
    const ws = wb.addWorksheet('S');
    ws.getCell('A1').value = { richText: [{ text: 'ri' }, { font: { bold: true }, text: 'ch' }] };
    ws.getCell('A2').value = 'plain & <special>';
    const file = Buffer.from(await wb.xlsx.writeBuffer());

    expectSameValues(await collectStreaming(file, 'S'), collectEager(file, 'S'));
    const streamed = await collectStreaming(file, 'S');
    // Cells whose string contains runs read back as `{ richText }` (ExcelJS shape).
    expect(streamed.get('A1')!.value).toEqual({
      richText: [{ text: 'ri' }, { font: { bold: true }, text: 'ch' }],
    });
    expect(streamed.get('A1')!.type).toBe('String');
    // Plain strings stay plain.
    expect(streamed.get('A2')!.value).toBe('plain & <special>');
  });

  it('streams a stored (method 0) archive', async () => {
    const xml =
      '<?xml version="1.0"?><worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">' +
      '<sheetData><row r="1"><c r="A1"><v>7</v></c><c r="B1" t="inlineStr"><is><t>x</t></is></c></row></sheetData>' +
      '</worksheet>';
    const file = makeXlsxWithSheet(xml, { level: 0 });
    expectSameValues(await collectStreaming(file, 'Sheet1'), collectEager(file, 'Sheet1'));
  });

  it('exposes the sheet list and rejects unknown sheets', async () => {
    const file = await makeExcelJsFile(3, 8);
    const wb = await StreamingWorkbook.open(file);
    try {
      expect(wb.sheetNames).toEqual(Workbook.read(file).sheetNames);
      expect(wb.getSheet('Data')!.name).toBe('Data');
      expect(wb.getSheet(1)!.name).toBe('Data');
      expect(wb.getSheet('Nope')).toBeUndefined();
      expect(() => wb.readRows('Nope')).toThrow(/Nope/);
      // Re-iterating the same sheet starts a fresh pass.
      const first = wb.getSheet('Data')!;
      expect(await collect(first.readRows())).toBe(await collect(first.readRows()));
    } finally {
      await wb.close();
    }
  });

  it('Workbook.openStream is a lazy alias for StreamingWorkbook.open', async () => {
    const file = makeMinimalXlsx();
    const wb = await Workbook.openStream(file);
    const rows: Array<{ number: number; cells: Array<{ ref: string; value: CellValue }> }> = [];
    for await (const row of wb.readRows('Sheet1')) rows.push(row);
    expect(rows).toHaveLength(2);
    expect(rows[0]!.cells[0]).toMatchObject({ ref: 'A1', value: 'Hello' });
    expect(rows[0]!.cells[1]).toMatchObject({ ref: 'B1', value: 42 });
    await wb.close();
  });

  it('streams directly from a Node file path (the archive is never fully loaded)', async () => {
    const { writeFile, unlink } = await import('node:fs/promises');
    const path = `${process.env.TEMP ?? process.env.TMP ?? '.'}/sheet-lite-stream-${Date.now()}.xlsx`;
    await writeFile(path, makeMinimalXlsx());
    try {
      const wb = await Workbook.openStream(path);
      try {
        const rows: Array<{ number: number }> = [];
        for await (const row of wb.readRows('Sheet1')) rows.push(row);
        expect(rows.map((r) => r.number)).toEqual([1, 2]);
        expect(wb.sheetNames).toEqual(['Sheet1']);
      } finally {
        await wb.close();
      }
    } finally {
      await unlink(path);
    }
  });
});

async function collect(gen: AsyncGenerator<StreamRow>): Promise<number> {
  let count = 0;
  for await (const _row of gen) count += 1;
  return count;
}

// ---------------------------------------------------------------------------
// streaming hardening (same rejections as the eager path)
// ---------------------------------------------------------------------------

describe('stream — hardening via the streaming path', () => {
  it('rejects zip64 markers at open()', async () => {
    const base = makeMinimalXlsx();
    const eocd = findEocd(base);
    const totalEntries = new DataView(base.buffer).getUint16(eocd + 10, true);

    const patched = new Uint8Array(base);
    new DataView(patched.buffer).setUint16(eocd + 10, 0xffff, true); // total entries -> zip64 marker
    const error = await openError(patched);
    expect(error).toBeInstanceOf(ZipUnsupportedError);
    expect((error as Error).message).toMatch(/zip64/);
    expect(totalEntries).toBeGreaterThan(0);
  });

  it('rejects a zip64 EOCD locator at open()', async () => {
    const base = makeMinimalXlsx();
    const eocd = findEocd(base);
    const body = base.subarray(0, eocd);
    const patched = new Uint8Array(body.length + 20 + 22);
    patched.set(body, 0);
    new DataView(patched.buffer).setUint32(body.length, 0x07064b50, true); // zip64 locator
    patched.set(base.subarray(eocd), body.length + 20);
    const error = await openError(patched);
    expect(error).toBeInstanceOf(ZipUnsupportedError);
    expect((error as Error).message).toMatch(/locator/);
  });

  it('handles data-descriptor (general purpose flag bit 3) entries like the eager reader', async () => {
    // `raw.ts` supports bit-3 entries (the central directory carries the real
    // CRC/sizes; the trailing descriptor is not part of `compressedSize`), so
    // streaming must accept them too — and agree on the values.
    const base = makeMinimalXlsx();
    const dv = new DataView(base.buffer);
    const eocd = findEocd(base);
    let pos = dv.getUint32(eocd + 16, true);
    const total = dv.getUint16(eocd + 10, true);
    for (let i = 0; i < total; i += 1) {
      const nameLen = dv.getUint16(pos + 28, true);
      const extraLen = dv.getUint16(pos + 30, true);
      const commentLen = dv.getUint16(pos + 32, true);
      const localOffset = dv.getUint32(pos + 42, true);
      dv.setUint16(pos + 8, dv.getUint16(pos + 8, true) | 0x0008, true); // central flags
      dv.setUint16(localOffset + 6, dv.getUint16(localOffset + 6, true) | 0x0008, true); // local flags
      pos += 46 + nameLen + extraLen + commentLen;
    }
    expectSameValues(await collectStreaming(base, 'Sheet1'), collectEager(base, 'Sheet1'));
  });

  it('rejects encrypted entries and traversal names at open()', async () => {
    const base = makeMinimalXlsx();
    const dv = new DataView(base.buffer);
    const firstRecord = dv.getUint32(findEocd(base) + 16, true);

    const encrypted = new Uint8Array(base);
    new DataView(encrypted.buffer).setUint16(firstRecord + 8, 0x0001, true);
    const encryptedError = await openError(encrypted);
    expect(encryptedError).toBeInstanceOf(ZipUnsupportedError);
    expect((encryptedError as Error).message).toMatch(/encrypt/i);

    const traversal = zipSync({ '../evil.txt': strToU8('x') });
    const traversalError = await openError(traversal);
    expect(traversalError).toBeInstanceOf(ZipSecurityError);
    expect((traversalError as Error).message).toMatch(/Unsafe entry rejected/);
  });

  it('enforces the total decompressed-size limit at open()', async () => {
    const zip = zipSync({ 'a.txt': strToU8('hello world') });
    let error: unknown;
    try {
      await StreamingWorkbook.open(zip, { maxTotalUncompressed: 3 });
    } catch (caught) {
      error = caught;
    }
    expect(error).toBeInstanceOf(ZipSecurityError);
    expect((error as Error).message).toMatch(/Total decompressed size exceeds limit/);
  });

  it('verifies CRC-32 while streaming and rejects corrupted entries', async () => {
    const xml =
      '<?xml version="1.0"?><worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">' +
      '<sheetData><row r="1"><c r="A1"><v>1</v></c><c r="B1"><v>2</v></c><c r="C1"><v>3</v></c></row></sheetData>' +
      '</worksheet>';
    const zip = makeXlsxWithSheet(xml, { level: 0 }); // stored: flip a data byte
    const dv = new DataView(zip.buffer);
    // Walk the central directory independently to find the sheet entry's local header.
    const eocd = findEocd(zip);
    let pos = dv.getUint32(eocd + 16, true);
    const total = dv.getUint16(eocd + 10, true);
    let localOffset = -1;
    for (let i = 0; i < total; i += 1) {
      const nameLen = dv.getUint16(pos + 28, true);
      const extraLen = dv.getUint16(pos + 30, true);
      const commentLen = dv.getUint16(pos + 32, true);
      const name = Buffer.from(zip.subarray(pos + 46, pos + 46 + nameLen)).toString('utf8');
      if (name === 'xl/worksheets/sheet1.xml') localOffset = dv.getUint32(pos + 42, true);
      pos += 46 + nameLen + extraLen + commentLen;
    }
    expect(localOffset).toBeGreaterThan(0);
    const nameLen = dv.getUint16(localOffset + 26, true);
    dv.setUint8(localOffset + 30 + nameLen, 88); // first data byte -> 'X'
    await expect(drainRows(zip)).rejects.toThrow(/CRC/);
  });

  it('rejects an archive that is not a ZIP', async () => {
    const error = await openError(strToU8('this is definitely not a zip file'));
    expect(error).toBeInstanceOf(ZipSecurityError);
  });
});

// ---------------------------------------------------------------------------
// memory ceiling
// ---------------------------------------------------------------------------

describe('stream — incremental decompression', () => {
  it('keeps the incremental parser buffer bounded', () => {
    const parser = new XmlPullParser();
    const row = '<row r="1"><c r="A1"><v>123</v></c><c r="B1" t="inlineStr"><is><t>x &amp; y</t></is></c></row>';
    let maxBuffered = 0;
    for (let i = 0; i < 100_000; i += 1) {
      parser.push(row);
      for (;;) {
        if (parser.next() === null) break;
      }
      if (parser.buffered > maxBuffered) maxBuffered = parser.buffered;
    }
    parser.end();
    for (;;) {
      if (parser.next() === null) break;
    }
    // 100k rows fed through one parser: the buffer must never grow with input.
    expect(maxBuffered).toBeLessThan(4096);
  });

  it('yields the first rows long before the whole entry has been read', async () => {
    const file = await makeExcelJsFile(15_000, 8);
    const inner = new BufferByteSource(file);

    class CountingSource implements ByteSource {
      bytesRead = 0;
      constructor(private readonly delegate: ByteSource) {}
      get size(): number {
        return this.delegate.size;
      }
      read(offset: number, length: number): Uint8Array {
        const out = this.delegate.read(offset, length);
        if (out instanceof Promise) throw new Error('test source must be synchronous');
        this.bytesRead += out.length;
        return out;
      }
    }

    const source = new CountingSource(inner);
    const dir = await readZipDirectory(source);
    const entry = dir.byName.get('xl/worksheets/sheet1.xml')!;
    expect(entry.compressedSize).toBeGreaterThan(128 * 1024); // the fixture is big enough to matter

    // Fresh counter: only the bytes read while streaming the entry count.
    const streaming = new CountingSource(inner);
    const iterator = streamEntryChunks(streaming, entry);
    const first = await iterator.next();
    expect(first.done).toBe(false);
    expect(first.value!.length).toBeGreaterThan(0);
    // A lazy reader: output flows after ~one 64 KiB chunk of input, not after
    // the whole entry — this is what keeps heap bounded for any file size.
    expect(streaming.bytesRead).toBeGreaterThan(0);
    expect(streaming.bytesRead).toBeLessThan(entry.compressedSize);
    expect(streaming.bytesRead).toBeLessThanOrEqual(2 * 64 * 1024);
    await iterator.return(undefined);
  });
});

describe('stream — memory ceiling', () => {
  const ROWS = 15_000;
  const COLS = 8;

  /** Best-effort full collection (available because vitest runs with --expose-gc). */
  async function settle(): Promise<void> {
    const gc = (globalThis as { gc?: () => void }).gc;
    for (let i = 0; i < 3 && gc; i += 1) {
      gc();
      await new Promise((resolve) => setTimeout(resolve, 25));
    }
  }

  it('reads a large sheet with a fraction of the eager reader\'s heap', { timeout: 180_000 }, async () => {
    const file = await makeExcelJsFile(ROWS, COLS);
    await settle();
    const startHeap = process.memoryUsage().heapUsed;
    const gc = (globalThis as { gc?: () => void }).gc;
    expect(typeof gc).toBe('function'); // vitest runs with --expose-gc

    // (a) eager: load the whole workbook and keep it alive. Reading its values
    // is covered by the parity tests above; here only memory matters.
    let eager: Workbook | null = Workbook.read(file);
    const ws = eager.getWorksheet('Data')!;
    expect(ws.getCell('A1').value).toBe('row-1');
    const eagerPeak = process.memoryUsage().heapUsed; // workbook retained => live memory
    const eagerDelta = eagerPeak - startHeap;
    eager = null;
    await settle();

    // (b) streaming: same rows, one at a time. Garbage is collected periodically
    // so the high-water mark reflects the reader's live memory, not short-lived
    // per-row allocations.
    const streamBaseline = process.memoryUsage().heapUsed;
    const streamWb = await StreamingWorkbook.open(file);
    let streamCells = 0;
    let streamPeak = 0;
    let sinceGc = 0;
    try {
      for await (const row of streamWb.readRows('Data')) {
        streamCells += row.cells.length;
        sinceGc += row.cells.length;
        if (sinceGc >= 4096) {
          sinceGc = 0;
          gc!();
          const used = process.memoryUsage().heapUsed - streamBaseline;
          if (used > streamPeak) streamPeak = used;
        }
      }
    } finally {
      await streamWb.close();
    }
    gc!();
    const tail = process.memoryUsage().heapUsed - streamBaseline;
    if (tail > streamPeak) streamPeak = tail;
    const streamDelta = Math.max(0, streamPeak);

    // Every cell is streamed exactly once (8 per row).
    expect(streamCells).toBe(ROWS * COLS);

    // Report the numbers so they stay visible in CI output.
    const mib = (bytes: number): string => `${(bytes / 1024 / 1024).toFixed(1)} MB`;
    console.log(
      `[memory] eager peak ${mib(eagerDelta)} vs stream peak ${mib(streamDelta)} ` +
        `(${ROWS} rows x ${COLS} cols)`,
    );

    // Ceiling justified by the design: the streaming reader only ever holds a
    // ~64 KiB compressed chunk, a few decompressed chunks, one row and the
    // (few-hundred-kB) shared strings. 32 MB is a very generous bound.
    expect(streamDelta).toBeLessThan(32 * 1024 * 1024);
    // The eager reader materializes the sheet XML + a DOM + a cell map.
    expect(eagerDelta).toBeGreaterThan(40 * 1024 * 1024);
    expect(streamDelta).toBeLessThan(eagerDelta / 4);
  });
});