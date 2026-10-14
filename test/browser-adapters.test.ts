import { describe, it, expect } from 'vitest';
import { Workbook, SheetLiteError } from '../src/index';
import { makeMinimalXlsx } from './fixtures/make';

describe('runtime — browser / edge adapters (Blob)', () => {
  it('reads a workbook from a Blob', async () => {
    const blob = new Blob([new Uint8Array(makeMinimalXlsx())]);
    const wb = await Workbook.readBlob(blob);
    expect(wb.sheetNames).toEqual(['Sheet1']);
    expect(wb.getWorksheet('Sheet1')!.getCell('B1').value).toBe(42);
  });

  it('accepts ArrayBuffer-backed Blobs (File uploads) and other parts', async () => {
    const bytes = makeMinimalXlsx();
    const buffer = new Uint8Array(bytes.length);
    buffer.set(bytes);
    const blob = new Blob([buffer.buffer]);
    const wb = await Workbook.readBlob(blob);
    expect(wb.getWorksheet('Sheet1')!.getCell('A1').value).toBe('Hello');
  });

  it('serializes to a Blob whose bytes equal toBuffer()', () => {
    const wb = Workbook.read(makeMinimalXlsx());
    wb.getWorksheet('Sheet1')!.getCell('A1').value = 'Hola';
    const blob = wb.toBlob();
    expect(blob).toBeInstanceOf(Blob);
    expect(blob.size).toBe(wb.toBuffer().length);
  });

  it('round-trips through Blob: toBlob -> readBlob', async () => {
    const wb = Workbook.read(makeMinimalXlsx());
    wb.getWorksheet('Sheet1')!.getCell('B2').value = 7.5;
    const reread = await Workbook.readBlob(wb.toBlob());
    expect(reread.getWorksheet('Sheet1')!.getCell('B2').value).toBe(7.5);
    expect(reread.getWorksheet('Sheet1')!.getCell('A1').value).toBe('Hello');
  });

  it('download() requires a browser runtime (clear error outside the browser)', () => {
    const wb = Workbook.read(makeMinimalXlsx());
    expect(() => wb.download('out.xlsx')).toThrow(SheetLiteError);
    try {
      wb.download('out.xlsx');
    } catch (error) {
      expect((error as SheetLiteError).code).toBe('ENV_UNSUPPORTED');
    }
  });
});
