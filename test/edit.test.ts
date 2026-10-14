import { describe, it, expect } from 'vitest';
import { Workbook } from '../src/index';
import { makeMinimalXlsx } from './fixtures/make';

const eq = (a: Uint8Array, b: Uint8Array): boolean => Buffer.from(a).equals(Buffer.from(b));

describe('v0.2 — editing without breaking the file', () => {
  it('changes one cell and keeps every other part byte-identical', () => {
    const original = makeMinimalXlsx();
    const wb = Workbook.read(original);
    wb.getWorksheet('Sheet1')!.getCell('B1').value = 100;

    const out = wb.toBuffer();
    expect(eq(out, original)).toBe(false);

    const reread = Workbook.read(out);
    const ws = reread.getWorksheet('Sheet1')!;
    expect(ws.getCell('B1').value).toBe(100);
    expect(ws.getCell('A1').value).toBe('Hello');
    expect(ws.getCell('A2').value).toBe('World');
    expect(ws.getCell('B2').value).toBeCloseTo(3.14);

    const before = wb.parts();
    const after = reread.parts();
    for (const name of before.names()) {
      if (name === 'xl/worksheets/sheet1.xml') continue;
      expect(eq(before.get(name)!, after.get(name)!), `part changed: ${name}`).toBe(true);
    }
  });

  it('appends a new row that can be read back', () => {
    const wb = Workbook.read(makeMinimalXlsx());
    wb.getWorksheet('Sheet1')!.addRow(['nuevo', 7]);

    const reread = Workbook.read(wb.toBuffer());
    const ws = reread.getWorksheet('Sheet1')!;
    expect(ws.getCell('A3').value).toBe('nuevo');
    expect(ws.getCell('B3').value).toBe(7);
    // original rows still intact
    expect(ws.getCell('A1').value).toBe('Hello');
    expect(ws.getCell('B2').value).toBeCloseTo(3.14);
  });

  it('writes the right XML for each value type', () => {
    const wb = Workbook.read(makeMinimalXlsx());
    const ws = wb.getWorksheet('Sheet1')!;
    ws.getCell('A1').value = true;
    ws.getCell('A2').value = null;

    const reread = Workbook.read(wb.toBuffer());
    const ws2 = reread.getWorksheet('Sheet1')!;
    expect(ws2.getCell('A1').value).toBe(true);
    expect(ws2.getCell('A1').type).toBe('Boolean');
    expect(ws2.getCell('A2').value).toBeNull();
  });

  it('preserves a formula on an untouched cell and can set one', () => {
    const wb = Workbook.read(makeMinimalXlsx());
    const ws = wb.getWorksheet('Sheet1')!;
    ws.getCell('B3').value = { formula: 'SUM(B1:B2)', result: 45.14 };

    const reread = Workbook.read(wb.toBuffer());
    const cell = reread.getWorksheet('Sheet1')!.getCell('B3');
    expect(cell.type).toBe('Formula');
    expect(cell.value).toEqual({ formula: 'SUM(B1:B2)', result: 45.14 });
  });

  it('no-op still returns the original bytes', () => {
    const original = makeMinimalXlsx();
    expect(eq(Workbook.read(original).toBuffer(), original)).toBe(true);
  });
});
