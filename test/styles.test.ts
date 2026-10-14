import { describe, it, expect } from 'vitest';
import ExcelJS from 'exceljs';
import * as XLSX from 'xlsx';
import { strFromU8 } from 'fflate';
import { Workbook } from '../src/index';

async function realFile(): Promise<Buffer> {
  const wb = new ExcelJS.Workbook();
  const ws = wb.addWorksheet('S');
  ws.getCell('A1').value = 'x';
  ws.getCell('A2').value = 1234.5;
  ws.getCell('A3').value = 0.42;
  return Buffer.from(await wb.xlsx.writeBuffer());
}

const sheetXmlOf = (buffer: Uint8Array): string =>
  strFromU8(Workbook.read(buffer).parts().get('xl/worksheets/sheet1.xml')!);

describe('v0.3 — styles and merges', () => {
  it('applies and re-reads a number format', async () => {
    const wb = Workbook.read(await realFile());
    wb.getWorksheet('S')!.getCell('A2').numFmt = '#,##0.00';
    const out = wb.toBuffer();

    const reread = Workbook.read(out).getWorksheet('S')!;
    expect(reread.getCell('A2').numFmt).toBe('#,##0.00');
    expect(reread.getCell('A2').value).toBe(1234.5);
    expect(() => XLSX.read(out, { type: 'buffer' })).not.toThrow();
  });

  it('applies a bold font and a solid fill', async () => {
    const wb = Workbook.read(await realFile());
    const cell = wb.getWorksheet('S')!.getCell('A1');
    cell.font = { bold: true };
    cell.fill = { type: 'pattern', pattern: 'solid', fgColor: 'FFFF0000' };
    const out = wb.toBuffer();

    const reread = Workbook.read(out).getWorksheet('S')!;
    expect(reread.getCell('A1').font.bold).toBe(true);
    expect(reread.getCell('A1').fill).toMatchObject({ pattern: 'solid', fgColor: 'FFFF0000' });
    expect(() => XLSX.read(out, { type: 'buffer' })).not.toThrow();
  });

  it('creates a merge group', async () => {
    const wb = Workbook.read(await realFile());
    wb.getWorksheet('S')!.mergeCells('C1:D1');
    const out = wb.toBuffer();
    expect(sheetXmlOf(out)).toContain('<mergeCells count="1"><mergeCell ref="C1:D1"/></mergeCells>');
    expect(() => XLSX.read(out, { type: 'buffer' })).not.toThrow();
  });

  it('edits value and style together, leaving other parts intact', async () => {
    const file = await realFile();
    const wb = Workbook.read(file);
    const cell = wb.getWorksheet('S')!.getCell('A3');
    cell.value = 0.9;
    cell.numFmt = '0%';

    const out = wb.toBuffer();
    const before = Workbook.read(file).parts();
    const after = Workbook.read(out).parts();
    for (const name of before.names()) {
      if (name === 'xl/worksheets/sheet1.xml' || name === 'xl/styles.xml') continue;
      expect(Buffer.from(before.get(name)!).equals(Buffer.from(after.get(name)!)), name).toBe(true);
    }
    const reread = Workbook.read(out).getWorksheet('S')!;
    expect(reread.getCell('A3').value).toBe(0.9);
    expect(reread.getCell('A3').numFmt).toBe('0%');
  });
});
