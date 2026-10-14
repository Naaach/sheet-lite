import { describe, it, expect } from 'vitest';
import ExcelJS from 'exceljs';
import * as XLSX from 'xlsx';
import { strFromU8 } from 'fflate';
import { Workbook } from '../src/index';

/** Build a realistic workbook with ExcelJS: values, date, style, formula,
 *  data validation, conditional formatting and merged cells. */
async function buildRealFile(): Promise<Buffer> {
  const wb = new ExcelJS.Workbook();
  const ws = wb.addWorksheet('Sheet1');

  ws.getCell('A1').value = 'Hello';
  ws.getCell('A1').font = { bold: true };
  ws.getCell('A2').value = 42;
  ws.getCell('A3').value = 1.5;
  ws.getCell('A4').value = new Date(Date.UTC(2026, 0, 15));
  ws.getCell('A4').numFmt = 'yyyy-mm-dd';
  ws.getCell('A5').value = true;
  ws.getCell('B1').value = { formula: 'SUM(A2:A3)', result: 43.5 };

  ws.getCell('C1').dataValidation = {
    type: 'list',
    allowBlank: true,
    formulae: ['"red,green,blue"'],
  };
  ws.addConditionalFormatting({
    ref: 'A2:A3',
    rules: [
      {
        type: 'cellIs',
        operator: 'greaterThan',
        formulae: ['10'],
        priority: 1,
        style: { fill: { type: 'pattern', pattern: 'solid', bgColor: { argb: 'FFFFC7CE' } } },
      },
    ],
  });
  ws.mergeCells('D1:E1');
  ws.getCell('D1').value = 'merged';

  const buffer = await wb.xlsx.writeBuffer();
  return Buffer.from(buffer);
}

describe('v0.2 — compatibility with real files and SheetJS', () => {
  it('reads the same values SheetJS reads', async () => {
    const file = await buildRealFile();
    const sheetjs = XLSX.read(file, { type: 'buffer', cellDates: true });
    const sjs = sheetjs.Sheets.Sheet1!;

    const ws = Workbook.read(file).getWorksheet('Sheet1')!;

    expect(ws.getCell('A1').value).toBe(sjs.A1.v);
    expect(ws.getCell('A2').value).toBe(sjs.A2.v);
    expect(ws.getCell('A3').value).toBe(sjs.A3.v);
    expect(ws.getCell('A5').value).toBe(sjs.A5.v);
    // Dates round-trip exactly to the value we wrote.
    expect((ws.getCell('A4').value as Date).toISOString()).toBe('2026-01-15T00:00:00.000Z');
  });

  it('reads formulas and their cached result', async () => {
    const ws = Workbook.read(await buildRealFile()).getWorksheet('Sheet1')!;
    expect(ws.getCell('B1').value).toEqual({ formula: 'SUM(A2:A3)', result: 43.5 });
  });

  it('preserves data validations, conditional formatting and merges when editing', async () => {
    const file = await buildRealFile();
    const wb = Workbook.read(file);
    wb.getWorksheet('Sheet1')!.getCell('A2').value = 100;

    const out = wb.toBuffer();
    const sheetXml = strFromU8(wb.parts().get('xl/worksheets/sheet1.xml')!);
    const outXml = strFromU8(Workbook.read(out).parts().get('xl/worksheets/sheet1.xml')!);

    expect(sheetXml).toContain('<dataValidations');
    expect(sheetXml).toContain('<conditionalFormatting');
    expect(sheetXml).toContain('<mergeCells');
    expect(outXml).toContain('<dataValidations');
    expect(outXml).toContain('<conditionalFormatting');
    expect(outXml).toContain('<mergeCells');

    // edited value persisted, and SheetJS can still parse the result
    const reread = Workbook.read(out).getWorksheet('Sheet1')!;
    expect(reread.getCell('A2').value).toBe(100);
    expect(reread.getCell('A1').value).toBe('Hello');
    expect(() => XLSX.read(out, { type: 'buffer' })).not.toThrow();
  });
});
