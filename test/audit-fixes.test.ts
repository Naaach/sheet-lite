import { describe, it, expect } from 'vitest';
import ExcelJS from 'exceljs';
import * as XLSX from 'xlsx';
import { strFromU8 } from 'fflate';
import { Workbook, serialToDate, dateToSerial, isDateFormatCode, decodeEntities } from '../src/index';
import { WorkbookWriter } from '../src/write';
import { makeMinimalXlsx, makeXlsxWithSheet } from './fixtures/make';

describe('audit fixes', () => {
  it('1900 leap-year: serial 1 is 1900-01-01 and round-trips', () => {
    expect(serialToDate(1, false).toISOString()).toBe('1900-01-01T00:00:00.000Z');
    expect(dateToSerial(new Date(Date.UTC(1900, 0, 1)), false)).toBe(1);
    expect(serialToDate(59, false).toISOString()).toBe('1900-02-28T00:00:00.000Z');
    expect(serialToDate(61, false).toISOString()).toBe('1900-03-01T00:00:00.000Z');
  });

  it('detects date formats correctly (locale prefixes vs accounting)', () => {
    expect(isDateFormatCode('[$-409]d-mmm-yy')).toBe(true);
    expect(isDateFormatCode('0.00;[Red]-0.00')).toBe(false);
    expect(isDateFormatCode('#,##0.00')).toBe(false);
    expect(isDateFormatCode('"due" dd')).toBe(true);
  });

  it('does not throw on out-of-range entities', () => {
    expect(decodeEntities('&#x110000;')).toBe('\uFFFD');
    expect(decodeEntities('&amp;&#65;')).toBe('&A');
  });

  it('writes mergeCells as a sibling AFTER </sheetData> (valid XML)', async () => {
    const wb = new ExcelJS.Workbook();
    wb.addWorksheet('S').getCell('A1').value = 'hi';
    const file = Buffer.from(await wb.xlsx.writeBuffer());

    const doc = Workbook.read(file);
    doc.getWorksheet('S')!.mergeCells('A1:C1');
    const out = doc.toBuffer();

    const xml = strFromU8(Workbook.read(out).parts().get('xl/worksheets/sheet1.xml')!);
    const sheetDataEnd = xml.indexOf('</sheetData>');
    const mergeIdx = xml.indexOf('<mergeCells');
    expect(sheetDataEnd).toBeGreaterThanOrEqual(0);
    expect(mergeIdx).toBeGreaterThan(sheetDataEnd);

    expect(() => XLSX.read(out, { type: 'buffer' })).not.toThrow();
    expect(Workbook.read(out).getWorksheet('S')!.getCell('A1').value).toBe('hi');
  });

  it('reads cells that omit the r attribute (positional)', () => {
    const xml =
      '<?xml version="1.0"?><worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">' +
      '<sheetData><row>' +
      '<c t="inlineStr"><is><t>a</t></is></c>' +
      '<c><v>1</v></c>' +
      '<c><v>2</v></c>' +
      '</row></sheetData></worksheet>';
    const ws = Workbook.read(makeXlsxWithSheet(xml)).getWorksheet('Sheet1')!;
    expect(ws.getCell('A1').value).toBe('a');
    expect(ws.getCell('B1').value).toBe(1);
    expect(ws.getCell('C1').value).toBe(2);
  });

  it('treats an empty <v> as null, not 0', () => {
    const xml =
      '<?xml version="1.0"?><worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">' +
      '<sheetData><row r="1"><c r="A1"><v></v></c></row></sheetData></worksheet>';
    const ws = Workbook.read(makeXlsxWithSheet(xml)).getWorksheet('Sheet1')!;
    expect(ws.getCell('A1').value).toBeNull();
  });

  it('round-trips a formula whose cached result is a string, boolean or error (eager writer)', () => {
    const wb = Workbook.read(makeMinimalXlsx());
    const ws = wb.getWorksheet('Sheet1')!;
    ws.getCell('A1').value = { formula: 'CONCAT("a","b")', result: 'ab' };
    ws.getCell('B1').value = { formula: 'B2>0', result: true };
    ws.getCell('A2').value = { formula: 'NA()', result: { error: '#N/A' } };

    const reread = Workbook.read(wb.toBuffer()).getWorksheet('Sheet1')!;
    // Before the fix these read back as `result: NaN` (or lost the error).
    expect(reread.getCell('A1').value).toEqual({ formula: 'CONCAT("a","b")', result: 'ab' });
    expect(reread.getCell('B1').value).toEqual({ formula: 'B2>0', result: true });
    expect(reread.getCell('A2').value).toEqual({ formula: 'NA()', result: { error: '#N/A' } });
  });

  it('round-trips formula results of every type through the incremental writer', async () => {
    const wb = await WorkbookWriter.create();
    wb.addWorksheet('S').addRow([
      { formula: 'CONCAT("a","b")', result: 'ab' },
      { formula: 'B1>0', result: true },
      { formula: 'NA()', result: { error: '#N/A' } },
      { formula: 'SUM(A1:A2)', result: 7 },
    ]);

    const reread = Workbook.read(await wb.toBuffer()).getWorksheet('S')!;
    expect(reread.getCell('A1').value).toEqual({ formula: 'CONCAT("a","b")', result: 'ab' });
    expect(reread.getCell('B1').value).toEqual({ formula: 'B1>0', result: true });
    expect(reread.getCell('C1').value).toEqual({ formula: 'NA()', result: { error: '#N/A' } });
    expect(reread.getCell('D1').value).toEqual({ formula: 'SUM(A1:A2)', result: 7 });
  });

  it('reopens a self-closing styles container instead of emitting invalid XML', () => {
    const styles =
      '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
      '<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">' +
      '<fonts/><fills/><borders/><cellStyleXfs/><cellXfs/>' +
      '</styleSheet>';
    const wb = Workbook.read(
      makeXlsxWithSheet(
        '<?xml version="1.0"?><worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">' +
          '<sheetData><row r="1"><c r="A1"><v>1</v></c></row></sheetData></worksheet>',
        { styles },
      ),
    );
    wb.getWorksheet('Sheet1')!.getCell('A1').font = { bold: true };
    const out = wb.toBuffer();

    const stylesXml = strFromU8(Workbook.read(out).parts().get('xl/styles.xml')!);
    // The self-closing containers are reopened, not left as `<fonts/ count=…>`.
    expect(stylesXml).toContain('<fonts count="1"><font><b/></font></fonts>');
    expect(stylesXml).toContain('<cellXfs count="1">');
    expect(stylesXml).toContain('</cellXfs>');
    expect(stylesXml).not.toContain('<fonts/');
    expect(stylesXml).not.toContain('<cellXfs/');
    // The written style is readable back.
    expect(Workbook.read(out).getWorksheet('Sheet1')!.getCell('A1').font).toEqual({ bold: true });
  });

  it('reopens a self-closing mergeCells container when adding a merge', () => {
    const xml =
      '<?xml version="1.0"?><worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">' +
      '<sheetData><row r="1"><c r="A1"><v>1</v></c></row></sheetData>' +
      '<mergeCells/>' +
      '</worksheet>';
    const wb = Workbook.read(makeXlsxWithSheet(xml));
    wb.getWorksheet('Sheet1')!.mergeCells('A1:C1');
    const out = wb.toBuffer();

    const sheetXml = strFromU8(Workbook.read(out).parts().get('xl/worksheets/sheet1.xml')!);
    expect(sheetXml).toContain('<mergeCells count="1"><mergeCell ref="A1:C1"/></mergeCells>');
    expect(sheetXml).not.toContain('<mergeCells/');
    expect(() => XLSX.read(out, { type: 'buffer' })).not.toThrow();
  });
});
