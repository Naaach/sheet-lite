import { describe, it, expect } from 'vitest';
import ExcelJS from 'exceljs';
import * as XLSX from 'xlsx';
import { strFromU8 } from 'fflate';
import { Workbook, type CellValue, type RichTextValue } from '../src/index';
import { StreamingWorkbook } from '../src/stream';
import { makeMinimalXlsx, makeXlsxWithSheet } from './fixtures/make';

// ---------------------------------------------------------------------------
// helpers
// ---------------------------------------------------------------------------

const stylesOf = (buffer: Uint8Array): string =>
  strFromU8(Workbook.read(buffer).parts().get('xl/styles.xml')!);

const sheetOf = (buffer: Uint8Array): string =>
  strFromU8(Workbook.read(buffer).parts().get('xl/worksheets/sheet1.xml')!);

async function excelJsFile(setup: (ws: ExcelJS.Worksheet) => void): Promise<Buffer> {
  const wb = new ExcelJS.Workbook();
  const ws = wb.addWorksheet('S');
  setup(ws);
  return Buffer.from(await wb.xlsx.writeBuffer());
}

/** ExcelJS's `load` typings use a legacy global `Buffer` interface. */
type ExcelLoadInput = Parameters<ExcelJS.Workbook['xlsx']['load']>[0];

async function loadInExcelJs(data: Uint8Array): Promise<ExcelJS.Workbook> {
  const ej = new ExcelJS.Workbook();
  await ej.xlsx.load(data as unknown as ExcelLoadInput);
  return ej;
}

/** Eager and streaming reads of every cell, as canonical snapshots. */
function eagerValues(file: Uint8Array, sheet: string): Map<string, CellValue> {
  const out = new Map<string, CellValue>();
  Workbook.read(file).getWorksheet(sheet)!.eachRow((row) => {
    row.eachCell((cell) => out.set(cell.ref, cell.value));
  });
  return out;
}

async function streamValues(file: Uint8Array, sheet: string): Promise<Map<string, CellValue>> {
  const wb = await StreamingWorkbook.open(file);
  try {
    const out = new Map<string, CellValue>();
    for await (const row of wb.readRows(sheet)) {
      for (const cell of row.cells) out.set(cell.ref, cell.value);
    }
    return out;
  } finally {
    await wb.close();
  }
}

// ---------------------------------------------------------------------------
// borders
// ---------------------------------------------------------------------------

describe('styles2 — borders', () => {
  it('writes a border and reads the same shape back through our reader', () => {
    const wb = Workbook.read(makeMinimalXlsx());
    const cell = wb.getWorksheet('Sheet1')!.getCell('B1');
    cell.value = 'boxed';
    cell.border = {
      top: { style: 'thin', color: { argb: 'FF00B0F0' } },
      bottom: { style: 'double' },
      left: { style: 'dashed' },
      right: { style: 'medium' },
      diagonal: { style: 'hair' },
      diagonalUp: true,
      diagonalDown: true,
    };
    const out = wb.toBuffer();

    const styles = stylesOf(out);
    expect(styles).toContain('<borders count="2"');
    expect(styles).toContain('<top style="thin"><color rgb="FF00B0F0"/></top>');
    expect(styles).toContain('applyBorder="1"');
    // All five sides are emitted in schema order (exactly what Excel writes).
    expect(styles).toMatch(/<border diagonalUp="1" diagonalDown="1"><left style="dashed"\/><right style="medium"\/><top style="thin"><color rgb="FF00B0F0"\/><\/top><bottom style="double"\/><diagonal style="hair"\/><\/border>/);

    const reread = Workbook.read(out).getWorksheet('Sheet1')!;
    expect(reread.getCell('B1').border).toEqual({
      top: { style: 'thin', color: { argb: 'FF00B0F0' } },
      bottom: { style: 'double' },
      left: { style: 'dashed' },
      right: { style: 'medium' },
      diagonal: { style: 'hair' },
      diagonalUp: true,
      diagonalDown: true,
    });
    // Untouched cells have no border.
    expect(reread.getCell('A1').border).toEqual({});
    expect(() => XLSX.read(out, { type: 'buffer' })).not.toThrow();
  });

  it('reuses an existing identical border instead of appending a duplicate', async () => {
    const file = await excelJsFile((ws) => {
      ws.getCell('A1').border = { left: { style: 'thin', color: { argb: 'FFFF0000' } } };
    });
    const wb = Workbook.read(new Uint8Array(file));
    const before = /<borders count="(\d+)"/.exec(stylesOf(new Uint8Array(file)))![1];
    wb.getWorksheet('S')!.getCell('C3').border = { left: { style: 'thin', color: { argb: 'FFFF0000' } } };
    const out = wb.toBuffer();
    const after = /<borders count="(\d+)"/.exec(stylesOf(out))![1];
    expect(after).toBe(before); // deduped against the existing style
    expect(Workbook.read(out).getWorksheet('S')!.getCell('C3').border).toEqual({
      left: { style: 'thin', color: { argb: 'FFFF0000' } },
    });
  });

  it('dedups borders set on several cells in one write', () => {
    const wb = Workbook.read(makeMinimalXlsx());
    const ws = wb.getWorksheet('Sheet1')!;
    ws.getCell('A1').border = { top: { style: 'thin' } };
    ws.getCell('B2').border = { top: { style: 'thin' } };
    ws.getCell('C3').border = { bottom: { style: 'thick' } };
    const styles = stylesOf(wb.toBuffer());
    expect(styles).toContain('<borders count="3"'); // 1 base + 2 distinct borders
    expect(styles.match(/<border[ />]/g)!.length).toBe(3);
  });

  it('merges one new side into a cell that already has a border (never drops sides)', async () => {
    const file = await excelJsFile((ws) => {
      ws.getCell('B2').value = 'x';
      ws.getCell('B2').border = {
        left: { style: 'thin', color: { argb: 'FFFF0000' } },
        right: { style: 'thin', color: { argb: 'FFFF0000' } },
      };
    });
    const wb = Workbook.read(new Uint8Array(file));
    wb.getWorksheet('S')!.getCell('B2').border = { top: { style: 'double' } };
    const out = wb.toBuffer();
    const border = Workbook.read(out).getWorksheet('S')!.getCell('B2').border;
    expect(border).toEqual({
      left: { style: 'thin', color: { argb: 'FFFF0000' } },
      right: { style: 'thin', color: { argb: 'FFFF0000' } },
      top: { style: 'double' },
    });
  });

  it('empty border and alignment objects are a no-op (byte-identical output)', () => {
    const file = makeMinimalXlsx();
    const wb = Workbook.read(file);
    const ws = wb.getWorksheet('Sheet1')!;
    ws.getCell('A1').border = {};
    ws.getCell('A1').alignment = {};
    expect(wb.toBuffer()).toEqual(file);
  });

  it('cross-reads borders with ExcelJS (diagonalUp/Down land in the diagonal model)', async () => {
    const wb = Workbook.read(makeMinimalXlsx());
    const ws = wb.getWorksheet('Sheet1')!;
    ws.getCell('B1').border = {
      top: { style: 'thin', color: { argb: 'FF00B0F0' } },
      diagonal: { style: 'hair' },
      diagonalUp: true,
      diagonalDown: true,
    };
    const out = Buffer.from(wb.toBuffer());
    const ej = await loadInExcelJs(out);
    const border = ej.getWorksheet('Sheet1')!.getCell('B1').border;
    expect(border).toMatchObject({
      top: { style: 'thin', color: { argb: 'FF00B0F0' } },
    });
    expect(border.diagonal).toMatchObject({ style: 'hair', up: true, down: true });
  });

  it('reads borders written by ExcelJS back into the same shape', async () => {
    const file = await excelJsFile((ws) => {
      ws.getCell('A2').border = {
        top: { style: 'mediumDashed' },
        // ExcelJS carries diagonalUp inside its diagonal model; the file gets diagonalUp="1".
        diagonal: { style: 'thick', color: { argb: 'FF123456' }, up: true },
      };
    });
    const border = Workbook.read(new Uint8Array(file)).getWorksheet('S')!.getCell('A2').border;
    expect(border).toEqual({
      top: { style: 'mediumDashed' },
      diagonal: { style: 'thick', color: { argb: 'FF123456' } },
      diagonalUp: true,
    });
  });
});

// ---------------------------------------------------------------------------
// alignment
// ---------------------------------------------------------------------------

describe('styles2 — alignment', () => {
  it('writes alignment inside the xf and reads the same shape back', () => {
    const wb = Workbook.read(makeMinimalXlsx());
    const ws = wb.getWorksheet('Sheet1')!;
    ws.getCell('A1').alignment = {
      horizontal: 'center',
      vertical: 'top',
      wrapText: true,
      shrinkToFit: true,
      indent: 1,
      textRotation: 45,
      readingOrder: 2,
    };
    const out = wb.toBuffer();

    const styles = stylesOf(out);
    expect(styles).toContain('applyAlignment="1"');
    expect(styles).toContain('<alignment horizontal="center" vertical="top" wrapText="1" shrinkToFit="1" indent="1" textRotation="45" readingOrder="2"/>');

    const reread = Workbook.read(out).getWorksheet('Sheet1')!;
    expect(reread.getCell('A1').alignment).toEqual({
      horizontal: 'center',
      vertical: 'top',
      wrapText: true,
      shrinkToFit: true,
      indent: 1,
      textRotation: 45,
      readingOrder: 2,
    });
    expect(reread.getCell('B1').alignment).toEqual({});
    expect(() => XLSX.read(out, { type: 'buffer' })).not.toThrow();
  });

  it('dedups xfs that differ only by alignment and reuses identical alignments', () => {
    const wb = Workbook.read(makeMinimalXlsx());
    const ws = wb.getWorksheet('Sheet1')!;
    ws.getCell('A1').alignment = { horizontal: 'right' };
    ws.getCell('B1').alignment = { horizontal: 'right' }; // same -> reuse
    ws.getCell('A2').alignment = { horizontal: 'left', wrapText: true }; // different -> new
    const styles = stylesOf(wb.toBuffer());
    expect(styles).toContain('<cellXfs count="3"'); // base + 2 distinct alignment xfs
    expect(styles.match(/<alignment /g)!.length).toBe(2);
  });

  it('keeps a cell\'s alignment when an unrelated numFmt is patched onto it', () => {
    const wb = Workbook.read(makeMinimalXlsx());
    const ws = wb.getWorksheet('Sheet1')!;
    ws.getCell('A1').alignment = { horizontal: 'center', vertical: 'bottom' };
    const wb2 = Workbook.read(wb.toBuffer());
    wb2.getWorksheet('Sheet1')!.getCell('A1').numFmt = '0.00';
    const out = wb2.toBuffer();
    const reread = Workbook.read(out).getWorksheet('Sheet1')!;
    expect(reread.getCell('A1').numFmt).toBe('0.00');
    expect(reread.getCell('A1').alignment).toEqual({ horizontal: 'center', vertical: 'bottom' });
    const styles = stylesOf(out);
    // The numFmt xf carries the alignment forward instead of dropping it.
    expect(styles).toMatch(/applyNumberFormat="1"><alignment horizontal="center" vertical="bottom"\/><\/xf>/);
  });

  it('cross-reads alignment with ExcelJS (OOXML center -> ExcelJS middle)', async () => {
    const wb = Workbook.read(makeMinimalXlsx());
    wb.getWorksheet('Sheet1')!.getCell('B1').alignment = {
      horizontal: 'center',
      vertical: 'center',
      wrapText: true,
    };
    const out = Buffer.from(wb.toBuffer());
    const ej = await loadInExcelJs(out);
    const alignment = ej.getWorksheet('Sheet1')!.getCell('B1').alignment;
    expect(alignment).toMatchObject({ horizontal: 'center', vertical: 'middle', wrapText: true });
  });

  it('reads ExcelJS alignments back in the OOXML shape (middle -> center)', async () => {
    const file = await excelJsFile((ws) => {
      ws.getCell('A1').alignment = {
        horizontal: 'center',
        vertical: 'middle',
        wrapText: true,
        indent: 2,
        textRotation: 90,
        readingOrder: 'rtl',
      };
    });
    const alignment = Workbook.read(new Uint8Array(file)).getWorksheet('S')!.getCell('A1').alignment;
    expect(alignment).toEqual({
      horizontal: 'center',
      vertical: 'center',
      wrapText: true,
      indent: 2,
      textRotation: 90,
      readingOrder: 2,
    });
  });
});

// ---------------------------------------------------------------------------
// rich text
// ---------------------------------------------------------------------------

const RICH: RichTextValue = {
  richText: [
    { text: 'Hello ' },
    { text: 'bold', font: { bold: true, size: 14, color: 'FFFF0000', name: 'Arial' } },
    { text: ' & <tail>', font: { italic: true } },
  ],
};

describe('styles2 — rich text', () => {
  it('writes runs and reads the same richText shape back through our reader', () => {
    const wb = Workbook.read(makeMinimalXlsx());
    const ws = wb.getWorksheet('Sheet1')!;
    ws.getCell('A1').value = RICH;
    ws.getCell('A2').value = 'plain stays plain';
    const out = wb.toBuffer();

    const sheet = sheetOf(out);
    expect(sheet).toContain('<c r="A1" t="inlineStr"');
    expect(sheet).toContain('<r><rPr><b/><sz val="14"/><color rgb="FFFF0000"/><rFont val="Arial"/></rPr><t>bold</t></r>');
    expect(sheet).toContain('<r><rPr><i/></rPr><t xml:space="preserve"> &amp; &lt;tail&gt;</t></r>');
    expect(sheet).toContain('<c r="A2" t="inlineStr"><is><t>plain stays plain</t></is></c>');

    const reread = Workbook.read(out).getWorksheet('Sheet1')!;
    expect(reread.getCell('A1').value).toEqual(RICH);
    expect(reread.getCell('A1').type).toBe('String');
    expect(reread.getCell('A2').value).toBe('plain stays plain');
    expect(reread.getCell('A2').type).toBe('String');
    expect(() => XLSX.read(out, { type: 'buffer' })).not.toThrow();
  });

  it('rich text with fonts round-trips fonts exactly', () => {
    const wb = Workbook.read(makeMinimalXlsx());
    wb.getWorksheet('Sheet1')!.getCell('B2').value = {
      richText: [{ text: 'x', font: { name: 'Courier', size: 9, bold: true, italic: true, underline: true, color: 'FF00FF00' } }],
    };
    const reread = Workbook.read(wb.toBuffer()).getWorksheet('Sheet1')!;
    expect(reread.getCell('B2').value).toEqual({
      richText: [{ text: 'x', font: { name: 'Courier', size: 9, bold: true, italic: true, underline: true, color: 'FF00FF00' } }],
    });
  });

  it('SheetJS parses the file and flattens the runs', () => {
    const wb = Workbook.read(makeMinimalXlsx());
    wb.getWorksheet('Sheet1')!.getCell('A1').value = RICH;
    const out = wb.toBuffer();
    const book = XLSX.read(out, { type: 'buffer' });
    expect(book.Sheets.Sheet1!.A1.v).toBe('Hello bold & <tail>');
  });

  it('ExcelJS cross-reads the runs', async () => {
    const wb = Workbook.read(makeMinimalXlsx());
    wb.getWorksheet('Sheet1')!.getCell('A1').value = RICH;
    const ej = await loadInExcelJs(wb.toBuffer());
    const value = ej.getWorksheet('Sheet1')!.getCell('A1').value as RichTextValue;
    expect(value.richText).toHaveLength(3);
    expect(value.richText[0]).toEqual({ text: 'Hello ' });
    expect(value.richText[1]).toMatchObject({
      text: 'bold',
      font: { bold: true, size: 14, name: 'Arial' },
    });
  });

  it('reads rich-text shared strings written by ExcelJS (runs + fonts)', async () => {
    const file = await excelJsFile((ws) => {
      ws.getCell('A1').value = { richText: [{ text: 'ri' }, { font: { bold: true, size: 12 }, text: 'ch' }] };
      ws.getCell('A2').value = 'plain & <special>';
    });
    const ws = Workbook.read(new Uint8Array(file)).getWorksheet('S')!;
    expect(ws.getCell('A1').value).toEqual({
      richText: [{ text: 'ri' }, { text: 'ch', font: { bold: true, size: 12 } }],
    });
    expect(ws.getCell('A2').value).toBe('plain & <special>');
  });

  it('eager and streaming readers agree on inline rich text (written by us)', async () => {
    const wb = Workbook.read(makeMinimalXlsx());
    const ws = wb.getWorksheet('Sheet1')!;
    ws.getCell('A1').value = RICH;
    ws.getCell('B1').value = { richText: [{ text: 'solo' }] };
    ws.getCell('A2').value = 42;
    const out = wb.toBuffer();

    const eager = eagerValues(out, 'Sheet1');
    const streamed = await streamValues(out, 'Sheet1');
    expect(streamed.size).toBe(eager.size);
    for (const [ref, value] of eager) {
      expect(streamed.get(ref), `streamed ${ref}`).toEqual(value);
    }
    expect(eager.get('A1')).toEqual(RICH);
    expect(eager.get('B1')).toEqual({ richText: [{ text: 'solo' }] });
  });

  it('eager and streaming readers agree on rich-text shared strings from ExcelJS', async () => {
    const file = await excelJsFile((ws) => {
      ws.getCell('A1').value = { richText: [{ text: 'ri' }, { font: { bold: true }, text: 'ch' }] };
      // ExcelJS's run-font color is `{ argb }`; ours reads back the flat ARGB string.
      ws.getCell('A2').value = { richText: [{ text: 'one ' }, { font: { color: { argb: 'FF00FF00' } }, text: 'two' }] };
    });
    const eager = eagerValues(new Uint8Array(file), 'S');
    const streamed = await streamValues(new Uint8Array(file), 'S');
    expect(streamed.size).toBe(eager.size);
    for (const [ref, value] of eager) {
      expect(streamed.get(ref), `streamed ${ref}`).toEqual(value);
    }
  });

  it('eager and streaming readers agree on entity-encoded formula strings (regression)', async () => {
    // `<f>` cached results with `t="str"` are entity-encoded in the XML; the
    // streaming reader used to return the raw encoded text.
    const xml =
      '<?xml version="1.0"?><worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">' +
      '<sheetData><row r="1">' +
      '<c r="A1" t="str"><f>COUNTA(B1:B2)</f><v>Brighton &amp; Hove &lt;3</v></c>' +
      '<c r="B1" t="e"><f>X</f><v>#DIV/0!</v></c>' +
      '</row></sheetData></worksheet>';
    const file = makeXlsxWithSheet(xml);
    const eager = eagerValues(file, 'Sheet1');
    const streamed = await streamValues(file, 'Sheet1');
    expect(eager.get('A1')).toEqual({ formula: 'COUNTA(B1:B2)', result: 'Brighton & Hove <3' });
    expect(streamed.get('A1')).toEqual(eager.get('A1'));
    expect(streamed.get('B1')).toEqual(eager.get('B1'));
  });

  it('honors xml:space for runs with leading/trailing whitespace and newlines', () => {
    const wb = Workbook.read(makeMinimalXlsx());
    wb.getWorksheet('Sheet1')!.getCell('A1').value = {
      richText: [{ text: '  padded  ' }, { text: 'line\nbreak', font: { bold: true } }],
    };
    const out = wb.toBuffer();
    expect(sheetOf(out)).toContain('<t xml:space="preserve">  padded  </t>');
    expect(sheetOf(out)).toContain('<t xml:space="preserve">line\nbreak</t>');
    const reread = Workbook.read(out).getWorksheet('Sheet1')!;
    expect(reread.getCell('A1').value).toEqual({
      richText: [{ text: '  padded  ' }, { text: 'line\nbreak', font: { bold: true } }],
    });
  });

  it('keeps reading a border/alignment/rich-text workbook clean after several edits (preservation)', () => {
    const file = makeMinimalXlsx();
    const wb = Workbook.read(file);
    const ws = wb.getWorksheet('Sheet1')!;
    ws.getCell('B1').value = RICH;
    ws.getCell('B1').border = { top: { style: 'thin' } };
    ws.getCell('A1').alignment = { horizontal: 'center' };

    const out = wb.toBuffer();
    const before = Workbook.read(file).parts();
    const after = Workbook.read(out).parts();
    for (const name of before.names()) {
      if (name === 'xl/worksheets/sheet1.xml' || name === 'xl/styles.xml') continue;
      expect(Buffer.from(before.get(name)!).equals(Buffer.from(after.get(name)!)), name).toBe(true);
    }
  });

  it('reading rich text without editing keeps the file byte-identical', async () => {
    const file = await excelJsFile((ws) => {
      ws.getCell('A1').value = { richText: [{ text: 'ri' }, { font: { bold: true }, text: 'ch' }] };
      ws.getCell('A2').border = { top: { style: 'thin' } };
      ws.getCell('A3').alignment = { horizontal: 'center' };
    });
    const wb = Workbook.read(new Uint8Array(file));
    expect(wb.getWorksheet('S')!.getCell('A1').value).toEqual({
      richText: [{ text: 'ri' }, { font: { bold: true }, text: 'ch' }],
    });
    expect(wb.getWorksheet('S')!.getCell('A2').border).toEqual({ top: { style: 'thin' } });
    expect(wb.getWorksheet('S')!.getCell('A3').alignment).toEqual({ horizontal: 'center' });
    expect(wb.toBuffer()).toEqual(new Uint8Array(file));
  });

  it('round-trips every supported border style', () => {
    const allStyles = [
      'thin',
      'medium',
      'thick',
      'dashed',
      'dotted',
      'double',
      'hair',
      'mediumDashed',
      'dashDot',
      'slantDashDot',
      'mediumDashDot',
      'dashDotDot',
      'mediumDashDotDot',
    ] as const;
    const wb = Workbook.read(makeMinimalXlsx());
    const ws = wb.getWorksheet('Sheet1')!;
    allStyles.forEach((style, index) => {
      ws.getCell(index + 1, 1).border = { top: { style } };
    });
    const reread = Workbook.read(wb.toBuffer()).getWorksheet('Sheet1')!;
    allStyles.forEach((style, index) => {
      expect(reread.getCell(index + 1, 1).border).toEqual({ top: { style } });
    });
    expect(() => XLSX.read(wb.toBuffer(), { type: 'buffer' })).not.toThrow();
  });
});