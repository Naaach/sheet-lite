import { Workbook, writeZip, type ZipEntry } from '../../src/index';

/**
 * The playground's sample files are built entirely in the browser:
 * a minimal-but-valid OPC container is assembled with the library's own
 * `writeZip`, then all content (values, dates, formulas, number formats,
 * merges) is written through the public Workbook/Worksheet/Cell API.
 */

const enc = new TextEncoder();

const part = (name: string, xml: string): ZipEntry => ({ name, data: enc.encode(xml) });

const CONTENT_TYPES =
  '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
  '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">' +
  '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>' +
  '<Default Extension="xml" ContentType="application/xml"/>' +
  '<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>' +
  '<Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>' +
  '<Override PartName="/xl/worksheets/sheet2.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>' +
  '<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>' +
  '</Types>';

const ROOT_RELS =
  '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
  '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
  '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>' +
  '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/>' +
  '<Relationship Id="rId3" Type="http://schemas.openxmlformats.org/package/2006/relationships/extended-properties" Target="docProps/app.xml"/>' +
  '</Relationships>';

const WORKBOOK =
  '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
  '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">' +
  '<sheets>' +
  '<sheet name="Sales" sheetId="1" r:id="rId1"/>' +
  '<sheet name="Readme" sheetId="2" r:id="rId2"/>' +
  '</sheets>' +
  '</workbook>';

const WORKBOOK_RELS =
  '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
  '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
  '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>' +
  '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet2.xml"/>' +
  '<Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>' +
  '</Relationships>';

// `sheetData` must be a real element pair (not self-closing) so the patch
// writer can splice new rows inside it.
const SHEET1 =
  '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
  '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">' +
  '<dimension ref="A1"/>' +
  '<sheetViews><sheetView workbookViewId="0"/></sheetViews>' +
  '<sheetData></sheetData>' +
  '</worksheet>';

const SHEET2 = SHEET1;

const STYLES =
  '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
  '<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">' +
  '<fonts count="1"><font><sz val="11"/><name val="Calibri"/></font></fonts>' +
  '<fills count="1"><fill><patternFill patternType="none"/></fill></fills>' +
  '<borders count="1"><border/></borders>' +
  '<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>' +
  '<cellXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/></cellXfs>' +
  '</styleSheet>';

const CORE =
  '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
  '<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/">' +
  '<dc:creator>sheet-lite playground</dc:creator>' +
  '<cp:lastModifiedBy>sheet-lite playground</cp:lastModifiedBy>' +
  '</cp:coreProperties>';

const APP =
  '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
  '<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties">' +
  '<Application>sheet-lite playground</Application>' +
  '</Properties>';

function archiveBytes(entries: ZipEntry[]): Uint8Array {
  // `writeZip` only reads `entries`/`byName`; `raw` is the pass-through buffer
  // of an existing package, which a synthesized container does not have.
  return writeZip({
    entries,
    byName: new Map(entries.map((entry) => [entry.name, entry])),
    raw: new Uint8Array(0),
  });
}

/** A deliberately malformed package: an entry with a `../` path-traversal name. */
export function createUnsafeBytes(): Uint8Array {
  return archiveBytes([
    part('[Content_Types].xml', CONTENT_TYPES),
    { name: '../evil.txt', data: enc.encode('boom') },
  ]);
}

/**
 * Generate the demo workbook:
 * Sales sheet — headers, numbers with formats, a real date, SUM/multiply
 * formulas, booleans and a merged banner row. Readme sheet — instructions
 * plus another merge, so sheet tabs and two parts get exercised.
 */
export function createSampleWorkbook(): Workbook {
  const wb = Workbook.read(
    archiveBytes([
      part('[Content_Types].xml', CONTENT_TYPES),
      part('_rels/.rels', ROOT_RELS),
      part('docProps/core.xml', CORE),
      part('docProps/app.xml', APP),
      part('xl/workbook.xml', WORKBOOK),
      part('xl/_rels/workbook.xml.rels', WORKBOOK_RELS),
      part('xl/worksheets/sheet1.xml', SHEET1),
      part('xl/worksheets/sheet2.xml', SHEET2),
      part('xl/styles.xml', STYLES),
    ]),
  );

  const sales = wb.getWorksheet('Sales')!;

  const headers = ['Region', 'Units', 'Unit price', 'Total', 'Launch date', 'Priority', 'Notes'];
  headers.forEach((header, index) => {
    const cell = sales.getCell(1, index + 1);
    cell.value = header;
    cell.font = { bold: true };
  });

  const rows: Array<[string, number, number, Date, boolean, string]> = [
    ['North', 120, 9.99, new Date(2026, 0, 15), true, 'Q1 launch'],
    ['South', 85, 14.5, new Date(2026, 1, 1), false, ''],
    ['East', 210, 7.25, new Date(2026, 1, 20), true, 'Repeat order'],
    ['West', 64, 22, new Date(2026, 2, 5), false, 'Awaiting PO'],
  ];
  rows.forEach(([region, units, price, date, priority, notes], i) => {
    const row = i + 2;
    sales.getCell(row, 1).value = region;
    sales.getCell(row, 2).value = units;
    const priceCell = sales.getCell(row, 3);
    priceCell.value = price;
    priceCell.numFmt = '#,##0.00';
    const totalCell = sales.getCell(row, 4);
    totalCell.value = { formula: `B${row}*C${row}`, result: Math.round(units * price * 100) / 100 };
    totalCell.numFmt = '#,##0.00';
    const dateCell = sales.getCell(row, 5);
    dateCell.value = date;
    // Built-in date format (id 14) on purpose: custom format codes would make
    // the library create a <numFmts> container, and StyleRegistry.toXml()
    // currently drops that element when new fonts are added in the same write
    // (two edits land on the same offset). See the report notes.
    dateCell.numFmt = 'mm-dd-yy';
    sales.getCell(row, 6).value = priority;
    sales.getCell(row, 7).value = notes;
  });

  const totalRow = rows.length + 2; // row 6
  sales.getCell(totalRow, 1).value = 'Total';
  sales.getCell(totalRow, 1).font = { bold: true };
  const unitsTotal = rows.reduce((sum, r) => sum + r[1], 0);
  const salesTotalCell = sales.getCell(totalRow, 4);
  salesTotalCell.value = {
    formula: `SUM(D2:D${totalRow - 1})`,
    result: Math.round(rows.reduce((sum, r) => sum + r[1] * r[2], 0) * 100) / 100,
  };
  salesTotalCell.numFmt = '#,##0.00';
  const unitsTotalCell = sales.getCell(totalRow, 2);
  unitsTotalCell.value = { formula: `SUM(B2:B${totalRow - 1})`, result: unitsTotal };
  unitsTotalCell.font = { bold: true };

  const bannerCell = sales.getCell(totalRow + 2, 1);
  bannerCell.value =
    'Generated in-browser by sheet-lite — edit anything below and watch the patch monitor.';
  bannerCell.font = { bold: true, color: 'FF2563EB' };
  sales.mergeCells(`A${totalRow + 2}:G${totalRow + 2}`);

  const readme = wb.getWorksheet('Readme')!;
  const lines = [
    'sheet-lite playground sample',
    'This file was generated in your browser using the library itself.',
    'Go to the Sales sheet, edit cells, toggle bold, add a row —',
    'then watch the patch monitor: only re-serialized parts change.',
    '',
    'Try: type =B2*2 into an empty cell, then download the result.',
  ];
  lines.forEach((line, i) => {
    const cell = readme.getCell(i + 1, 1);
    cell.value = line;
    if (i === 0) cell.font = { bold: true, size: 14 };
  });
  readme.mergeCells('A8:D9');
  readme.getCell(8, 1).value = 'Merged A8:D9 — merges are written as <mergeCell> on save.';
  readme.getCell(8, 1).font = { italic: true };

  return wb;
}
