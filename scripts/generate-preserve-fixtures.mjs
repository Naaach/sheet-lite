#!/usr/bin/env node
/**
 * Generates the small `.xlsx` fixtures used by `test/preservation.test.ts`.
 *
 * The README promises that data validations, conditional formatting, tables,
 * images, charts, pivot tables and comments are *preserved, not created*. That
 * promise needs files that actually contain those features, so this script
 * builds them with the devDependencies (ExcelJS, plus fflate for the parts
 * ExcelJS cannot emit) and writes them to `test/fixtures/preserve/`.
 *
 * ## Determinism
 *
 * The output must be byte-identical on every run, otherwise the committed
 * fixtures cannot be regenerated without a diff:
 *
 *  - the workbook `created`/`modified` timestamps are pinned to fixed dates;
 *  - every ZIP entry is re-packed by fflate with one fixed MS-DOS mtime
 *    (ExcelJS streams through `archiver`, whose entry timestamps are not
 *    pinned), so the container bytes never depend on "now";
 *  - no random ids, no auto-incrementing relationship ids: every `rId`,
 *    drawing id, chart axis id and style name is a literal in this file.
 *
 * ## Usage
 *
 *     node scripts/generate-preserve-fixtures.mjs           # write fixtures
 *     node scripts/generate-preserve-fixtures.mjs --check   # verify, exit 1 on drift
 *
 * `--check` regenerates in memory and compares against the committed bytes;
 * it is what CI (or you, before committing) should run to prove the fixtures
 * are reproducible.
 */
import { existsSync, mkdirSync, readFileSync, writeFileSync } from 'node:fs';
import { dirname, join } from 'node:path';
import { fileURLToPath } from 'node:url';
import { createHash } from 'node:crypto';
import ExcelJS from 'exceljs';
import { strToU8, unzipSync, zipSync } from 'fflate';

const here = dirname(fileURLToPath(import.meta.url));
const OUT_DIR = join(here, '..', 'test', 'fixtures', 'preserve');

/** Fixed MS-DOS timestamp stamped on every entry of every fixture. */
const ZIP_MTIME = new Date(Date.UTC(2020, 0, 1, 0, 0, 0));
/** Fixed workbook timestamps, so `docProps/core.xml` never drifts. */
const CREATED = new Date(Date.UTC(2020, 0, 1, 0, 0, 0));
const MODIFIED = new Date(Date.UTC(2020, 0, 2, 0, 0, 0));

/** A real 1x1 PNG (~70 B), small enough to keep the fixtures tiny. */
const PNG_BASE64 =
  'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8z8BQDwAEhQGAhKmMIQAAAABJRU5ErkJggg==';

const DEC = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>';

const REL_NS = 'http://schemas.openxmlformats.org/package/2006/relationships';
const rel = (id, type, target) =>
  `<Relationship Id="${id}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/${type}" Target="${target}"/>`;
const rels = (...entries) =>
  `${DEC}<Relationships xmlns="${REL_NS}">${entries.join('')}</Relationships>`;

/**
 * Re-pack a ZIP with a fixed mtime and without directory entries, so the bytes
 * depend only on the part contents.
 */
function normalize(input) {
  const decoded = unzipSync(new Uint8Array(input));
  const parts = {};
  for (const name of Object.keys(decoded).sort()) {
    if (name.endsWith('/')) continue; // directory entry
    parts[name] = decoded[name];
  }
  return zipSync(parts, { level: 6, mtime: ZIP_MTIME });
}

/** Read one part of a package as a UTF-8 string. */
const text = (parts, name) => new TextDecoder().decode(parts[name]);
const setText = (parts, name, value) => {
  parts[name] = strToU8(value);
};
const insertBefore = (xml, closingTag, snippet) => {
  const at = xml.lastIndexOf(closingTag);
  if (at === -1) throw new Error(`generator bug: ${closingTag} not found`);
  return xml.slice(0, at) + snippet + xml.slice(at);
};

/** Shared ExcelJS workbook setup: pinned timestamps and authors. */
function newWorkbook() {
  const wb = new ExcelJS.Workbook();
  wb.created = CREATED;
  wb.modified = MODIFIED;
  wb.creator = 'sheet-lite fixture generator';
  wb.lastModifiedBy = 'sheet-lite fixture generator';
  wb.title = 'sheet-lite preservation fixture';
  return wb;
}

// --------------------------------------------------------------------------
// 1. data validations + conditional formatting
// --------------------------------------------------------------------------
async function buildValidations() {
  const wb = newWorkbook();
  const ws = wb.addWorksheet('Data');
  ws.getColumn(1).width = 18;
  ws.getColumn(2).width = 10;

  ws.getCell('A1').value = 'Item';
  ws.getCell('B1').value = 'Qty';
  ws.getCell('A1').font = { bold: true };
  ws.getCell('B1').font = { bold: true };

  const rows = [
    ['Widget', 3],
    ['Gadget', 12],
    ['Doohickey', 40],
    ['Thingamajig', 1],
  ];
  rows.forEach(([name, qty], index) => {
    const row = index + 2;
    ws.getCell(`A${row}`).value = name;
    ws.getCell(`B${row}`).value = qty;
    // inline list validation + a numeric validation
    ws.getCell(`C${row}`).dataValidation = {
      type: 'list',
      allowBlank: true,
      formulae: ['"red,green,blue"'],
    };
    ws.getCell(`D${row}`).dataValidation = {
      type: 'whole',
      operator: 'between',
      allowBlank: true,
      formulae: [0, 100],
      showErrorMessage: true,
      errorStyle: 'stop',
      errorTitle: 'Out of range',
      error: 'Enter a whole number between 0 and 100',
    };
  });

  // conditional formatting: a cellIs rule + a colour scale over the same range
  ws.addConditionalFormatting({
    ref: 'B2:B5',
    rules: [
      {
        type: 'cellIs',
        operator: 'greaterThan',
        formulae: [10],
        priority: 1,
        style: { font: { bold: true, color: { argb: 'FF9C0006' } } },
      },
    ],
  });
  ws.addConditionalFormatting({
    ref: 'B2:B5',
    rules: [
      {
        type: 'colorScale',
        priority: 2,
        cfvo: [{ type: 'min' }, { type: 'percentile', value: 50 }, { type: 'max' }],
        color: [{ argb: 'FFF8696B' }, { argb: 'FFFFEB84' }, { argb: 'FF63BE7B' }],
      },
    ],
  });

  // an existing merge, so the "add a merge" edit has company
  ws.mergeCells('A7:B7');
  ws.getCell('A7').value = 'merged header';
  ws.getCell('A7').alignment = { horizontal: 'center' };

  return normalize(await wb.xlsx.writeBuffer());
}

// --------------------------------------------------------------------------
// 2. table + comments/notes (legacy VML) + autofilter
// --------------------------------------------------------------------------
async function buildTableAndComments() {
  const wb = newWorkbook();
  const ws = wb.addWorksheet('Sales');

  ws.addTable({
    name: 'SalesTable',
    ref: 'A1',
    headerRow: true,
    totalsRow: false,
    style: { theme: 'TableStyleMedium2', showRowStripes: true },
    columns: [{ name: 'Product' }, { name: 'Amount' }, { name: 'Note' }],
    rows: [
      ['Widget', 3, 'first'],
      ['Gadget', 12, 'second'],
      ['Doohickey', 40, 'third'],
    ],
  });

  // Cell notes -> xl/comments1.xml + xl/drawings/vmlDrawing1.vml
  ws.getCell('A1').note = 'Product column header';
  ws.getCell('C2').note = 'check this row';
  ws.getCell('A1').note = {
    texts: [
      { font: { bold: true, size: 11, name: 'Calibri' }, text: 'sheet-lite note: ' },
      { font: { italic: true, size: 11, name: 'Calibri' }, text: 'kept verbatim' },
    ],
    margins: { insetmode: 'custom', inset: [0.13, 0.13, 0.25, 0.25] },
  };

  ws.getCell('E1').value = 'untouched';
  ws.mergeCells('E1:F1');

  return normalize(await wb.xlsx.writeBuffer());
}

// --------------------------------------------------------------------------
// 3. image + chart (ExcelJS emits the image; the chart part is injected)
// --------------------------------------------------------------------------
/** Minimal but complete clustered bar chart with cached values, no external data. */
function chartXml() {
  const cats = ['Widget', 'Gadget', 'Doohickey'];
  const vals = [3, 12, 40];
  const strCache = cats
    .map((v, i) => `<c:pt idx="${i}"><c:v>${v}</c:v></c:pt>`)
    .join('');
  const numCache = vals
    .map((v, i) => `<c:pt idx="${i}"><c:v>${v}</c:v></c:pt>`)
    .join('');
  return (
    `${DEC}<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"` +
    ` xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"` +
    ` xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">` +
    `<c:date1904 val="0"/><c:lang val="en-US"/><c:roundedCorners val="0"/>` +
    `<c:chart>` +
    `<c:title><c:tx><c:rich><a:bodyPr/><a:lstStyle/><a:p><a:r><a:t>Units sold</a:t></a:r></a:p></c:rich></c:tx><c:overlay val="0"/></c:title>` +
    `<c:autoTitleDeleted val="0"/>` +
    `<c:plotArea><c:layout/><c:barChart><c:barDir val="col"/><c:grouping val="clustered"/><c:varyColors val="0"/>` +
    `<c:ser><c:idx val="0"/><c:order val="0"/>` +
    `<c:tx><c:strRef><c:f>Data!$B$1</c:f><c:strCache><c:ptCount val="1"/><c:pt idx="0"><c:v>Qty</c:v></c:pt></c:strCache></c:strRef></c:tx>` +
    `<c:cat><c:strRef><c:f>Data!$A$2:$A$4</c:f><c:strCache><c:ptCount val="${cats.length}"/>${strCache}</c:strCache></c:strRef></c:cat>` +
    `<c:val><c:numRef><c:f>Data!$B$2:$B$4</c:f><c:numCache><c:formatCode>General</c:formatCode><c:ptCount val="${vals.length}"/>${numCache}</c:numCache></c:numRef></c:val>` +
    `</c:ser><c:gapWidth val="150"/><c:axId val="111111111"/><c:axId val="222222222"/></c:barChart>` +
    `<c:catAx><c:axId val="111111111"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:delete val="0"/><c:axPos val="b"/><c:crossAx val="222222222"/></c:catAx>` +
    `<c:valAx><c:axId val="222222222"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:delete val="0"/><c:axPos val="l"/><c:crossAx val="111111111"/></c:valAx>` +
    `</c:plotArea><c:legend><c:legendPos val="r"/><c:overlay val="0"/></c:legend>` +
    `<c:plotVisOnly val="1"/><c:dispBlanksAs val="gap"/></c:chart>` +
    `</c:chartSpace>`
  );
}

const CHART_RELS = rels(
  rel('rId1', 'chartColorStyle', 'colors1.xml'),
  rel('rId2', 'chartStyle', 'style1.xml'),
);

const COLORS1 = `${DEC}<cs:colorStyle xmlns:cs="http://schemas.microsoft.com/office/drawing/2012/chartStyle" meth="cycle" id="10"><a:schemeClr xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" val="accent1"/><a:schemeClr xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" val="accent2"/><cs:variation/><cs:variation><a:lumMod xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" val="60000"/></cs:variation></cs:colorStyle>`;

const STYLE1 = `${DEC}<cs:chartStyle xmlns:cs="http://schemas.microsoft.com/office/drawing/2012/chartStyle" id="201"><cs:axisTitle><cs:defRPr/></cs:axisTitle><cs:categoryAxis><cs:defRPr/></cs:categoryAxis><cs:chartArea><cs:defRPr/></cs:chartArea><cs:dataLabel><cs:defRPr/></cs:dataLabel><cs:dataPoint><cs:defRPr/></cs:dataPoint><cs:downBar><cs:defRPr/></cs:downBar><cs:gridlineMajor><cs:defRPr/></cs:gridlineMajor><cs:gridlineMinor><cs:defRPr/></cs:gridlineMinor><cs:hiLoLine><cs:defRPr/></cs:hiLoLine><cs:leaderLine><cs:defRPr/></cs:leaderLine><cs:legendEntry><cs:defRPr/></cs:legendEntry><cs:plotArea><cs:defRPr/></cs:plotArea><cs:serAxis><cs:defRPr/></cs:serAxis><cs:title><cs:defRPr/></cs:title><cs:trendline><cs:defRPr/></cs:trendline><cs:upBar><cs:defRPr/></cs:upBar><cs:upDownBars><cs:defRPr/></cs:upDownBars><cs:valueAxis><cs:defRPr/></cs:valueAxis></cs:chartStyle>`;

/** Anchor that references the chart from the sheet drawing. */
const CHART_ANCHOR =
  '<xdr:twoCellAnchor>' +
  '<xdr:from><xdr:col>7</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>0</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from>' +
  '<xdr:to><xdr:col>13</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>14</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:to>' +
  '<xdr:graphicFrame macro="">' +
  '<xdr:nvGraphicFramePr><xdr:cNvPr id="2" name="Chart 1"/><xdr:cNvGraphicFramePr/></xdr:nvGraphicFramePr>' +
  '<xdr:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/></xdr:xfrm>' +
  '<a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart">' +
  '<c:chart xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" r:id="rId2"/>' +
  '</a:graphicData></a:graphic></xdr:graphicFrame><xdr:clientData/></xdr:twoCellAnchor>';

async function buildImageAndChart() {
  const wb = newWorkbook();
  const ws = wb.addWorksheet('Data');
  ws.getCell('A1').value = 'Product';
  ws.getCell('B1').value = 'Qty';
  ws.getCell('A1').font = { bold: true };
  [
    ['Widget', 3],
    ['Gadget', 12],
    ['Doohickey', 40],
  ].forEach(([name, qty], index) => {
    const row = index + 2;
    ws.getCell(`A${row}`).value = name;
    ws.getCell(`B${row}`).value = qty;
  });

  // ExcelJS owns the picture: xl/media/image1.png + drawing rel rId1.
  ws.addImage(wb.addImage({ extension: 'png', base64: PNG_BASE64 }), 'E1:F5');

  const parts = unzipSync(new Uint8Array(await wb.xlsx.writeBuffer()));

  // ExcelJS cannot emit charts, so wire one into the existing drawing by hand.
  setText(parts, 'xl/charts/chart1.xml', chartXml());
  setText(parts, 'xl/charts/_rels/chart1.xml.rels', CHART_RELS);
  setText(parts, 'xl/charts/colors1.xml', COLORS1);
  setText(parts, 'xl/charts/style1.xml', STYLE1);
  setText(parts, 'xl/drawings/_rels/drawing1.xml.rels',
    rels(rel('rId1', 'image', '../media/image1.png'), rel('rId2', 'chart', '../charts/chart1.xml')));
  setText(parts, 'xl/drawings/drawing1.xml',
    insertBefore(text(parts, 'xl/drawings/drawing1.xml'), '</xdr:wsDr>', CHART_ANCHOR));
  setText(parts, '[Content_Types].xml', insertBefore(text(parts, '[Content_Types].xml'), '</Types>',
    '<Override PartName="/xl/charts/chart1.xml" ContentType="application/vnd.openxmlformats-officedocument.drawingml.chart+xml"/>' +
    '<Override PartName="/xl/charts/colors1.xml" ContentType="application/vnd.ms-office.chartcolorstyle+xml"/>' +
    '<Override PartName="/xl/charts/style1.xml" ContentType="application/vnd.ms-office.chartstyle+xml"/>'));

  return normalize(zipSync(parts, { level: 6, mtime: ZIP_MTIME }));
}

// --------------------------------------------------------------------------
// 4. pivot table (hand-built: ExcelJS has no pivot support at all)
// --------------------------------------------------------------------------
const PIVOT_CT = {
  pivotCacheDefinition:
    'application/vnd.openxmlformats-officedocument.spreadsheetml.pivotCacheDefinition+xml',
  pivotCacheRecords:
    'application/vnd.openxmlformats-officedocument.spreadsheetml.pivotCacheRecords+xml',
  pivotTable: 'application/vnd.openxmlformats-officedocument.spreadsheetml.pivotTable+xml',
};
const WORKSHEET_CT =
  'application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml';

function buildPivot() {
  const main =
    'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml';
  const stylesCt =
    'application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml';

  const contentTypes =
    `${DEC}<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">` +
    '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>' +
    '<Default Extension="xml" ContentType="application/xml"/>' +
    `<Override PartName="/xl/workbook.xml" ContentType="${main}"/>` +
    `<Override PartName="/xl/worksheets/sheet1.xml" ContentType="${WORKSHEET_CT}"/>` +
    `<Override PartName="/xl/worksheets/sheet2.xml" ContentType="${WORKSHEET_CT}"/>` +
    `<Override PartName="/xl/styles.xml" ContentType="${stylesCt}"/>` +
    `<Override PartName="/xl/pivotCache/pivotCacheDefinition1.xml" ContentType="${PIVOT_CT.pivotCacheDefinition}"/>` +
    `<Override PartName="/xl/pivotCache/pivotCacheRecords1.xml" ContentType="${PIVOT_CT.pivotCacheRecords}"/>` +
    `<Override PartName="/xl/pivotTables/pivotTable1.xml" ContentType="${PIVOT_CT.pivotTable}"/>` +
    '</Types>';

  const workbook =
    `${DEC}<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">` +
    '<bookViews><workbookView activeTab="1"/></bookViews>' +
    '<sheets><sheet name="Raw" sheetId="1" r:id="rId1"/><sheet name="Summary" sheetId="2" r:id="rId2"/></sheets>' +
    '<pivotCaches><pivotCache cacheId="1" r:id="rId3"/></pivotCaches>' +
    '<calcPr calcId="171027"/></workbook>';

  const workbookRels = rels(
    rel('rId1', 'worksheet', 'worksheets/sheet1.xml'),
    rel('rId2', 'worksheet', 'worksheets/sheet2.xml'),
    rel('rId3', 'pivotCacheDefinition', 'pivotCache/pivotCacheDefinition1.xml'),
    rel('rId4', 'styles', 'styles.xml'),
  );

  const inline = (ref, value) => `<c r="${ref}" t="inlineStr"><is><t>${value}</t></is></c>`;
  const number = (ref, value) => `<c r="${ref}"><v>${value}</v></c>`;
  const sheet = (dimension, rows) =>
    `${DEC}<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">` +
    `<dimension ref="${dimension}"/><sheetViews><sheetView workbookViewId="0"/></sheetViews>` +
    `<sheetData>${rows}</sheetData></worksheet>`;

  const raw = sheet(
    'A1:B5',
    `<row r="1">${inline('A1', 'Product')}${inline('B1', 'Amount')}</row>` +
      `<row r="2">${inline('A2', 'Widget')}${number('B2', 10)}</row>` +
      `<row r="3">${inline('A3', 'Gadget')}${number('B3', 25)}</row>` +
      `<row r="4">${inline('A4', 'Doohickey')}${number('B4', 7)}</row>` +
      `<row r="5">${inline('A5', 'Thingamajig')}${number('B5', 40)}</row>`,
  );

  // Summary sheet: a header row plus the pivot table area A3:B8.
  const summary = sheet(
    'A1:B8',
    `<row r="1">${inline('A1', 'Pivot report')}</row>` +
      `<row r="3">${inline('A3', 'Product')}${inline('B3', 'Sum of Amount')}</row>` +
      `<row r="4">${inline('A4', 'Widget')}${number('B4', 10)}</row>` +
      `<row r="5">${inline('A5', 'Gadget')}${number('B5', 25)}</row>` +
      `<row r="6">${inline('A6', 'Doohickey')}${number('B6', 7)}</row>` +
      `<row r="7">${inline('A7', 'Thingamajig')}${number('B7', 40)}</row>` +
      `<row r="8">${inline('A8', 'Grand Total')}${number('B8', 82)}</row>`,
  );

  const pivotCacheDefinition =
    `${DEC}<pivotCacheDefinition xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"` +
    ' r:id="rId1" refreshOnLoad="1" recordCount="4" createdVersion="6" refreshedVersion="6" minRefreshableVersion="3" upgradedVersion="0">' +
    '<cacheSource type="worksheet"><worksheetSource ref="A1:B5" sheet="Raw"/></cacheSource>' +
    '<cacheFields count="2">' +
    '<cacheField name="Product" numFmtId="0"><sharedItems count="4">' +
    '<s v="Widget"/><s v="Gadget"/><s v="Doohickey"/><s v="Thingamajig"/></sharedItems></cacheField>' +
    '<cacheField name="Amount" numFmtId="0"><sharedItems containsSemiMixedTypes="0" containsString="0" containsNumber="1" containsInteger="1" minValue="7" maxValue="40"/></cacheField>' +
    '</cacheFields></pivotCacheDefinition>';

  const pivotCacheRecords =
    `${DEC}<pivotCacheRecords xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" count="4">` +
    '<r><x v="0"/><n v="10"/></r><r><x v="1"/><n v="25"/></r><r><x v="2"/><n v="7"/></r><r><x v="3"/><n v="40"/></r>' +
    '</pivotCacheRecords>';

  const pivotTable =
    `${DEC}<pivotTableDefinition xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" name="PivotTable1" cacheId="1"` +
    ' applyNumberFormats="0" applyBorderFormats="0" applyFontFormats="0" applyPatternFormats="0"' +
    ' applyAlignmentFormats="0" applyWidthHeightFormats="1" dataCaption="Values" updatedVersion="6"' +
    ' minRefreshableVersion="3" useAutoFormatting="1" itemPrintTitles="1" createdVersion="6" indent="0"' +
    ' outline="1" outlineData="1" multipleFieldFilters="0">' +
    '<location ref="A3:B8" firstHeaderRow="1" firstDataRow="1" firstDataCol="1"/>' +
    '<pivotFields count="2">' +
    '<pivotField axis="axisRow" showAll="0"><items count="5"><item x="0"/><item x="1"/><item x="2"/><item x="3"/><item t="default"/></items></pivotField>' +
    '<pivotField dataField="1" showAll="0"/>' +
    '</pivotFields>' +
    '<rowFields count="1"><field x="0"/></rowFields>' +
    '<rowItems count="5"><i><x/></i><i><x v="1"/></i><i><x v="2"/></i><i><x v="3"/></i><i t="grand"><x/></i></rowItems>' +
    '<colItems count="1"><i/></colItems>' +
    '<dataFields count="1"><dataField name="Sum of Amount" fld="1" baseField="0" baseItem="0"/></dataFields>' +
    '<pivotTableStyleInfo name="PivotStyleLight16" showRowHeaders="1" showColHeaders="1" showRowStripes="0" showColStripes="0" showLastColumn="1"/>' +
    '</pivotTableDefinition>';

  const styles =
    `${DEC}<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">` +
    '<fonts count="1"><font><sz val="11"/><name val="Calibri"/></font></fonts>' +
    '<fills count="1"><fill><patternFill patternType="none"/></fill></fills>' +
    '<borders count="1"><border/></borders>' +
    '<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>' +
    '<cellXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/></cellXfs>' +
    '</styleSheet>';

  const parts = {
    '[Content_Types].xml': strToU8(contentTypes),
    '_rels/.rels': strToU8(rels(rel('rId1', 'officeDocument', 'xl/workbook.xml'))),
    'xl/workbook.xml': strToU8(workbook),
    'xl/_rels/workbook.xml.rels': strToU8(workbookRels),
    'xl/worksheets/sheet1.xml': strToU8(raw),
    'xl/worksheets/sheet2.xml': strToU8(summary),
    'xl/worksheets/_rels/sheet2.xml.rels': strToU8(
      rels(rel('rId1', 'pivotTable', '../pivotTables/pivotTable1.xml')),
    ),
    'xl/styles.xml': strToU8(styles),
    'xl/pivotCache/pivotCacheDefinition1.xml': strToU8(pivotCacheDefinition),
    'xl/pivotCache/_rels/pivotCacheDefinition1.xml.rels': strToU8(
      rels(rel('rId1', 'pivotCacheRecords', 'pivotCacheRecords1.xml')),
    ),
    'xl/pivotCache/pivotCacheRecords1.xml': strToU8(pivotCacheRecords),
    'xl/pivotTables/pivotTable1.xml': strToU8(pivotTable),
  };
  return zipSync(parts, { level: 6, mtime: ZIP_MTIME });
}

// --------------------------------------------------------------------------
const FIXTURES = [
  ['validations-and-conditional-formatting.xlsx', buildValidations],
  ['table-and-comments.xlsx', buildTableAndComments],
  ['image-and-chart.xlsx', buildImageAndChart],
  ['pivot-table.xlsx', buildPivot],
];

const sha256 = (bytes) => createHash('sha256').update(bytes).digest('hex');

async function main() {
  const check = process.argv.includes('--check');
  if (!check) mkdirSync(OUT_DIR, { recursive: true });

  let drifted = 0;
  for (const [name, build] of FIXTURES) {
    const bytes = Buffer.from(await build());
    const path = join(OUT_DIR, name);
    if (check) {
      const current = existsSync(path) ? readFileSync(path) : null;
      if (current === null || !current.equals(bytes)) {
        drifted += 1;
        console.error(`DRIFT  ${name}: ${current ? sha256(current).slice(0, 16) : 'missing'} -> ${sha256(bytes).slice(0, 16)}`);
      } else {
        console.log(`ok     ${name}  ${(bytes.length / 1024).toFixed(1)} kB  ${sha256(bytes).slice(0, 16)}`);
      }
      continue;
    }
    writeFileSync(path, bytes);
    console.log(`wrote  ${name}  ${(bytes.length / 1024).toFixed(1)} kB  ${sha256(bytes).slice(0, 16)}`);
  }

  if (check) {
    if (drifted > 0) {
      console.error(`\n${drifted} fixture(s) differ from the committed bytes — run: node scripts/generate-preserve-fixtures.mjs`);
      process.exit(1);
    }
    console.log('\nAll committed fixtures match a fresh generation (deterministic).');
  }
}

await main();