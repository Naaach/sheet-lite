/**
 * The minimal OPC parts a workbook cannot exist without.
 *
 * Eager-path only: `Workbook.addWorksheet` creates these when the workbook it
 * opened lacks them. The incremental writer (`sheet-lite/write`) builds its
 * package from its own constants and is loaded on demand — keeping this module
 * out of the writer's import graph keeps the gzip-budgeted main entry at one
 * fewer shared chunk.
 */
export const XML_DECL = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>';
export const SHEET_NS =
  'xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" ' +
  'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"';

export const MAIN_CONTENT_TYPE =
  'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml';
export const SHEET_CONTENT_TYPE =
  'application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml';
export const STYLES_CONTENT_TYPE =
  'application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml';
export const WORKSHEET_REL =
  'http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet';
export const STYLES_REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles';
export const XML_CONTENT_TYPE = 'application/xml';
export const RELS_CONTENT_TYPE = 'application/vnd.openxmlformats-package.relationships+xml';

/** Minimal valid styleSheet, used when a source workbook has no styles.xml. */
export const MINIMAL_STYLES =
  XML_DECL +
  `<styleSheet ${SHEET_NS}>` +
  '<fonts count="1"><font><sz val="11"/><name val="Calibri"/></font></fonts>' +
  '<fills count="1"><fill><patternFill patternType="none"/></fill></fills>' +
  '<borders count="1"><border/></borders>' +
  '<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>' +
  '<cellXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/></cellXfs>' +
  '</styleSheet>';

/** `[Content_Types].xml` skeleton (only the workbook override; callers add the rest). */
export function minimalContentTypes(): string {
  return (
    XML_DECL +
    '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">' +
    `<Default Extension="rels" ContentType="${RELS_CONTENT_TYPE}"/>` +
    `<Default Extension="xml" ContentType="${XML_CONTENT_TYPE}"/>` +
    `<Override PartName="/xl/workbook.xml" ContentType="${MAIN_CONTENT_TYPE}"/>` +
    '</Types>'
  );
}

/** `xl/_rels/workbook.xml.rels` skeleton (callers add the relationships). */
export function minimalWorkbookRels(): string {
  return (
    XML_DECL +
    '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
    '</Relationships>'
  );
}