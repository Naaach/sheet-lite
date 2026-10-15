/**
 * Frozen OPC / SpreadsheetML constants the writer emits into the parts it
 * creates.
 *
 * These are specification strings (namespaces, content types, relationship
 * types) — the eager writer keeps its own copies for the same reason it keeps
 * its own minimal parts: the writer is an on-demand chunk and anything it
 * shares with the gzip-budgeted main entry has to be imported by that entry.
 * The one *rule* both writers need — Excel's worksheet-name limits — is not
 * duplicated: it lives in `opc/sheet-name.ts` and is shared.
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
export const OFFICE_DOCUMENT_REL =
  'http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument';
export const XML_CONTENT_TYPE = 'application/xml';
export const RELS_CONTENT_TYPE = 'application/vnd.openxmlformats-package.relationships+xml';

export const OFFICE_XMLNS = 'xmlns="http://schemas.openxmlformats.org/package/2006/relationships"';
export const CONTENT_TYPES_XMLNS = 'xmlns="http://schemas.openxmlformats.org/package/2006/content-types"';