import { strFromU8 } from 'fflate';
import { resolvePartPath } from '../opc/relationships';
import { SheetLiteError } from '../zip/errors';
import { parseXml, child, childrenOf, descendants } from '../xml/xml';
import { XmlPullParser } from '../xml/pull-parser';
import { RichTextRunCollector } from './rich-text';
import { isBuiltinDateFormat, isDateFormatCode } from '../model/dates';
import type { RichTextValue } from '../model/value';

export interface SheetRef {
  name: string;
  sheetId: string;
  rId: string;
  /** Absolute part name, e.g. "/xl/worksheets/sheet1.xml". */
  partName: string;
}

export interface PackageMeta {
  date1904: boolean;
  sheets: SheetRef[];
  /** Plain string, or `{ richText: [...] }` when the shared string has runs. */
  sharedStrings: Array<string | RichTextValue>;
  /** numFmtId per index into `cellXfs`. */
  cellXfs: number[];
  numFmts: Map<number, string>;
  /** Indexes into `cellXfs` whose format represents a date/time. */
  dateStyleIndexes: Set<number>;
}

/** Minimal part access: satisfied by `PartStore` and by the streaming reader's small-parts map. */
export interface PartLookup {
  get(name: string): Uint8Array | undefined;
}

function decode(data: Uint8Array): string {
  return strFromU8(data);
}

function rootElement(xml: string, name: string) {
  const doc = parseXml(xml);
  return doc.children.find((n) => n.name === name) ?? doc;
}

function readRels(store: PartLookup, relsPath: string, basePart: string): Map<string, string> {
  const map = new Map<string, string>();
  const data = store.get(relsPath);
  if (!data) return map;
  for (const rel of descendants(parseXml(decode(data)), 'Relationship')) {
    const id = rel.attrs.Id;
    const target = rel.attrs.Target;
    if (!id || !target) continue;
    map.set(id, resolvePartPath(basePart, target));
  }
  return map;
}

/** Decode a shared-strings XML part into the string table (incrementally).
 * Entries with `<r>` runs decode to `{ richText: [...] }` via the shared
 * collector, so eager and streaming readers see the same shapes. */
function parseSharedStrings(data: Uint8Array, out: Array<string | RichTextValue>): void {
  const parser = new XmlPullParser();
  const decoder = new TextDecoder();
  let inSi = false;
  let collector = new RichTextRunCollector();

  const drain = (): void => {
    for (;;) {
      const event = parser.next();
      if (event === null) return;
      if (event.kind === 'text') {
        if (inSi) collector.text(event.text);
        continue;
      }
      if (event.kind === 'open') {
        if (event.name === 'si') {
          inSi = true;
          collector = new RichTextRunCollector();
          if (event.selfClosing) {
            inSi = false;
            out.push('');
          }
          continue;
        }
        if (!inSi) continue;
        switch (event.name) {
          case 'r':
            collector.openRun();
            if (event.selfClosing) collector.closeRun();
            break;
          case 'rPh':
            collector.openSkip();
            if (event.selfClosing) collector.closeSkip();
            break;
          case 'rPr':
            collector.openRPr();
            if (event.selfClosing) collector.closeRPr();
            break;
          case 't':
            collector.openT();
            if (event.selfClosing) collector.closeT();
            break;
          case 'b':
          case 'i':
          case 'u':
          case 'sz':
          case 'color':
          case 'rFont':
            collector.fontElement(event.name, event.attrs);
            break;
          default:
            break;
        }
        continue;
      }
      if (event.kind === 'close' && inSi) {
        switch (event.name) {
          case 't':
            collector.closeT();
            break;
          case 'r':
            collector.closeRun();
            break;
          case 'rPh':
            collector.closeSkip();
            break;
          case 'rPr':
            collector.closeRPr();
            break;
          case 'si':
            inSi = false;
            out.push(collector.value());
            break;
          default:
            break;
        }
      }
    }
  };

  const CHUNK = 1 << 16;
  for (let at = 0; at < data.length; at += CHUNK) {
    parser.push(decoder.decode(data.subarray(at, at + CHUNK), { stream: true }));
    drain();
  }
  parser.push(decoder.decode());
  parser.end();
  drain();
}

export function readPackageMeta(store: PartLookup): PackageMeta {
  const workbookData = store.get('xl/workbook.xml');
  if (!workbookData) {
    throw new SheetLiteError('Not an XLSX package: missing xl/workbook.xml', 'NOT_XLSX');
  }

  const workbook = rootElement(decode(workbookData), 'workbook');
  const workbookPr = child(workbook, 'workbookPr');
  const date1904 = workbookPr?.attrs.date1904 === '1' || workbookPr?.attrs.date1904 === 'true';

  const rels = readRels(store, 'xl/_rels/workbook.xml.rels', '/xl/workbook.xml');
  const sheets: SheetRef[] = [];
  const sheetsNode = child(workbook, 'sheets');
  for (const sheet of childrenOf(sheetsNode ?? workbook, 'sheet')) {
    const rId = sheet.attrs['r:id'] ?? '';
    const partName = rels.get(rId);
    if (!partName) continue;
    sheets.push({
      name: sheet.attrs.name ?? '',
      sheetId: sheet.attrs.sheetId ?? '',
      rId,
      partName,
    });
  }

  const sharedStrings: string[] = [];
  const ssData = store.get('xl/sharedStrings.xml');
  if (ssData) parseSharedStrings(ssData, sharedStrings);

  const cellXfs: number[] = [];
  const numFmts = new Map<number, string>();
  const stylesData = store.get('xl/styles.xml');
  if (stylesData) {
    const styleSheet = rootElement(decode(stylesData), 'styleSheet');
    const numFmtsNode = child(styleSheet, 'numFmts');
    if (numFmtsNode) {
      for (const nf of childrenOf(numFmtsNode, 'numFmt')) {
        numFmts.set(Number(nf.attrs.numFmtId), nf.attrs.formatCode ?? '');
      }
    }
    const xfsNode = child(styleSheet, 'cellXfs');
    if (xfsNode) {
      for (const xf of childrenOf(xfsNode, 'xf')) {
        cellXfs.push(Number(xf.attrs.numFmtId ?? '0'));
      }
    }
  }

  const dateStyleIndexes = new Set<number>();
  cellXfs.forEach((numFmtId, index) => {
    if (isBuiltinDateFormat(numFmtId)) {
      dateStyleIndexes.add(index);
      return;
    }
    const code = numFmts.get(numFmtId);
    if (code && isDateFormatCode(code)) dateStyleIndexes.add(index);
  });

  return { date1904, sheets, sharedStrings, cellXfs, numFmts, dateStyleIndexes };
}
