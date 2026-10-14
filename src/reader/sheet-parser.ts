import { parseXml, child, childrenOf, textOf, type XmlNode } from '../xml/xml';
import { serialToDate } from '../model/dates';
import { typeOf, type CellValue, type ValueTypeName } from '../model/value';
import type { PackageMeta } from './package-reader';

export interface CellRecord {
  ref: string;
  row: number;
  col: number;
  /** Byte range of the `<c>` element in the original sheet XML. */
  start: number;
  end: number;
  styleIndex: number;
  value: CellValue;
  type: ValueTypeName;
}

export interface RowRecord {
  r: number;
  start: number;
  openEnd: number;
  closeStart: number;
  end: number;
  selfClosing: boolean;
}

export interface MergeChild {
  ref: string;
  start: number;
  end: number;
}

export interface ParsedSheet {
  xml: string;
  cells: Map<string, CellRecord>;
  rows: Map<number, RowRecord>;
  /** Index of the `</sheetData>` close tag in the original XML. */
  sheetDataClose: number;
  /** Index just after `</sheetData>` (where sibling elements such as mergeCells go). */
  sheetDataEnd: number;
  mergeCells: { start: number; openEnd: number; closeStart: number; end: number } | null;
  mergeChildren: MergeChild[];
  maxRow: number;
  maxCol: number;
}

export function columnToNumber(letters: string): number {
  let n = 0;
  for (const ch of letters.toUpperCase()) n = n * 26 + (ch.charCodeAt(0) - 64);
  return n;
}

export function numberToColumn(n: number): string {
  let s = '';
  let value = n;
  while (value > 0) {
    const rem = (value - 1) % 26;
    s = String.fromCharCode(65 + rem) + s;
    value = Math.floor((value - 1) / 26);
  }
  return s;
}

export function refToCoords(ref: string): { row: number; col: number } {
  const match = /^([A-Za-z]+)(\d+)$/.exec(ref);
  if (!match) return { row: 0, col: 0 };
  return { row: Number(match[2]), col: columnToNumber(match[1]!) };
}

export function coordsToRef(row: number, col: number): string {
  return `${numberToColumn(col)}${row}`;
}

function computeValue(c: XmlNode, meta: PackageMeta): CellValue {
  const t = c.attrs.t;
  const styleIndex = c.attrs.s ? Number(c.attrs.s) : 0;
  const vNode = child(c, 'v');
  const rawText = vNode ? textOf(vNode) : undefined;
  const hasText = rawText !== undefined && rawText !== '';
  const fNode = child(c, 'f');

  let value: CellValue = null;
  if (t === 's') {
    value = hasText ? meta.sharedStrings[Number(rawText)] ?? '' : '';
  } else if (t === 'inlineStr') {
    const isNode = child(c, 'is');
    value = isNode ? textOf(isNode) : '';
  } else if (t === 'str') {
    value = rawText ?? '';
  } else if (t === 'b') {
    value = hasText ? rawText === '1' : null;
  } else if (t === 'e') {
    value = hasText ? { error: rawText } : null;
  } else if (hasText) {
    const numeric = Number(rawText);
    value = meta.dateStyleIndexes.has(styleIndex)
      ? serialToDate(numeric, meta.date1904)
      : numeric;
  }

  if (fNode) return { formula: textOf(fNode), result: value ?? undefined };
  return value;
}

export function parseSheet(xml: string, meta: PackageMeta): ParsedSheet {
  const document = parseXml(xml);
  const worksheet = document.children.find((n) => n.name === 'worksheet') ?? document;
  const sheetData = child(worksheet, 'sheetData');

  const cells = new Map<string, CellRecord>();
  const rows = new Map<number, RowRecord>();
  let maxRow = 0;
  let maxCol = 0;

  if (sheetData) {
    let implicitRow = 0;
    for (const rowNode of childrenOf(sheetData, 'row')) {
      implicitRow += 1;
      const r = rowNode.attrs.r ? Number(rowNode.attrs.r) : implicitRow;
      rows.set(r, {
        r,
        start: rowNode.start,
        openEnd: rowNode.openEnd,
        closeStart: rowNode.closeStart,
        end: rowNode.end,
        selfClosing: rowNode.closeStart === rowNode.openEnd,
      });
      if (r > maxRow) maxRow = r;

      let prevCol = 0;
      for (const c of childrenOf(rowNode, 'c')) {
        let ref = c.attrs.r;
        let col: number;
        if (ref) {
          col = refToCoords(ref).col;
        } else {
          col = prevCol + 1;
          ref = coordsToRef(r, col);
        }
        prevCol = col;
        const { row } = refToCoords(ref);
        const value = computeValue(c, meta);
        cells.set(ref, {
          ref,
          row,
          col,
          start: c.start,
          end: c.end,
          styleIndex: c.attrs.s ? Number(c.attrs.s) : 0,
          value,
          type: typeOf(value),
        });
        if (col > maxCol) maxCol = col;
        if (row > maxRow) maxRow = row;
      }
    }
  }

  const mergeCellsNode = child(worksheet, 'mergeCells') ?? null;
  const mergeChildren: MergeChild[] = [];
  if (mergeCellsNode) {
    for (const merge of childrenOf(mergeCellsNode, 'mergeCell')) {
      mergeChildren.push({ ref: merge.attrs.ref ?? '', start: merge.start, end: merge.end });
    }
  }

  const closeIdx = xml.lastIndexOf('</sheetData>');
  const sheetDataClose = closeIdx === -1 ? xml.lastIndexOf('</worksheet>') : closeIdx;
  const sheetDataEnd = closeIdx === -1 ? sheetDataClose : closeIdx + '</sheetData>'.length;
  return {
    xml,
    cells,
    rows,
    sheetDataClose,
    sheetDataEnd,
    mergeCells: mergeCellsNode
      ? {
          start: mergeCellsNode.start,
          openEnd: mergeCellsNode.openEnd,
          closeStart: mergeCellsNode.closeStart,
          end: mergeCellsNode.end,
        }
      : null,
    mergeChildren,
    maxRow,
    maxCol,
  };
}
