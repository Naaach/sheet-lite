import { parseXml, child, childrenOf, textOf, type XmlNode } from '../xml/xml';
import { columnToNumber, coordsToRef, numberToColumn, refToCoords } from '../model/refs';
import { serialToDate } from '../model/dates';
import { typeOf, type CellValue, type ValueTypeName } from '../model/value';
import { parseInlineRichText } from './rich-text';
import type { PackageMeta } from './package-reader';

export { columnToNumber, coordsToRef, numberToColumn, refToCoords };

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
  /** Row height from the `ht` attribute, when set. */
  height?: number;
}

export interface MergeChild {
  ref: string;
  start: number;
  end: number;
}

export interface DimensionRecord {
  start: number;
  end: number;
  /** Raw ref attribute, e.g. "A1:D4". */
  ref: string;
}

export interface ColRecord {
  min: number;
  max: number;
  /** All attributes of the `<col>` element (width, style, hidden, customWidth, ...). */
  attrs: Record<string, string>;
  start: number;
  end: number;
}

export interface PaneRecord {
  start: number;
  end: number;
  attrs: Record<string, string>;
}

export interface SheetViewRecord {
  start: number;
  openEnd: number;
  closeStart: number;
  end: number;
  attrs: Record<string, string>;
  pane: PaneRecord | null;
}

export interface SheetViewsRecord {
  start: number;
  openEnd: number;
  closeStart: number;
  end: number;
  views: SheetViewRecord[];
}

export interface ParsedSheet {
  xml: string;
  cells: Map<string, CellRecord>;
  rows: Map<number, RowRecord>;
  /** Index of `<sheetData` (or of `</worksheet>` when sheetData is absent). */
  sheetDataStart: number;
  /** True when sheetData is self-closing (`<sheetData/>`): new rows must replace it. */
  sheetDataSelfClosing: boolean;
  /** Index of the `</sheetData>` close tag in the original XML. */
  sheetDataClose: number;
  /** Index just after `</sheetData>` (where sibling elements such as mergeCells go). */
  sheetDataEnd: number;
  dimension: DimensionRecord | null;
  sheetFormatPrStart: number | null;
  cols: { start: number; openEnd: number; closeStart: number; end: number; entries: ColRecord[] } | null;
  sheetViews: SheetViewsRecord | null;
  mergeCells: { start: number; openEnd: number; closeStart: number; end: number } | null;
  mergeChildren: MergeChild[];
  maxRow: number;
  maxCol: number;
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
    if (!isNode) value = '';
    else value = parseInlineRichText(isNode) ?? textOf(isNode);
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
        height: rowNode.attrs.ht !== undefined ? Number(rowNode.attrs.ht) : undefined,
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

  const dimensionNode = child(worksheet, 'dimension');
  const dimension: DimensionRecord | null = dimensionNode
    ? { start: dimensionNode.start, end: dimensionNode.end, ref: dimensionNode.attrs.ref ?? '' }
    : null;

  const sheetFormatPr = child(worksheet, 'sheetFormatPr');

  const colsNode = child(worksheet, 'cols');
  let cols: ParsedSheet['cols'] = null;
  if (colsNode) {
    const entries: ColRecord[] = [];
    for (const col of childrenOf(colsNode, 'col')) {
      const min = Number(col.attrs.min);
      const max = col.attrs.max !== undefined ? Number(col.attrs.max) : min;
      if (!Number.isFinite(min) || !Number.isFinite(max) || max < min) continue;
      entries.push({ min, max, attrs: { ...col.attrs }, start: col.start, end: col.end });
    }
    cols = {
      start: colsNode.start,
      openEnd: colsNode.openEnd,
      closeStart: colsNode.closeStart,
      end: colsNode.end,
      entries,
    };
  }

  const sheetViewsNode = child(worksheet, 'sheetViews');
  let sheetViews: ParsedSheet['sheetViews'] = null;
  if (sheetViewsNode) {
    const views: SheetViewRecord[] = [];
    for (const view of childrenOf(sheetViewsNode, 'sheetView')) {
      const paneNode = child(view, 'pane');
      views.push({
        start: view.start,
        openEnd: view.openEnd,
        closeStart: view.closeStart,
        end: view.end,
        attrs: { ...view.attrs },
        pane: paneNode
          ? { start: paneNode.start, end: paneNode.end, attrs: { ...paneNode.attrs } }
          : null,
      });
    }
    sheetViews = {
      start: sheetViewsNode.start,
      openEnd: sheetViewsNode.openEnd,
      closeStart: sheetViewsNode.closeStart,
      end: sheetViewsNode.end,
      views,
    };
  }

  // Positions are taken from the parsed nodes so self-closing `<sheetData/>`
  // (new sheets, sheets without rows) is handled: rows can never be appended
  // after a self-closing sheetData, the element must be replaced instead.
  let sheetDataStart: number;
  let sheetDataSelfClosing = false;
  let sheetDataClose: number;
  let sheetDataEnd: number;
  if (sheetData) {
    sheetDataStart = sheetData.start;
    sheetDataSelfClosing = sheetData.closeStart === sheetData.openEnd;
    sheetDataClose = sheetDataSelfClosing ? sheetData.end : sheetData.closeStart;
    sheetDataEnd = sheetData.end;
  } else {
    const fallback = xml.lastIndexOf('</worksheet>');
    sheetDataStart = fallback === -1 ? xml.length : fallback;
    sheetDataClose = sheetDataStart;
    sheetDataEnd = sheetDataStart;
  }
  return {
    xml,
    cells,
    rows,
    sheetDataStart,
    sheetDataSelfClosing,
    sheetDataClose,
    sheetDataEnd,
    dimension,
    sheetFormatPrStart: sheetFormatPr ? sheetFormatPr.start : null,
    cols,
    sheetViews,
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
