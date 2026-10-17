import { escapeXml } from '../xml/xml';
import type { CellValue } from '../model/value';
import type { CellStyleOverrides } from '../model/style';
import type { SheetViewProps } from '../model/view';
import { assertMergeRange } from '../model/refs';
import { coordsToRef, refToCoords, type ColRecord, type ParsedSheet } from '../reader/sheet-parser';
import type { StyleRegistry } from './style-registry';
import { generateCellXml } from './cell-xml';
import { paneXmlFor } from './cell-xml';

export interface CellEdit {
  hasValue: boolean;
  value: CellValue;
  style?: CellStyleOverrides;
}

/** Width/hidden overrides for one column. `null` clears the property. */
export interface ColEdit {
  width?: number | null;
  hidden?: boolean | null;
}

/**
 * Everything a write can change in one worksheet part. Every collection may be
 * empty; the writer only emits edits for entries that are actually present.
 */
export interface SheetPatch {
  changes: Map<string, CellEdit>;
  newMerges: string[];
  removedMerges: string[];
  /** 1-based column -> width/hidden overrides. */
  colEdits: Map<number, ColEdit>;
  /** 1-based row -> height (`null` clears a previously set height). */
  rowHeights: Map<number, number | null>;
  /** Requested sheet views (frozen panes); `null` leaves the XML untouched. */
  views: SheetViewProps[] | null;
}

interface Op {
  pos: number;
  end: number;
  text: string;
  /**
   * Tiebreak for edits anchored at the same position (applied after sorting by
   * descending `pos`): lower `order` is applied first and therefore ends up
   * LATER in the output; the edit that must appear first in the document gets
   * the highest `order` at that position. Example at `<sheetData>`:
   * a created `<cols>` (order 1) must precede a created `<sheetViews>` text,
   * so the views edit gets order 2 and the cols edit order 1.
   */
  order?: number;
}

/**
 * Patch a worksheet: only the changed cells/merges/columns/rows/views are
 * rewritten; every untouched byte of the original XML stays exactly as it was.
 */
export function serializeSheet(
  sheet: ParsedSheet,
  patch: SheetPatch,
  date1904: boolean,
  registry: StyleRegistry | null,
): string {
  const ops: Op[] = [];
  const insertsByPos = new Map<number, { col: number; xml: string }[]>();
  const newRows = new Map<number, { col: number; xml: string }[]>();
  const selfClosingRows = new Map<number, { col: number; xml: string }[]>();

  for (const [ref, edit] of patch.changes) {
    const existing = sheet.cells.get(ref);
    const baseIndex = existing ? existing.styleIndex : 0;
    const styleIndex =
      edit.style && registry ? registry.resolveIndex(baseIndex, edit.style) : baseIndex;
    const value = edit.hasValue ? edit.value : existing ? existing.value : null;
    if (!edit.hasValue && styleIndex === baseIndex) {
      // Style-only edit that resolves to the cell's existing style (e.g. an
      // empty override, or a border/alignment identical to the base one):
      // nothing changed, keep the original bytes.
      continue;
    }
    let xml: string;
    if ((value === null || value === undefined) && edit.style && styleIndex) {
      // An empty cell that only carries styling still needs to exist.
      xml = `<c r="${ref}" s="${styleIndex}"/>`;
    } else {
      xml = generateCellXml(ref, value, styleIndex, date1904);
    }

    if (existing) {
      ops.push({ pos: existing.start, end: existing.end, text: xml });
      continue;
    }
    const { row, col } = refToCoords(ref);
    const rowRec = sheet.rows.get(row);
    if (rowRec && !rowRec.selfClosing) {
      const list = insertsByPos.get(rowRec.closeStart) ?? [];
      list.push({ col, xml });
      insertsByPos.set(rowRec.closeStart, list);
    } else if (rowRec && rowRec.selfClosing) {
      const list = selfClosingRows.get(row) ?? [];
      list.push({ col, xml });
      selfClosingRows.set(row, list);
    } else {
      const list = newRows.get(row) ?? [];
      list.push({ col, xml });
      newRows.set(row, list);
    }
  }

  for (const [pos, list] of insertsByPos) {
    list.sort((a, b) => a.col - b.col);
    ops.push({ pos, end: pos, text: list.map((entry) => entry.xml).join('') });
  }

  for (const [rowNum, list] of selfClosingRows) {
    list.sort((a, b) => a.col - b.col);
    const rowRec = sheet.rows.get(rowNum)!;
    const openTag = (patchedRowOpen(sheet, patch, rowNum) ?? sheet.xml.slice(rowRec.start, rowRec.openEnd))
      .replace(/\/>\s*$/, '>');
    ops.push({
      pos: rowRec.start,
      end: rowRec.end,
      text: openTag + list.map((entry) => entry.xml).join('') + '</row>',
    });
  }

  // Row heights on existing rows: patch only the opening tag. Self-closing
  // rows that received cells are handled by the loop above.
  for (const rowNum of patch.rowHeights.keys()) {
    const rowRec = sheet.rows.get(rowNum);
    if (!rowRec || selfClosingRows.has(rowNum)) continue;
    const open = patchedRowOpen(sheet, patch, rowNum)!;
    ops.push({
      pos: rowRec.start,
      end: rowRec.selfClosing ? rowRec.end : rowRec.openEnd,
      text: open,
    });
  }

  // New rows: merge pending heights into the opening tag; rows that only carry
  // a height are emitted as self-closing elements.
  const newRowNumbers = new Set<number>(newRows.keys());
  for (const rowNum of patch.rowHeights.keys()) {
    if (!sheet.rows.has(rowNum)) newRowNumbers.add(rowNum);
  }
  const newRowXml: string[] = [];
  for (const rowNum of [...newRowNumbers].sort((a, b) => a - b)) {
    const list = (newRows.get(rowNum) ?? []).sort((a, b) => a.col - b.col);
    const height = patch.rowHeights.get(rowNum);
    const attrs = height !== undefined && height !== null ? ` ht="${height}" customHeight="1"` : '';
    const cells = list.map((entry) => entry.xml).join('');
    newRowXml.push(cells ? `<row r="${rowNum}"${attrs}>${cells}</row>` : `<row r="${rowNum}"${attrs}/>`);
  }
  if (newRowXml.length > 0) {
    if (sheet.sheetDataSelfClosing) {
      ops.push({
        pos: sheet.sheetDataStart,
        end: sheet.sheetDataEnd,
        text: `<sheetData>${newRowXml.join('')}</sheetData>`,
      });
    } else {
      ops.push({ pos: sheet.sheetDataClose, end: sheet.sheetDataClose, text: newRowXml.join('') });
    }
  }

  applyColEdits(sheet, patch.colEdits, ops);
  applyViews(sheet, patch.views, ops);
  applyDimension(sheet, patch, ops);
  applyMergeEdits(sheet, patch.newMerges, patch.removedMerges, ops);

  ops.sort((a, b) => b.pos - a.pos || (a.order ?? 0) - (b.order ?? 0));
  let out = sheet.xml;
  for (const op of ops) out = out.slice(0, op.pos) + op.text + out.slice(op.end);
  return out;
}

/** Opening tag of a row with the pending height applied, or null when untouched. */
function patchedRowOpen(sheet: ParsedSheet, patch: SheetPatch, rowNum: number): string | null {
  const rowRec = sheet.rows.get(rowNum);
  if (!rowRec || !patch.rowHeights.has(rowNum)) return null;
  return patchRowOpenTag(sheet.xml.slice(rowRec.start, rowRec.openEnd), patch.rowHeights.get(rowNum)!);
}

function patchRowOpenTag(open: string, height: number | null): string {
  const selfClosing = open.endsWith('/>');
  const body = selfClosing ? open.slice(0, -2) : open.slice(0, -1);
  const close = selfClosing ? '/>' : '>';
  if (height === null || height === undefined) {
    return body.replace(/\s+ht="[^"]*"/g, '').replace(/\s+customHeight="[^"]*"/g, '') + close;
  }
  let out = body
    .replace(/\s+ht="[^"]*"/g, ` ht="${height}"`)
    .replace(/\s+customHeight="[^"]*"/g, ' customHeight="1"');
  if (!/ ht="/.test(out)) out += ` ht="${height}"`;
  if (!/ customHeight="/.test(out)) out += ' customHeight="1"';
  return out + close;
}

function applyColEdits(sheet: ParsedSheet, colEdits: Map<number, ColEdit>, ops: Op[]): void {
  if (colEdits.size === 0) return;
  const cols = sheet.cols;
  if (!cols) {
    const entries = [...colEdits.keys()]
      .sort((a, b) => a - b)
      .map((col) => colEntryXml(col, col, colEditAttrs({}, colEdits.get(col)!)));
    // cols precedes sheetData in the worksheet schema.
    ops.push({ pos: sheet.sheetDataStart, end: sheet.sheetDataStart, text: `<cols>${entries.join('')}</cols>`, order: 1 });
    return;
  }

  // Split and regenerate only the entries that cover an edited column.
  const editedEntries = new Map<ColRecord, number[]>();
  const insertions = new Map<number, { col: number; xml: string }[]>();
  for (const [col, edit] of colEdits) {
    const entry = cols.entries.find((e) => col >= e.min && col <= e.max);
    if (entry) {
      const list = editedEntries.get(entry) ?? [];
      list.push(col);
      editedEntries.set(entry, list);
    } else {
      const next = cols.entries.find((e) => e.min > col);
      const pos = next ? next.start : cols.closeStart;
      const list = insertions.get(pos) ?? [];
      list.push({ col, xml: colEntryXml(col, col, colEditAttrs({}, edit)) });
      insertions.set(pos, list);
    }
  }

  for (const [entry, editedCols] of editedEntries) {
    ops.push({ pos: entry.start, end: entry.end, text: splitColEntry(entry, editedCols, colEdits) });
  }
  for (const [pos, list] of insertions) {
    list.sort((a, b) => a.col - b.col);
    ops.push({ pos, end: pos, text: list.map((entry) => entry.xml).join(''), order: 1 });
  }
}

/** Apply width/hidden overrides onto (a copy of) an existing attribute set. */
function colEditAttrs(base: Record<string, string>, edit: ColEdit): Record<string, string> {
  const attrs = { ...base };
  if (edit.width !== undefined) {
    if (edit.width === null) {
      delete attrs.width;
      delete attrs.customWidth;
    } else {
      attrs.width = String(edit.width);
      attrs.customWidth = '1';
    }
  }
  if (edit.hidden !== undefined) {
    if (edit.hidden === null || edit.hidden === false) delete attrs.hidden;
    else attrs.hidden = '1';
  }
  return attrs;
}

function colEntryXml(min: number, max: number, attrs: Record<string, string>): string {
  const parts = [`min="${min}"`, `max="${max}"`];
  for (const [key, value] of Object.entries(attrs)) {
    if (key === 'min' || key === 'max' || value === undefined) continue;
    parts.push(`${key}="${escapeXml(value)}"`);
  }
  return `<col ${parts.join(' ')}/>`;
}

/**
 * Rewrite one `<col>` range so the edited columns carry their overrides while
 * every other column in the range keeps its original attributes untouched.
 */
function splitColEntry(entry: ColRecord, editedCols: number[], colEdits: Map<number, ColEdit>): string {
  let segments: { from: number; to: number }[] = [{ from: entry.min, to: entry.max }];
  for (const col of editedCols) {
    const next: { from: number; to: number }[] = [];
    for (const seg of segments) {
      if (col < seg.from || col > seg.to) {
        next.push(seg);
        continue;
      }
      if (col > seg.from) next.push({ from: seg.from, to: col - 1 });
      next.push({ from: col, to: col });
      if (col < seg.to) next.push({ from: col + 1, to: seg.to });
    }
    segments = next;
  }

  let xml = '';
  for (const seg of segments) {
    const edited = seg.from === seg.to && colEdits.has(seg.from);
    const attrs = colEditAttrs(entry.attrs, edited ? colEdits.get(seg.from)! : {});
    xml += colEntryXml(seg.from, seg.to, attrs);
  }
  return xml;
}

/** The `<pane>` element for the requested first view, or null when unfrozen. */
/** Patch the `<sheetViews>/<sheetView>/<pane>` triple of a sheet. */
function applyViews(sheet: ParsedSheet, views: SheetViewProps[] | null, ops: Op[]): void {
  if (!views) return;
  const paneXml = paneXmlFor(views[0]);
  const record = sheet.sheetViews;
  if (record) {
    const view = record.views[0];
    if (!view) {
      // Container without a sheetView child: put one (with the pane) inside.
      if (!paneXml) return;
      const inner = `<sheetView workbookViewId="0">${paneXml}</sheetView>`;
      if (record.closeStart === record.openEnd) {
        ops.push({ pos: record.start, end: record.end, text: `<sheetViews>${inner}</sheetViews>` });
      } else {
        ops.push({ pos: record.closeStart, end: record.closeStart, text: inner });
      }
      return;
    }
    if (view.pane) {
      // Replace or drop the existing pane; everything else stays byte-identical.
      ops.push({ pos: view.pane.start, end: view.pane.end, text: paneXml ?? '' });
    } else if (paneXml) {
      // pane is the first child of sheetView (before selection etc.).
      if (view.closeStart === view.openEnd) {
        // Self-closing sheetView: reopen it with the pane inside.
        const attrs = Object.entries(view.attrs)
          .map(([key, value]) => `${key}="${escapeXml(value)}"`)
          .join(' ');
        ops.push({ pos: view.start, end: view.end, text: `<sheetView ${attrs}>${paneXml}</sheetView>` });
      } else {
        ops.push({ pos: view.openEnd, end: view.openEnd, text: paneXml, order: 1 });
      }
    }
  } else if (paneXml) {
    // sheetViews precedes sheetFormatPr/cols/sheetData in the worksheet schema.
    const pos = sheet.sheetFormatPrStart ?? sheet.cols?.start ?? sheet.sheetDataStart;
    ops.push({
      pos,
      end: pos,
      text: `<sheetViews><sheetView workbookViewId="0">${paneXml}</sheetView></sheetViews>`,
      order: 2,
    });
  }
}

/**
 * Keep `<dimension>` in sync when new cells extend the used range: SheetJS and
 * other consumers bound their reads by the declared ref, so a stale dimension
 * would make the added cells invisible to them.
 */
function applyDimension(sheet: ParsedSheet, patch: SheetPatch, ops: Op[]): void {
  const dim = sheet.dimension;
  if (!dim) return;
  const match = /^([A-Za-z]+)(\d+)(?::([A-Za-z]+)(\d+))?$/.exec(dim.ref);
  if (!match) return;
  let maxRow = sheet.maxRow;
  let maxCol = sheet.maxCol;
  for (const ref of patch.changes.keys()) {
    const { row, col } = refToCoords(ref);
    if (row > maxRow) maxRow = row;
    if (col > maxCol) maxCol = col;
  }
  const endRow = Math.max(Number(match[2]), maxRow);
  const endCol = Math.max(refToCoords(`${match[3] ?? match[1]}${match[4] ?? match[2]}`).col, maxCol);
  const startRef = `${match[1]}${match[2]}`;
  if (endRow > Number(match[4] ?? match[2]) || endCol > refToCoords(`${match[3] ?? match[1]}${match[4] ?? match[2]}`).col) {
    ops.push({
      pos: dim.start,
      end: dim.end,
      text: `<dimension ref="${startRef}:${coordsToRef(endRow, endCol)}"/>`,
    });
  }
}

function applyMergeEdits(sheet: ParsedSheet, newMerges: string[], removedMerges: string[], ops: Op[]): void {
  const container = sheet.mergeCells;
  let present = sheet.mergeChildren;

  const removedSet = new Set(removedMerges);
  if (removedMerges.length > 0) {
    for (const child of present) {
      if (removedSet.has(child.ref)) ops.push({ pos: child.start, end: child.end, text: '' });
    }
    present = present.filter((child) => !removedSet.has(child.ref));
  }

  const hasAdd = newMerges.length > 0;
  // Validated here (not when `mergeCells()` was called) so the eager writer
  // rejects exactly what the incremental one rejects, with the same code.
  const additions = newMerges
    .map((ref) => `<mergeCell ref="${escapeXml(assertMergeRange(ref))}"/>`)
    .join('');

  if (container) {
    if (!hasAdd && removedMerges.length === 0) return;
    const count = present.length + newMerges.length;
    if (container.closeStart === container.openEnd) {
      // Self-closing <mergeCells/>: reopen it so the additions land INSIDE the
      // element (inserting after `/>` would leave them as siblings).
      if (!hasAdd) return; // no children to remove, nothing to add
      const body = sheet.xml.slice(container.start, container.openEnd).replace(/\/>$/, '>');
      const withCount = /\bcount="\d+"/.test(body)
        ? body.replace(/\bcount="\d+"/, `count="${count}"`)
        : body.replace(/>$/, ` count="${count}">`);
      ops.push({ pos: container.start, end: container.end, text: `${withCount}${additions}</mergeCells>` });
      return;
    }
    const open = sheet.xml.slice(container.start, container.openEnd);
    const updated = /\bcount="\d+"/.test(open)
      ? open.replace(/\bcount="\d+"/, `count="${count}"`)
      : open.replace(/>$/, ` count="${count}">`);
    ops.push({ pos: container.start, end: container.openEnd, text: updated });
    // Additions go at the end of the container; removals are separate range ops.
    if (hasAdd) ops.push({ pos: container.closeStart, end: container.closeStart, text: additions });
  } else if (hasAdd) {
    // mergeCells is a sibling of sheetData, so it goes AFTER </sheetData>.
    ops.push({
      pos: sheet.sheetDataEnd,
      end: sheet.sheetDataEnd,
      text: `<mergeCells count="${newMerges.length}">${additions}</mergeCells>`,
    });
  }
}