import type { PackageMeta } from '../reader/package-reader';
import {
  columnToNumber,
  coordsToRef,
  parseSheet,
  refToCoords,
  type ParsedSheet,
} from '../reader/sheet-parser';
import { serializeSheet, type CellEdit, type ColEdit, type SheetPatch } from '../writer/sheet-writer';
import type { StyleRegistry } from '../writer/style-registry';
import type {
  AlignmentProps,
  BorderProps,
  CellStyleOverrides,
  FillProps,
  FontProps,
} from '../model/style';
import type { SheetViewProps } from '../model/view';
import { type CellValue } from '../model/value';
import { SheetLiteError } from '../zip/errors';
import { Cell } from './cell';
import { Column } from './column';

export class Row {
  constructor(
    readonly worksheet: Worksheet,
    readonly number: number,
    readonly cells: Cell[],
  ) {}

  getCell(col: number): Cell {
    return this.cells.find((c) => c.col === col) ?? new Cell(this.worksheet, coordsToRef(this.number, col));
  }

  eachCell(callback: (cell: Cell, col: number) => void): void {
    for (const cell of this.cells) callback(cell, cell.col);
  }

  get height(): number | undefined {
    return this.worksheet.getRowHeight(this.number);
  }

  set height(value: number | undefined) {
    this.worksheet.setRowHeight(this.number, value);
  }
}

/**
 * A single sheet. Reading values and patching cells, styles and merges while
 * preserving everything else in the sheet XML.
 */
const hasKeys = (value: unknown): boolean =>
  typeof value === 'object' && value !== null && Object.keys(value).length > 0;

export class Worksheet {
  readonly name: string;
  readonly id: number;
  /** Store key of the sheet part (no leading slash), e.g. "xl/worksheets/sheet1.xml". */
  readonly partName: string;

  private readonly parsed: ParsedSheet;
  private readonly meta: PackageMeta;
  private readonly registry: StyleRegistry | null;
  private readonly changes = new Map<string, CellEdit>();
  private readonly newMerges = new Set<string>();
  private readonly removedMerges = new Set<string>();
  private readonly colEdits = new Map<number, ColEdit>();
  private readonly rowHeights = new Map<number, number | null>();
  private pendingViews: SheetViewProps[] | null = null;

  constructor(
    name: string,
    id: number,
    partName: string,
    xml: string,
    meta: PackageMeta,
    registry: StyleRegistry | null,
  ) {
    this.name = name;
    this.id = id;
    this.partName = partName;
    this.meta = meta;
    this.registry = registry;
    this.parsed = parseSheet(xml, meta);
  }

  get rowCount(): number {
    let max = this.parsed.maxRow;
    for (const ref of this.changes.keys()) {
      const { row } = refToCoords(ref);
      if (row > max) max = row;
    }
    return max;
  }

  get columnCount(): number {
    let max = this.parsed.maxCol;
    for (const ref of this.changes.keys()) {
      const { col } = refToCoords(ref);
      if (col > max) max = col;
    }
    return max;
  }

  getCell(ref: string): Cell;
  getCell(row: number, col: number): Cell;
  getCell(a: string | number, b?: number): Cell {
    const ref = typeof a === 'number' ? coordsToRef(a, b!) : a.toUpperCase();
    return new Cell(this, ref);
  }

  getRow(rowNumber: number): Row {
    const cells: Cell[] = [];
    for (const record of this.parsed.cells.values()) {
      if (record.row === rowNumber) cells.push(new Cell(this, record.ref));
    }
    for (const ref of this.changes.keys()) {
      if (refToCoords(ref).row === rowNumber && !this.parsed.cells.has(ref)) {
        cells.push(new Cell(this, ref));
      }
    }
    cells.sort((x, y) => x.col - y.col);
    return new Row(this, rowNumber, cells);
  }

  eachRow(callback: (row: Row, rowNumber: number) => void): void {
    // Group cells by row once (O(cells)) instead of re-scanning the whole cell
    // map for every row, which made iteration O(rows x cells).
    const byRow = new Map<number, Cell[]>();
    for (const record of this.parsed.cells.values()) {
      let list = byRow.get(record.row);
      if (!list) {
        list = [];
        byRow.set(record.row, list);
      }
      list.push(new Cell(this, record.ref));
    }
    for (const ref of this.changes.keys()) {
      if (this.parsed.cells.has(ref)) continue;
      const { row } = refToCoords(ref);
      let list = byRow.get(row);
      if (!list) {
        list = [];
        byRow.set(row, list);
      }
      list.push(new Cell(this, ref));
    }
    for (const n of [...byRow.keys()].sort((x, y) => x - y)) {
      const cells = byRow.get(n)!.sort((x, y) => x.col - y.col);
      callback(new Row(this, n, cells), n);
    }
  }

  /** Append a row of values (column A, B, C, …). */
  addRow(values: CellValue[]): Row {
    const rowNumber = this.rowCount + 1;
    values.forEach((value, index) => this.setCellValue(coordsToRef(rowNumber, index + 1), value));
    return this.getRow(rowNumber);
  }

  /** Merge a range, e.g. "A1:C1". */
  mergeCells(ref: string): void {
    this.newMerges.add(ref.toUpperCase());
    this.removedMerges.delete(ref.toUpperCase());
  }

  /** Remove a merge by its range reference. */
  unmergeCells(ref: string): void {
    const key = ref.toUpperCase();
    this.newMerges.delete(key);
    this.removedMerges.add(key);
  }

  /** A single column (1-based number or letter), ExcelJS-style. */
  getColumn(col: number | string): Column {
    const n = typeof col === 'number' ? Math.floor(col) : columnToNumber(col);
    if (!Number.isFinite(n) || n < 1 || n > 16384) {
      throw new SheetLiteError(`Invalid column ${String(col)}: must be a number 1-16384 or a letter (A-XFD)`, 'BAD_COLUMN');
    }
    return new Column(this, n);
  }

  get views(): SheetViewProps[] {
    if (this.pendingViews) return this.pendingViews.map((view) => ({ ...view }));
    return readSheetViews(this.parsed);
  }

  /** ExcelJS-style frozen panes: `ws.views = [{ state: 'frozen', xSplit: 2, ySplit: 3 }]`. */
  set views(views: SheetViewProps[] | undefined) {
    this.pendingViews = (views ?? []).map((view) => ({ ...view }));
  }

  // --- internal ---

  getColumnWidth(col: number): number | undefined {
    for (const entry of this.parsed.cols?.entries ?? []) {
      if (col >= entry.min && col <= entry.max) {
        const width = entry.attrs.width;
        return width !== undefined ? Number(width) : undefined;
      }
    }
    return undefined;
  }

  setColumnWidth(col: number, width: number | undefined): void {
    // Excel's own range is 0-255: anything else (NaN, negative, non-finite,
    // > 255) is written straight into `<col width="..">` and makes Excel
    // offer to repair the file. `undefined` still clears the width.
    if (width !== undefined && !(width >= 0 && width <= 255)) {
      throw new SheetLiteError(`Invalid column width ${width}: must be a number 0-255`, 'BAD_COLUMN_WIDTH');
    }
    this.mergeColEdit(col, { width: width === undefined ? null : width });
  }

  getColumnHidden(col: number): boolean {
    for (const entry of this.parsed.cols?.entries ?? []) {
      if (col >= entry.min && col <= entry.max) {
        return entry.attrs.hidden === '1' || entry.attrs.hidden === 'true';
      }
    }
    return false;
  }

  setColumnHidden(col: number, hidden: boolean): void {
    this.mergeColEdit(col, { hidden: hidden ? true : null });
  }

  private mergeColEdit(col: number, edit: ColEdit): void {
    const existing = this.colEdits.get(col) ?? {};
    this.colEdits.set(col, { ...existing, ...edit });
  }

  getRowHeight(row: number): number | undefined {
    return this.parsed.rows.get(row)?.height;
  }

  setRowHeight(row: number, height: number | undefined): void {
    if (height !== undefined && (!Number.isFinite(height) || height < 0)) {
      throw new SheetLiteError(`Invalid row height ${height}: must be a non-negative number`, 'BAD_ROW_HEIGHT');
    }
    this.rowHeights.set(row, height === undefined ? null : height);
  }

  getCellValue(ref: string): CellValue {
    const edit = this.changes.get(ref);
    if (edit?.hasValue) return edit.value;
    return this.parsed.cells.get(ref)?.value ?? null;
  }

  getCellNumFmt(ref: string): string {
    const edit = this.changes.get(ref);
    if (edit?.style?.numFmt !== undefined) return edit.style.numFmt;
    if (!this.registry) return 'General';
    return this.registry.numFmtCode(this.baseStyleIndex(ref));
  }

  getCellFont(ref: string): FontProps {
    const base = this.registry ? this.registry.fontProps(this.baseStyleIndex(ref)) ?? {} : {};
    const edit = this.changes.get(ref);
    return edit?.style?.font ? { ...base, ...edit.style.font } : base;
  }

  getCellFill(ref: string): FillProps | undefined {
    const base = this.registry ? this.registry.fillProps(this.baseStyleIndex(ref)) : undefined;
    const edit = this.changes.get(ref);
    return edit?.style?.fill ? { ...base, ...edit.style.fill } : base;
  }

  getCellBorder(ref: string): BorderProps {
    const base = this.registry ? this.registry.borderProps(this.baseStyleIndex(ref)) ?? {} : {};
    const edit = this.changes.get(ref);
    return edit?.style?.border ? { ...base, ...edit.style.border } : base;
  }

  getCellAlignment(ref: string): AlignmentProps {
    const base = this.registry ? this.registry.alignmentProps(this.baseStyleIndex(ref)) ?? {} : {};
    const edit = this.changes.get(ref);
    return edit?.style?.alignment ? { ...base, ...edit.style.alignment } : base;
  }

  setCellValue(ref: string, value: CellValue): void {
    const key = ref.toUpperCase();
    const existing = this.changes.get(key) ?? { hasValue: false, value: null };
    this.changes.set(key, { ...existing, hasValue: true, value });
  }

  setCellStyle(ref: string, patch: CellStyleOverrides): void {
    const key = ref.toUpperCase();
    const existing = this.changes.get(key) ?? { hasValue: false, value: null };
    const style: CellStyleOverrides = { ...existing.style };
    if (patch.numFmt !== undefined) style.numFmt = patch.numFmt;
    if (patch.font) style.font = { ...style.font, ...patch.font };
    if (patch.fill) style.fill = { ...style.fill, ...patch.fill };
    if (patch.border) style.border = { ...style.border, ...patch.border };
    if (patch.alignment) style.alignment = { ...style.alignment, ...patch.alignment };
    if (
      !existing.hasValue &&
      style.numFmt === undefined &&
      !hasKeys(style.font) &&
      !hasKeys(style.fill) &&
      !hasKeys(style.border) &&
      !hasKeys(style.alignment)
    ) {
      // A style-only edit with no effective overrides: don't dirty the sheet.
      if (existing.style !== undefined) this.changes.delete(key);
      return;
    }
    this.changes.set(key, { ...existing, style });
  }

  isDirty(): boolean {
    return (
      this.changes.size > 0 ||
      this.newMerges.size > 0 ||
      this.removedMerges.size > 0 ||
      this.colEdits.size > 0 ||
      this.rowHeights.size > 0 ||
      this.pendingViews !== null
    );
  }

  serialize(registry: StyleRegistry | null): string {
    const patch: SheetPatch = {
      changes: this.changes,
      newMerges: [...this.newMerges],
      removedMerges: [...this.removedMerges],
      colEdits: this.colEdits,
      rowHeights: this.rowHeights,
      views: this.pendingViews,
    };
    return serializeSheet(this.parsed, patch, this.meta.date1904, registry);
  }

  private baseStyleIndex(ref: string): number {
    return this.parsed.cells.get(ref)?.styleIndex ?? 0;
  }
}

function readSheetViews(sheet: ParsedSheet): SheetViewProps[] {
  const record = sheet.sheetViews;
  if (!record) return [];
  return record.views.map((view) => {
    const props: SheetViewProps = { workbookViewId: Number(view.attrs.workbookViewId ?? 0) };
    if (view.pane) {
      const pane = view.pane.attrs;
      props.state = (pane.state as SheetViewProps['state'] | undefined) ?? 'split';
      if (pane.xSplit !== undefined) props.xSplit = Number(pane.xSplit);
      if (pane.ySplit !== undefined) props.ySplit = Number(pane.ySplit);
      if (pane.topLeftCell !== undefined) props.topLeftCell = pane.topLeftCell;
    } else {
      props.state = 'normal';
      if (view.attrs.topLeftCell !== undefined) props.topLeftCell = view.attrs.topLeftCell;
    }
    return props;
  });
}
