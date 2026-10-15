/**
 * A worksheet being created, written one row at a time (ExcelJS-shaped).
 *
 * `addRow()` appends a row; the row's XML is deflated and streamed out when
 * the NEXT row is added (or when the workbook is committed), so at most one row
 * is alive at a time and the sheet's XML never exists as a whole. That one-row
 * window is also why styles and the row height of the most recent row can still
 * be changed after `addRow()` returns.
 */
import { mergeStylePatch, type CellStyleOverrides } from '../model/style';
import type { SheetViewProps } from '../model/view';
import type { CellValue } from '../model/value';
import { SheetLiteError } from '../zip/errors';
import type { StreamingZipWriter } from '../zip/zip-stream';
import {
  MAX_COLUMNS,
  MAX_ROWS,
  SheetStreamWriter,
  streamCellXml,
  type ColumnDef,
} from '../writer/sheet-stream-writer';
import { WriteCell } from './write-cell';
import { WriteColumn } from './write-column';

/** Services a worksheet needs from its workbook. */
export interface WriteContext {
  /**
   * Make this sheet the sheet currently being written, closing the entry of the
   * previous one. Throws `SHEET_INTERLEAVED` when this sheet already has rows:
   * a ZIP entry's bytes must be contiguous, so rows have to be appended to one
   * sheet at a time.
   */
  activate(sheet: WriteWorksheet): void;
  /** Resolve a style patch to a `styles.xml` `cellXfs` index (0 = default). */
  resolveStyle(patch: CellStyleOverrides): number;
  /** Serial dates are written with this date system (1900 unless asked otherwise). */
  readonly date1904: boolean;
  /** True once the workbook has been committed. */
  readonly sealed: boolean;
}

/** One appended row: its cells, its height and the styles it applies to them. */
export class WriteRow {
  private written = false;
  private readonly cells = new Map<number, WriteCell>();
  private readonly style: CellStyleOverrides = {};
  private heightValue: number | undefined;

  constructor(
    readonly worksheet: WriteWorksheet,
    readonly number: number,
    values: readonly CellValue[],
  ) {
    values.forEach((value, index) => {
      this.getCell(index + 1).value = value;
    });
  }

  /** Cells written so far, in column order. */
  get cellList(): WriteCell[] {
    return [...this.cells.values()].sort((a, b) => a.col - b.col);
  }

  /** Get (and create if needed) the cell at a 1-based column index. */
  getCell(col: number): WriteCell {
    assertColumn(col);
    let cell = this.cells.get(col);
    if (!cell) {
      cell = new WriteCell(this, col, null);
      this.cells.set(col, cell);
    }
    return cell;
  }

  eachCell(callback: (cell: WriteCell, col: number) => void): void {
    for (const cell of this.cellList) callback(cell, cell.col);
  }

  get height(): number | undefined {
    return this.heightValue;
  }

  set height(value: number | undefined) {
    if (value !== undefined && (!Number.isFinite(value) || value < 0)) {
      throw new SheetLiteError(
        `Invalid row height ${value}: must be a non-negative number`,
        'BAD_ROW_HEIGHT',
      );
    }
    this.assertOpen();
    this.heightValue = value;
  }

  /** Row-wide style patch, applied to every cell that does not override it. */
  get rowStyle(): CellStyleOverrides {
    return this.style;
  }

  get numFmt(): string | undefined {
    return this.style.numFmt;
  }

  set numFmt(value: string | undefined) {
    this.setProp('numFmt', value);
  }

  get font(): CellStyleOverrides['font'] {
    return this.style.font;
  }

  set font(value: CellStyleOverrides['font']) {
    this.setProp('font', value);
  }

  get fill(): CellStyleOverrides['fill'] {
    return this.style.fill;
  }

  set fill(value: CellStyleOverrides['fill']) {
    this.setProp('fill', value);
  }

  get border(): CellStyleOverrides['border'] {
    return this.style.border;
  }

  set border(value: CellStyleOverrides['border']) {
    this.setProp('border', value);
  }

  get alignment(): CellStyleOverrides['alignment'] {
    return this.style.alignment;
  }

  set alignment(value: CellStyleOverrides['alignment']) {
    this.setProp('alignment', value);
  }

  /** True once the row has been streamed into the sheet part. */
  get flushed(): boolean {
    return this.written;
  }

  /** @internal Serialize the row's cells (`<c>` markup, in column order). */
  toCellsXml(context: WriteContext): string {
    let xml = '';
    for (const cell of this.cellList) {
      const patch = mergeStylePatch(this.style, cell.cellStyle);
      const styleIndex = context.resolveStyle(patch);
      xml += streamCellXml(cell.ref, cell.value, styleIndex, context.date1904);
    }
    return xml;
  }

  /** @internal Mark the row as serialized. */
  markWritten(): void {
    this.written = true;
  }

  /** Merge one property over the row's patch; `undefined` clears it. */
  private setProp<K extends keyof CellStyleOverrides>(key: K, value: CellStyleOverrides[K]): void {
    this.assertOpen();
    const next = value === undefined ? { ...this.style } : mergeStylePatch(this.style, { [key]: value } as CellStyleOverrides);
    if (value === undefined) delete next[key];
    Object.assign(this.style, next);
  }

  private assertOpen(): void {
    if (this.written) {
      throw new SheetLiteError(
        `Cannot change row ${this.number} after it has been written`,
        'ROW_WRITTEN',
      );
    }
  }
}

/** A sheet being written: add rows, set column widths, merges and frozen panes. */
export class WriteWorksheet {
  readonly name: string;
  readonly id: number;
  /** Store key of the sheet part, e.g. "xl/worksheets/sheet1.xml". */
  readonly partName: string;

  private readonly stream: SheetStreamWriter;
  private readonly columnDefs = new Map<number, ColumnDef>();
  private viewsValue: SheetViewProps[] | null;
  private rows = 0;
  private pending: WriteRow | null = null;
  private finalized = false;

  constructor(
    name: string,
    id: number,
    partName: string,
    zip: StreamingZipWriter,
    date1904: boolean,
    private readonly context: WriteContext,
    views?: SheetViewProps[] | null,
  ) {
    this.name = name;
    this.id = id;
    this.partName = partName;
    this.viewsValue = views ?? null;
    this.stream = new SheetStreamWriter(zip, partName, date1904);
    if (this.viewsValue) this.stream.setViews(this.viewsValue);
  }

  /** Number of rows appended so far. */
  get rowCount(): number {
    return this.rows;
  }

  /** True once this sheet has been closed (it can no longer be written to). */
  get done(): boolean {
    return this.finalized;
  }

  /**
   * Append a row: `addRow(['name', 42, new Date()])`. Values are written to
   * columns A, B, C … in order; `null`/`undefined` and `undefined` holes leave
   * the cell empty (an empty row writes `<row r="N"/>`).
   *
   * The returned row can still be styled — and given a height — until the next
   * `addRow()` or until the workbook is committed.
   *
   * @throws `SHEET_FINALIZED` after `commit()`, `BAD_ROW` past row 1 048 576
   * and `BAD_COLUMN` past column XFD.
   */
  addRow(values: readonly CellValue[] = []): WriteRow {
    this.assertWritable();
    this.context.activate(this);
    this.flush();
    if (values.length > MAX_COLUMNS) {
      throw new SheetLiteError(
        `Cannot write ${values.length} cells in row ${this.rows + 1}: a row holds at most ${MAX_COLUMNS} columns`,
        'BAD_COLUMN',
      );
    }
    const row = new WriteRow(this, this.rows + 1, values);
    this.rows += 1;
    this.pending = row;
    return row;
  }

  /** Merge a range, e.g. `mergeCells('A1:C1')`. Applied when the part is closed. */
  mergeCells(ref: string): void {
    this.assertWritable();
    this.stream.mergeCells(ref);
  }

  /** A single column (1-based number or letter), ExcelJS-style. */
  getColumn(col: number | string): WriteColumn {
    return new WriteColumn(this, col);
  }

  /** The frozen-pane views requested for this sheet. */
  get views(): SheetViewProps[] {
    return this.viewsValue ? this.viewsValue.map((view) => ({ ...view })) : [];
  }

  /**
   * ExcelJS-style frozen panes. Must be set before the first row is written
   * (`PROLOGUE_LOCKED` afterwards): `<sheetViews>` precedes `<sheetData>`.
   *
   * The field is updated only once the stream has accepted the value, so
   * `views` always reports what was actually written, never a rejected set.
   */
  set views(views: SheetViewProps[] | undefined) {
    const next = (views ?? []).length > 0 ? views!.map((view) => ({ ...view })) : null;
    this.stream.setViews(next); // throws PROLOGUE_LOCKED once the part has started
    this.viewsValue = next;
  }

  /** @internal Column definition of the prologue. */
  getColumnDef(col: number): ColumnDef {
    return this.columnDefs.get(col) ?? {};
  }

  /** @internal Set a column's width / visibility in the prologue. */
  setColumnDef(col: number, def: ColumnDef): void {
    assertColumn(col);
    this.stream.setColumn(col, def);
    const merged: ColumnDef = { ...this.getColumnDef(col), ...def };
    for (const key of Object.keys(merged) as Array<keyof ColumnDef>) {
      if (merged[key] === undefined) delete merged[key];
    }
    this.columnDefs.set(col, merged);
  }

  /** @internal Flush the pending row (serialize it into the sheet part). */
  flush(): void {
    const row = this.pending;
    if (!row || row.flushed) return;
    this.stream.writeRow(row.number, row.height, row.toCellsXml(this.context));
    row.markWritten();
    this.pending = null;
  }

  /** @internal Close the sheet part (epilogue + ZIP entry). */
  finalize(): void {
    if (this.finalized) return;
    this.flush();
    this.stream.finish();
    this.finalized = true;
  }

  /** @internal Reject writes past the sheet limit or after `commit()`. */
  assertWritable(): void {
    if (this.finalized || this.context.sealed) {
      throw new SheetLiteError(
        `Cannot write to worksheet "${this.name}": its output has already been written`,
        'SHEET_FINALIZED',
      );
    }
    if (this.rows >= MAX_ROWS) {
      throw new SheetLiteError(
        `Cannot add a row to worksheet "${this.name}": a sheet holds at most ${MAX_ROWS} rows`,
        'BAD_ROW',
      );
    }
  }
}

/** @internal Column bound check shared by `WriteWorksheet` and `WriteRow`. */
function assertColumn(col: number): void {
  if (!Number.isInteger(col) || col < 1 || col > MAX_COLUMNS) {
    throw new SheetLiteError(
      `Invalid column ${col}: must be an integer 1-${MAX_COLUMNS}`,
      'BAD_COLUMN',
    );
  }
}