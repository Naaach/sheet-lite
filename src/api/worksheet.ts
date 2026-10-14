import type { PackageMeta } from '../reader/package-reader';
import { coordsToRef, parseSheet, refToCoords, type ParsedSheet } from '../reader/sheet-parser';
import { serializeSheet, type CellEdit } from '../writer/sheet-writer';
import type { StyleRegistry } from '../writer/style-registry';
import type { CellStyleOverrides, FillProps, FontProps } from '../model/style';
import { type CellValue } from '../model/value';
import { Cell } from './cell';

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
}

/**
 * A single sheet. Reading values and patching cells, styles and merges while
 * preserving everything else in the sheet XML.
 */
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

  // --- internal ---

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
    this.changes.set(key, { ...existing, style });
  }

  isDirty(): boolean {
    return this.changes.size > 0 || this.newMerges.size > 0 || this.removedMerges.size > 0;
  }

  serialize(registry: StyleRegistry | null): string {
    return serializeSheet(
      this.parsed,
      this.changes,
      this.meta.date1904,
      registry,
      [...this.newMerges],
      [...this.removedMerges],
    );
  }

  private baseStyleIndex(ref: string): number {
    return this.parsed.cells.get(ref)?.styleIndex ?? 0;
  }
}
