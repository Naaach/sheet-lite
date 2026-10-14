import type { PackageMeta } from '../reader/package-reader';
import { coordsToRef, parseSheet, refToCoords, type ParsedSheet } from '../reader/sheet-parser';
import { serializeSheet } from '../writer/sheet-writer';
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
 * A single sheet. v0.2 supports reading values and *patching* cells/rows while
 * preserving everything else in the sheet XML.
 */
export class Worksheet {
  readonly name: string;
  readonly id: number;
  /** Store key of the sheet part (no leading slash), e.g. "xl/worksheets/sheet1.xml". */
  readonly partName: string;

  private readonly parsed: ParsedSheet;
  private readonly meta: PackageMeta;
  private readonly changes = new Map<string, CellValue>();

  constructor(name: string, id: number, partName: string, xml: string, meta: PackageMeta) {
    this.name = name;
    this.id = id;
    this.partName = partName;
    this.meta = meta;
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
    const rowNumbers = new Set<number>();
    for (const record of this.parsed.cells.values()) rowNumbers.add(record.row);
    for (const ref of this.changes.keys()) rowNumbers.add(refToCoords(ref).row);
    for (const n of [...rowNumbers].sort((x, y) => x - y)) {
      callback(this.getRow(n), n);
    }
  }

  /** Append a row of values (column A, B, C, …). */
  addRow(values: CellValue[]): Row {
    const rowNumber = this.rowCount + 1;
    values.forEach((value, index) => {
      this.setCellValue(coordsToRef(rowNumber, index + 1), value);
    });
    return this.getRow(rowNumber);
  }

  // --- internal ---

  getCellValue(ref: string): CellValue {
    if (this.changes.has(ref)) return this.changes.get(ref)!;
    return this.parsed.cells.get(ref)?.value ?? null;
  }

  setCellValue(ref: string, value: CellValue): void {
    this.changes.set(ref.toUpperCase(), value);
  }

  isDirty(): boolean {
    return this.changes.size > 0;
  }

  serialize(): string {
    return serializeSheet(this.parsed, this.changes, this.meta.date1904);
  }
}
