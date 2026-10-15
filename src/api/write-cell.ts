/**
 * A cell of a row being written.
 *
 * Mirrors the read-side `Cell` property names (`value`, `numFmt`, `font`,
 * `fill`, `border`, `alignment`) so code can be shared between reading and
 * writing. Style setters accumulate into one `CellStyleOverrides` patch, which
 * is merged with the row's own patch and resolved against the workbook's style
 * registry when the row is flushed — identical styles therefore collapse to a
 * single `styles.xml` entry.
 *
 * The cell must not be kept after the next `addRow()` (or after the workbook is
 * committed): its content has already been serialized and flushed (`ROW_WRITTEN`).
 */
import { coordsToRef } from '../model/refs';
import { mergeStylePatch, type CellStyleOverrides } from '../model/style';
import type { CellValue } from '../model/value';
import { SheetLiteError } from '../zip/errors';
import { assertWritableValue } from '../writer/sheet-stream-writer';
import type { WriteRow } from './write-worksheet';

export class WriteCell {
  private value_: CellValue = null;
  private style: CellStyleOverrides = {};

  constructor(
    private readonly owner: WriteRow,
    readonly col: number,
    value: CellValue,
  ) {
    this.value = value;
  }

  /** Reference of the cell, e.g. "B7". */
  get ref(): string {
    return coordsToRef(this.owner.number, this.col);
  }

  get row(): number {
    return this.owner.number;
  }

  get value(): CellValue {
    return this.value_;
  }

  set value(value: CellValue) {
    this.assertOpen();
    assertWritableValue(value, this.ref);
    this.value_ = value;
  }

  /** Style overrides of this cell alone (merged with the row's on flush). */
  get cellStyle(): CellStyleOverrides {
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

  toString(): string {
    const value = this.value;
    return value === null || value === undefined ? '' : String(value);
  }

  /** Merge one property over the cell's patch; `undefined` clears it. */
  private setProp<K extends keyof CellStyleOverrides>(key: K, value: CellStyleOverrides[K]): void {
    this.assertOpen();
    const next = value === undefined ? { ...this.style } : mergeStylePatch(this.style, { [key]: value } as CellStyleOverrides);
    if (value === undefined) delete next[key];
    this.style = next;
  }

  private assertOpen(): void {
    if (this.owner.flushed) {
      throw new SheetLiteError(
        `Cannot change ${this.ref} after its row has been written`,
        'ROW_WRITTEN',
      );
    }
  }
}
