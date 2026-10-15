/**
 * Incremental serializer for ONE worksheet part.
 *
 * The eager writer (`writer/sheet-writer.ts`) patches an existing sheet XML in
 * memory; this one *produces* the part from nothing, emitting the prologue,
 * one `<row>` at a time and the epilogue, so the XML never exists as a whole.
 * Values are rendered by the SAME `generateCellXml` the patch writer uses, so
 * both write paths escape, space-preserve and encode dates/rich text/formulas
 * identically by construction.
 *
 * Schema order is respected (`sheetViews`, `cols`, `sheetData`), which is why
 * column widths and views are prologue properties: they must be known before
 * the first row is written (`PROLOGUE_LOCKED` afterwards). Merges are a
 * sibling of `sheetData`, so they can be added at any time and are emitted in
 * the epilogue.
 */
import type { CellValue } from '../model/value';
import type { SheetViewProps } from '../model/view';
import { SheetLiteError } from '../zip/errors';
import { assertMergeRange } from '../model/refs';
import { escapeXml, illegalXmlChar } from '../xml/escape';
import type { StreamingZipWriter } from '../zip/zip-stream';
import { generateCellXml, paneXmlFor } from './cell-xml';

/** Excel's hard limits: 1 048 576 rows x 16 384 columns (XFD). */
export const MAX_ROWS = 1_048_576;
export const MAX_COLUMNS = 16_384;

/** Width/visibility of one column, emitted in the prologue. */
export interface ColumnDef {
  width?: number;
  hidden?: boolean;
}

function assertWritableText(value: string, where: string): void {
  const illegal = illegalXmlChar(value);
  if (illegal) {
    throw new SheetLiteError(
      `Cannot write ${where}: the string contains ${illegal}, which is not valid in XML`,
      'BAD_TEXT',
    );
  }
}

/**
 * Reject values that would produce a file Excel refuses to open, instead of
 * writing `<v>NaN</v>` and letting the repair prompt happen on the user's side.
 * Called when a value is set, so the error points at the offending cell.
 */
export function assertWritableValue(value: CellValue, ref: string): void {
  if (typeof value === 'number') {
    if (!Number.isFinite(value)) {
      throw new SheetLiteError(
        `Cannot write ${ref}: ${String(value)} is not a valid cell number (NaN and Infinity are rejected)`,
        'BAD_NUMBER',
      );
    }
    return;
  }
  if (typeof value === 'string') {
    assertWritableText(value, ref);
    return;
  }
  if (value instanceof Date) {
    if (Number.isNaN(value.getTime())) {
      throw new SheetLiteError(`Cannot write ${ref}: the Date is invalid`, 'BAD_DATE');
    }
    return;
  }
  if (typeof value !== 'object' || value === null) return;
  if ('richText' in value) {
    value.richText.forEach((run, index) => assertWritableText(run.text ?? '', `${ref} run ${index}`));
    return;
  }
  if ('formula' in value) {
    assertWritableText(value.formula, `${ref} formula`);
    const result = value.result;
    if (typeof result === 'string') assertWritableText(result, `${ref} formula result`);
    else if (typeof result === 'number' && !Number.isFinite(result)) {
      throw new SheetLiteError(`Cannot write ${ref}: the formula result is not a finite number`, 'BAD_NUMBER');
    } else if (result instanceof Date && Number.isNaN(result.getTime())) {
      throw new SheetLiteError(`Cannot write ${ref}: the formula result is an invalid Date`, 'BAD_DATE');
    }
  }
}

/**
 * XML for one written cell. Empty string when the cell carries neither a value
 * nor a style (such cells are simply not emitted, like the patch writer does).
 * Values are validated when they are set (`assertWritableValue`).
 */
export function streamCellXml(
  ref: string,
  value: CellValue,
  styleIndex: number,
  date1904: boolean,
): string {
  if (value === null || value === undefined) {
    // A styled but empty cell still has to exist in the XML.
    return styleIndex ? `<c r="${ref}" s="${styleIndex}"/>` : '';
  }
  return generateCellXml(ref, value, styleIndex, date1904);
}

/** Streams one worksheet part into an open ZIP entry. */
export class SheetStreamWriter {
  private readonly encoder = new TextEncoder();
  private readonly cols = new Map<number, ColumnDef>();
  private readonly merges = new Set<string>();
  private views: SheetViewProps[] | null = null;
  private started = false;
  private ended = false;

  constructor(
    private readonly zip: StreamingZipWriter,
    /** Store key of the part, e.g. "xl/worksheets/sheet1.xml". */
    readonly partName: string,
    private readonly date1904: boolean,
  ) {}

  /** True once the first row (or `finish()`) has flushed the prologue. */
  get sealed(): boolean {
    return this.started;
  }

  /** True once the epilogue has been written and the ZIP entry is closed. */
  get done(): boolean {
    return this.ended;
  }

  /** Frozen panes. Must be set before the first row is written. */
  setViews(views: SheetViewProps[] | null): void {
    this.assertPrologue('views');
    this.views = views && views.length > 0 ? views : null;
  }

  /**
   * Set a column's width and/or visibility. Must be called before the first row
   * is written (the `<cols>` container precedes `<sheetData>` in the schema).
   *
   * @throws `BAD_COLUMN_WIDTH` outside Excel's 0-255 range, `PROLOGUE_LOCKED`
   *   once the part has started.
   */
  setColumn(col: number, def: ColumnDef): void {
    this.assertPrologue('column width');
    if (!Number.isFinite(col) || col < 1 || col > MAX_COLUMNS) {
      throw new SheetLiteError(`Invalid column ${col}: must be a number 1-${MAX_COLUMNS}`, 'BAD_COLUMN');
    }
    // Excel's own range is 0-255 (same rule and message as the eager writer's
    // `Worksheet.setColumnWidth`): anything else lands in `<col width="..">`
    // and makes Excel offer to repair the file.
    if (def.width !== undefined && !(def.width >= 0 && def.width <= 255)) {
      throw new SheetLiteError(
        `Invalid column width ${def.width}: must be a number 0-255`,
        'BAD_COLUMN_WIDTH',
      );
    }
    const current = this.cols.get(col);
    this.cols.set(col, { ...current, ...def });
  }

  /**
   * Merge a range (e.g. "A1:C1"); emitted after `</sheetData>`. The same range
   * twice costs one `<mergeCell>`.
   *
   * @throws `BAD_MERGE_REF` when `ref` is not a cell range.
   */
  mergeCells(ref: string): void {
    this.merges.add(assertMergeRange(ref));
  }

  /** Append one row; `cells` is the already-serialized `<c>` markup. */
  writeRow(number: number, height: number | undefined, cells: string): void {
    if (this.ended) {
      throw new SheetLiteError(
        `Cannot write row ${number} of ${this.partName}: the part has already been written`,
        'SHEET_FINALIZED',
      );
    }
    if (!Number.isFinite(number) || number < 1 || number > MAX_ROWS) {
      throw new SheetLiteError(`Invalid row ${number}: must be a number 1-${MAX_ROWS}`, 'BAD_ROW');
    }
    if (!this.started) this.writePrologue();
    const attrs = height === undefined ? '' : ` ht="${height}" customHeight="1"`;
    this.zip.write(
      this.encoder.encode(
        cells === '' ? `<row r="${number}"${attrs}/>` : `<row r="${number}"${attrs}>${cells}</row>`,
      ),
    );
  }

  /** Write the epilogue and close the ZIP entry. */
  finish(): void {
    if (this.ended) return;
    if (!this.started) this.writePrologue();
    this.zip.write(this.encoder.encode('</sheetData>'));
    if (this.merges.size > 0) {
      const refs = [...this.merges].map((ref) => `<mergeCell ref="${escapeXml(ref)}"/>`).join('');
      this.zip.write(
        this.encoder.encode(`<mergeCells count="${this.merges.size}">${refs}</mergeCells>`),
      );
    }
    this.zip.write(this.encoder.encode('</worksheet>'));
    this.zip.end();
    this.ended = true;
  }

  private assertPrologue(what: string): void {
    if (this.started) {
      throw new SheetLiteError(
        `Cannot set ${what} on ${this.partName} after its first row was written: the sheet prologue is already emitted`,
        'PROLOGUE_LOCKED',
      );
    }
  }

  private writePrologue(): void {
    this.started = true;
    this.zip.begin(this.partName);

    let views = '';
    const pane = paneXmlFor(this.views?.[0]);
    if (pane || this.views) {
      const workbookViewId = this.views?.[0]?.workbookViewId ?? 0;
      views = `<sheetViews><sheetView workbookViewId="${escapeXml(String(workbookViewId))}">${pane ?? ''}</sheetView></sheetViews>`;
    }

    let cols = '';
    if (this.cols.size > 0) {
      const entries = [...this.cols.keys()]
        .sort((a, b) => a - b)
        .map((col) => {
          const def = this.cols.get(col)!;
          const attrs = [`min="${col}"`, `max="${col}"`];
          if (def.width !== undefined) attrs.push(`width="${def.width}"`, 'customWidth="1"');
          if (def.hidden) attrs.push('hidden="1"');
          return `<col ${attrs.join(' ')}/>`;
        })
        .join('');
      cols = `<cols>${entries}</cols>`;
    }

    this.zip.write(
      this.encoder.encode(
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
          '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"' +
          ' xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">' +
          views +
          cols +
          '<sheetData>',
      ),
    );
  }
}