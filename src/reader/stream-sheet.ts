/**
 * Streaming worksheet reader.
 *
 * A push-based counterpart of `reader/sheet-parser.ts`: instead of parsing a
 * whole sheet into a DOM, an incremental pull-parser consumes UTF-8 text
 * chunks and a small state machine decodes one `<row>` at a time. Only the
 * current row's cells are ever materialized, so a sheet of any size can be
 * walked with (almost) constant memory.
 *
 * Value rules are intentionally identical to `computeValue` in
 * `sheet-parser.ts` — shared strings, inline strings, formula strings,
 * booleans, errors, numbers and date-formatted numbers (1900/1904), plus
 * formulas with their cached result — so eager and streaming reads always
 * agree on values and types.
 */
import { decodeEntities } from '../xml/xml';
import { XmlPullParser, type XmlEvent } from '../xml/pull-parser';
import { serialToDate } from '../model/dates';
import { typeOf, type CellValue, type ValueTypeName } from '../model/value';
import { coordsToRef, refToCoords } from './sheet-parser';
import type { PackageMeta } from './package-reader';

export interface StreamCell {
  ref: string;
  row: number;
  col: number;
  value: CellValue;
  type: ValueTypeName;
}

export interface StreamRow {
  /** Row number (`r` attribute, or the implicit 1-based position). */
  number: number;
  cells: StreamCell[];
}

// ---------------------------------------------------------------------------
// Row/cell state machine
// ---------------------------------------------------------------------------

interface CellBuild {
  attrs: Record<string, string>;
  /** Raw (still entity-encoded) text of the first `<v>`, `<f>`, `<is>`. */
  v: string | undefined;
  f: string | undefined;
  is: string | undefined;
}

interface RowBuild {
  r: number;
  prevCol: number;
  cells: StreamCell[];
}

export class SheetRowReader {
  private parser = new XmlPullParser();
  private decoder = new TextDecoder();
  private inSheetData = false;
  private implicitRow = 0;
  private row: RowBuild | null = null;
  private cell: CellBuild | null = null;
  /** Element currently collecting text (`v` / `f` / `is`). */
  private target: 'v' | 'f' | 'is' | null = null;
  private targetText = '';

  constructor(private readonly meta: PackageMeta) {}

  /** Unconsumed XML bytes currently buffered (diagnostics). */
  get bufferedLength(): number {
    return this.parser.buffered;
  }

  /** Feed one decompressed chunk (UTF-8 bytes). */
  write(chunk: Uint8Array): void {
    this.parser.push(this.decoder.decode(chunk, { stream: true }));
  }

  /** Signal end-of-stream; trailing malformed text is surfaced on `next()`. */
  end(): void {
    this.parser.push(this.decoder.decode());
    this.parser.end();
  }

  /** Next completed row, or `null` when more input is needed / the sheet ended. */
  next(): StreamRow | null {
    for (;;) {
      const event = this.parser.next();
      if (event === null) {
        // Malformed-but-recoverable input: flush a still-open row at EOF.
        if (this.parser.ended && this.row) {
          if (this.cell) this.finishCell();
          return this.finishRow();
        }
        return null;
      }
      const row = this.handle(event);
      if (row) return row;
    }
  }

  private handle(event: XmlEvent): StreamRow | null {
    if (event.kind === 'text') {
      if (this.target) this.targetText += event.text;
      return null;
    }

    if (event.kind === 'open') {
      switch (event.name) {
        case 'sheetData':
          this.inSheetData = true;
          return null;
        case 'row': {
          if (!this.inSheetData) return null;
          this.implicitRow += 1;
          const r = event.attrs.r !== undefined ? Number(event.attrs.r) : this.implicitRow;
          this.row = { r, prevCol: 0, cells: [] };
          if (event.selfClosing) return this.finishRow();
          return null;
        }
        case 'c': {
          if (!this.row) return null;
          this.cell = { attrs: event.attrs, v: undefined, f: undefined, is: undefined };
          if (event.selfClosing) this.finishCell();
          return null;
        }
        case 'v':
        case 'f':
        case 'is': {
          if (!this.cell) return null;
          this.target = event.name;
          this.targetText = '';
          if (event.selfClosing) this.closeTarget();
          return null;
        }
        default:
          // `<t>` inside `<is>`, rich-text `<r>` runs, … keep collecting text.
          return null;
      }
    }

    switch (event.name) {
      case 'sheetData':
        this.inSheetData = false;
        return null;
      case 'v':
      case 'f':
      case 'is':
        if (this.target === event.name) this.closeTarget();
        return null;
      case 'c':
        if (this.cell) this.finishCell();
        return null;
      case 'row':
        return this.row ? this.finishRow() : null;
      default:
        return null;
    }
  }

  private closeTarget(): void {
    const cell = this.cell;
    if (!cell) return;
    if (this.target === 'v' && cell.v === undefined) cell.v = this.targetText;
    else if (this.target === 'f' && cell.f === undefined) cell.f = this.targetText;
    else if (this.target === 'is' && cell.is === undefined) cell.is = this.targetText;
    this.target = null;
    this.targetText = '';
  }

  /** Decode one cell with the SAME rules as `computeValue` in sheet-parser.ts. */
  private finishCell(): void {
    const cell = this.cell!;
    const row = this.row!;
    this.cell = null;

    const attrs = cell.attrs;
    let ref = attrs.r;
    let col: number;
    if (ref) {
      col = refToCoords(ref).col;
    } else {
      col = row.prevCol + 1;
      ref = coordsToRef(row.r, col);
    }
    row.prevCol = col;
    const { row: refRow } = refToCoords(ref);
    const styleIndex = attrs.s ? Number(attrs.s) : 0;

    const rawText = cell.v;
    const hasText = rawText !== undefined && rawText !== '';

    let value: CellValue = null;
    if (attrs.t === 's') {
      value = hasText ? this.meta.sharedStrings[Number(rawText)] ?? '' : '';
    } else if (attrs.t === 'inlineStr') {
      value = cell.is !== undefined ? decodeEntities(cell.is) : '';
    } else if (attrs.t === 'str') {
      value = rawText ?? '';
    } else if (attrs.t === 'b') {
      value = hasText ? rawText === '1' : null;
    } else if (attrs.t === 'e') {
      value = hasText ? { error: rawText! } : null;
    } else if (hasText) {
      const numeric = Number(rawText);
      value = this.meta.dateStyleIndexes.has(styleIndex)
        ? serialToDate(numeric, this.meta.date1904)
        : numeric;
    }

    if (cell.f !== undefined) {
      value = { formula: decodeEntities(cell.f), result: value ?? undefined };
    }

    row.cells.push({ ref, row: refRow, col, value, type: typeOf(value) });
  }

  private finishRow(): StreamRow {
    const row = this.row!;
    this.row = null;
    return { number: row.r, cells: row.cells };
  }
}