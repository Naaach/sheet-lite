import type { Cell, Worksheet } from '../../src/index';
import {
  columnLabel,
  displayValue,
  editSeed,
  escapeHtml,
  excelWidthPx,
  fontStyle,
  parseCellInput,
  rowHeightPx,
} from './format';

/** Hard caps so a huge workbook cannot take down the tab (see README). */
const ROW_STEP = 500;
const MIN_ROWS = 24;
const MAX_COLS = 40;
const MIN_COLS = 10;

export interface Selection {
  row: number;
  col: number;
}

/** Presentation hints for one column, derived from `ws.getColumn(n)`. */
export interface ColumnMeta {
  /** `Column.width` in Excel character units. */
  width: number | undefined;
  /** `Column.hidden`. */
  hidden: boolean;
}

export interface GridCallbacks {
  /** The selected cell changed (Cell is a fresh view object for that ref). */
  onSelected(cell: Cell): void;
  /** A value was committed into a cell. */
  onEdited(cell: Cell): void;
  /** Width / hidden of a column (optional: the grid just uses defaults). */
  columnStyle?(col: number): ColumnMeta;
  /** Explicit row height in points (`Row.height`), or undefined. */
  rowHeight?(row: number): number | undefined;
}

const TYPE_CLASS: Record<string, string> = {
  Number: 'num',
  Date: 'date',
  Boolean: 'bool',
  Formula: 'formula',
  Error: 'err',
};

/**
 * Editable spreadsheet grid.
 * Click selects, double-click / Enter / F2 / typing edits, arrows navigate.
 * Rendering is capped (see constants); "show more" grows it in steps.
 */
export class Grid {
  private sheet: Worksheet | null = null;
  private table: HTMLTableElement | null = null;
  private selection: Selection | null = null;
  private editing: { row: number; col: number; input: HTMLInputElement } | null = null;
  private renderRows = 0;
  private renderCols = 0;
  private totalRows = 0;
  private totalCols = 0;
  /** Per-column presentation resolved once per rebuild (see `columnStyle`). */
  private colMeta: ColumnMeta[] = [];

  constructor(
    private readonly scroll: HTMLElement,
    private readonly cb: GridCallbacks,
  ) {
    scroll.addEventListener('keydown', (e) => this.onKeyDown(e));
    scroll.addEventListener('click', (e) => {
      const td = (e.target as HTMLElement).closest<HTMLTableCellElement>('td[data-r]');
      if (td) this.select(Number(td.dataset.r), Number(td.dataset.c));
    });
    scroll.addEventListener('dblclick', (e) => {
      const td = (e.target as HTMLElement).closest<HTMLTableCellElement>('td[data-r]');
      if (td) {
        const row = Number(td.dataset.r);
        const col = Number(td.dataset.c);
        this.select(row, col);
        this.startEdit('');
      }
    });
  }

  get selected(): Selection | null {
    return this.selection;
  }

  /** Rows currently rendered. */
  get rowWindow(): number {
    return this.renderRows;
  }

  /** Total rows with content in the sheet. */
  get totalRowCount(): number {
    return this.totalRows;
  }

  /** Columns currently rendered. */
  get colWindow(): number {
    return this.renderCols;
  }

  /** Total columns with content in the sheet. */
  get totalColCount(): number {
    return this.totalCols;
  }

  get truncated(): boolean {
    return this.renderRows < this.totalRows || this.renderCols < this.totalCols;
  }

  statusSummary(): string {
    const rowText =
      this.renderRows < this.totalRows
        ? `rows 1–${this.renderRows.toLocaleString()} of ${this.totalRows.toLocaleString()}`
        : `rows 1–${this.renderRows.toLocaleString()}`;
    const colText =
      this.renderCols < this.totalCols
        ? `${this.renderCols} of ${this.totalCols} columns`
        : `${this.renderCols} column${this.renderCols === 1 ? '' : 's'}`;
    return `${rowText} · ${colText}`;
  }

  /** Full rebuild for a sheet. Resets the row cap and selects A1. */
  render(sheet: Worksheet): void {
    this.cancelEdit();
    this.sheet = sheet;
    this.totalRows = Math.max(sheet.rowCount, 1);
    this.totalCols = Math.max(sheet.columnCount, 1);
    this.renderRows = Math.min(Math.max(this.totalRows, MIN_ROWS), ROW_STEP);
    this.renderCols = Math.min(Math.max(this.totalCols, MIN_COLS), MAX_COLS);
    this.rebuild();
    this.select(1, 1);
  }

  /** Grow the rendered row window (keeps the current selection). */
  showMoreRows(): void {
    if (!this.sheet) return;
    this.renderRows = Math.min(this.renderRows + ROW_STEP, this.totalRows);
    this.rebuildKeepSelection();
  }

  /** Re-render after a column width / hidden column / row height change. */
  refreshStructure(): void {
    if (!this.sheet) return;
    this.rebuildKeepSelection();
  }

  /** Re-render one cell in place after an out-of-band edit (styles, formats). */
  refresh(row: number, col: number): void {
    this.refreshCell(row, col);
  }

  // --- internals -----------------------------------------------------------

  private rebuildKeepSelection(): void {
    const keep = this.selection ?? { row: 1, col: 1 };
    this.rebuild();
    this.select(keep.row, keep.col);
  }

  private rebuild(): void {
    const sheet = this.sheet;
    if (!sheet) return;

    const cols = this.renderCols;
    this.colMeta = Array.from({ length: cols }, (_, i) => this.cb.columnStyle?.(i + 1) ?? {
      width: undefined,
      hidden: false,
    });

    const head = [`<th class="corner" aria-hidden="true"></th>`];
    const colgroup = [`<col style="width:44px">`];
    for (let c = 1; c <= cols; c++) {
      const meta = this.colMeta[c - 1]!;
      head.push(
        `<th scope="col"${meta.hidden ? ' class="col-hidden"' : ''}>${columnLabel(c)}</th>`,
      );
      colgroup.push(
        meta.width === undefined ? '<col>' : `<col style="width:${excelWidthPx(meta.width)}px">`,
      );
    }

    const body: string[] = [];
    for (let r = 1; r <= this.renderRows; r++) {
      const height = this.cb.rowHeight?.(r);
      const style = height === undefined ? '' : ` style="height:${rowHeightPx(height)}px"`;
      body.push(
        `<tr${style}><th scope="row">${r}</th>${this.rowCellsHtml(sheet, r, cols)}</tr>`,
      );
    }

    this.table = document.createElement('table');
    this.table.className = 'sheet';
    this.table.innerHTML =
      `<colgroup>${colgroup.join('')}</colgroup>` +
      `<thead><tr>${head.join('')}</tr></thead><tbody>${body.join('')}</tbody>`;
    this.scroll.replaceChildren(this.table);
  }

  private rowCellsHtml(sheet: Worksheet, row: number, cols: number): string {
    let html = '';
    for (let c = 1; c <= cols; c++) {
      const cell = sheet.getCell(row, c);
      html += this.cellHtml(cell, row, c);
    }
    return html;
  }

  private cellHtml(cell: Cell, row: number, col: number): string {
    const value = cell.value;
    const selected = this.selection?.row === row && this.selection?.col === col ? ' selected' : '';
    let classes = `${TYPE_CLASS[cell.type] ?? ''}${selected}${this.cellExtraClass(cell)}`;
    let title = '';

    if (cell.type === 'Formula') {
      const formula = value && typeof value === 'object' && 'formula' in value ? value.formula : '';
      title = ` title="=${escapeHtml(formula)}"`;
    }

    return `<td data-r="${row}" data-c="${col}" class="${classes}"${title}>${cellContentHtml(cell)}</td>`;
  }

  /** Cell-level classes that are not part of the value type (bold, hidden column). */
  private cellExtraClass(cell: Cell): string {
    const value = cell.value;
    let classes = '';
    if (value !== null && typeof value === 'object' && 'richText' in value) {
      classes += ' rich';
    } else if (value !== null) {
      const font = cell.font;
      if (font.bold) classes += ' bold';
      if (font.italic) classes += ' italic';
    }
    if (this.colMeta[cell.col - 1]?.hidden) classes += ' col-hidden';
    return classes;
  }

  private tdAt(row: number, col: number): HTMLTableCellElement | null {
    return this.table?.querySelector<HTMLTableCellElement>(`td[data-r="${row}"][data-c="${col}"]`) ?? null;
  }

  select(row: number, col: number): void {
    if (!this.sheet) return;
    row = Math.min(Math.max(row, 1), this.renderRows);
    col = Math.min(Math.max(col, 1), this.renderCols);
    if (this.selection) {
      this.tdAt(this.selection.row, this.selection.col)?.classList.remove('selected');
    }
    this.selection = { row, col };
    const td = this.tdAt(row, col);
    td?.classList.add('selected');
    this.cb.onSelected(this.sheet.getCell(row, col));
  }

  private moveSelection(dr: number, dc: number): void {
    if (!this.selection) return;
    this.select(this.selection.row + dr, this.selection.col + dc);
    this.tdAt(this.selection.row, this.selection.col)?.scrollIntoView({
      block: 'nearest',
      inline: 'nearest',
    });
  }

  /** Re-render a single cell after an edit (no full rebuild). */
  private refreshCell(row: number, col: number): void {
    const td = this.tdAt(row, col);
    if (!td || !this.sheet) return;
    const cell = this.sheet.getCell(row, col);
    td.className = `${TYPE_CLASS[cell.type] ?? ''}${this.selection?.row === row && this.selection?.col === col ? ' selected' : ''}${this.cellExtraClass(cell)}`;
    td.removeAttribute('title');
    if (cell.type === 'Formula') {
      const value = cell.value;
      const formula = value && typeof value === 'object' && 'formula' in value ? value.formula : '';
      td.setAttribute('title', `=${formula}`);
    }
    td.innerHTML = cellContentHtml(cell);
  }

  private onKeyDown(e: KeyboardEvent): void {
    if (!this.sheet || !this.selection) return;

    if (this.editing) {
      if (e.key === 'Enter') {
        e.preventDefault();
        this.commitEdit();
        this.moveSelection(1, 0);
      } else if (e.key === 'Tab') {
        e.preventDefault();
        this.commitEdit();
        this.moveSelection(0, e.shiftKey ? -1 : 1);
      } else if (e.key === 'Escape') {
        e.preventDefault();
        this.cancelEdit();
        this.scroll.focus();
      }
      return;
    }

    switch (e.key) {
      case 'ArrowUp':
        e.preventDefault();
        this.moveSelection(-1, 0);
        break;
      case 'ArrowDown':
        e.preventDefault();
        this.moveSelection(1, 0);
        break;
      case 'ArrowLeft':
        e.preventDefault();
        this.moveSelection(0, -1);
        break;
      case 'ArrowRight':
        e.preventDefault();
        this.moveSelection(0, 1);
        break;
      case 'Tab':
        e.preventDefault();
        this.moveSelection(0, e.shiftKey ? -1 : 1);
        break;
      case 'Enter':
      case 'F2':
        e.preventDefault();
        this.startEdit('');
        break;
      case 'Delete':
      case 'Backspace':
        e.preventDefault();
        this.clearSelected();
        break;
      default:
        if (e.key.length === 1 && !e.ctrlKey && !e.metaKey && !e.altKey) {
          e.preventDefault();
          this.startEdit(e.key);
        }
    }
  }

  private startEdit(initial: string): void {
    const sel = this.selection;
    if (!this.sheet || !sel || this.editing) return;
    const td = this.tdAt(sel.row, sel.col);
    if (!td) return;

    const cell = this.sheet.getCell(sel.row, sel.col);
    const input = document.createElement('input');
    input.className = 'cell-editor';
    input.value = initial || editSeed(cell.value);
    input.setAttribute('aria-label', `Edit cell ${cell.address}`);
    td.classList.add('editing');
    td.replaceChildren(input);
    input.focus();
    if (initial) input.setSelectionRange(initial.length, initial.length);
    else input.select();

    const entry = { row: sel.row, col: sel.col, input };
    this.editing = entry;
    input.addEventListener('blur', () => {
      if (this.editing === entry) this.commitEdit(false);
    });
  }

  private commitEdit(refocus = true): void {
    const entry = this.editing;
    if (!entry || !this.sheet) return;
    this.editing = null;

    const text = entry.input.value;
    const td = this.tdAt(entry.row, entry.col);
    td?.classList.remove('editing');
    entry.input.remove();

    const cell = this.sheet.getCell(entry.row, entry.col);
    cell.value = parseCellInput(text);
    this.refreshCell(entry.row, entry.col);
    this.cb.onEdited(cell);
    if (refocus) this.scroll.focus();
  }

  private cancelEdit(): void {
    const entry = this.editing;
    if (!entry) return;
    this.editing = null;
    const td = this.tdAt(entry.row, entry.col);
    td?.classList.remove('editing');
    entry.input.remove();
    this.refreshCell(entry.row, entry.col);
  }

  private clearSelected(): void {
    const sel = this.selection;
    if (!this.sheet || !sel) return;
    const cell = this.sheet.getCell(sel.row, sel.col);
    cell.value = null;
    this.refreshCell(sel.row, sel.col);
    this.cb.onEdited(cell);
  }
}

/**
 * Inner HTML of a cell: rich-text runs keep their own font styling, formulas
 * get the `f` badge, everything else is escaped plain text.
 */
function cellContentHtml(cell: Cell): string {
  const value = cell.value;

  if (value !== null && typeof value === 'object' && 'richText' in value) {
    return value.richText
      .map((run) => {
        const style = fontStyle(run.font);
        return `<span${style ? ` style="${style}"` : ''}>${escapeHtml(run.text)}</span>`;
      })
      .join('');
  }

  const text = escapeHtml(displayValue(value, cell.numFmt));
  if (cell.type === 'Formula') {
    const formula = value && typeof value === 'object' && 'formula' in value ? value.formula : '';
    return `<span class="fx-badge" title="=${escapeHtml(formula)}">f</span>${text}`;
  }
  return text;
}
