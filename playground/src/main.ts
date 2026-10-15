import './style.css';

import { Workbook, SheetLiteError } from '../../src/index';
import type {
  AlignmentProps,
  BorderProps,
  Cell,
  RichTextRun,
  SheetViewProps,
  Worksheet,
} from '../../src/index';
import { Grid } from './grid';
import { buildPatchReport, snapshotParts } from './patch';
import { createSampleWorkbook, createUnsafeBytes } from './sample';
import {
  columnLabel,
  displayDate,
  escapeHtml,
  fontStyle,
  formatBytes,
  formatDelta,
} from './format';

// --- DOM -------------------------------------------------------------------

function el<T extends HTMLElement>(id: string): T {
  const node = document.getElementById(id);
  if (!node) throw new Error(`Missing #${id} in index.html`);
  return node as T;
}

const fileChip = el('fileChip');
const dirtyChip = el('dirtyChip');
const btnOpen = el<HTMLButtonElement>('btnOpen');
const fileInput = el<HTMLInputElement>('fileInput');
const btnSample = el<HTMLButtonElement>('btnSample');
const btnUnsafe = el<HTMLButtonElement>('btnUnsafe');
const btnDownload = el<HTMLButtonElement>('btnDownload');
const errorBar = el('errorBar');
const errorName = el('errorName');
const errorCode = el('errorCode');
const errorMessage = el('errorMessage');
const errorClose = el<HTMLButtonElement>('errorClose');
const tabs = el('tabs');
const gridNotice = el('gridNotice');
const noticeText = el('noticeText');
const btnMoreRows = el<HTMLButtonElement>('btnMoreRows');
const gridScroll = el('gridScroll');
const statusLine = el('statusLine');
const inspRef = el('inspRef');
const inspValue = el('inspValue');
const inspType = el('inspType');
const inspNumFmt = el('inspNumFmt');
const inspFont = el('inspFont');
const inspFill = el('inspFill');
const inspBorder = el('inspBorder');
const inspAlignment = el('inspAlignment');
const actionSheet = el('actionSheet');
const btnBold = el<HTMLButtonElement>('btnBold');
const btnNumFmt = el<HTMLButtonElement>('btnNumFmt');
const btnAddRow = el<HTMLButtonElement>('btnAddRow');
const btnMerge = el<HTMLButtonElement>('btnMerge');
const btnUnmerge = el<HTMLButtonElement>('btnUnmerge');
const btnAlignLeft = el<HTMLButtonElement>('btnAlignLeft');
const btnAlignCenter = el<HTMLButtonElement>('btnAlignCenter');
const btnAlignRight = el<HTMLButtonElement>('btnAlignRight');
const btnWrap = el<HTMLButtonElement>('btnWrap');
const btnBoxBorder = el<HTMLButtonElement>('btnBoxBorder');
const btnEdgeBorder = el<HTMLButtonElement>('btnEdgeBorder');
const btnRichText = el<HTMLButtonElement>('btnRichText');
const structSheet = el('structSheet');
const newSheetName = el<HTMLInputElement>('newSheetName');
const btnAddSheet = el<HTMLButtonElement>('btnAddSheet');
const btnRemoveSheet = el<HTMLButtonElement>('btnRemoveSheet');
const colNumber = el<HTMLInputElement>('colNumber');
const colWidth = el<HTMLInputElement>('colWidth');
const colHidden = el<HTMLInputElement>('colHidden');
const btnColumn = el<HTMLButtonElement>('btnColumn');
const colState = el('colState');
const rowNumber = el<HTMLInputElement>('rowNumber');
const rowHeightInput = el<HTMLInputElement>('rowHeight');
const btnRowHeight = el<HTMLButtonElement>('btnRowHeight');
const rowState = el('rowState');
const freezeX = el<HTMLInputElement>('freezeX');
const freezeY = el<HTMLInputElement>('freezeY');
const btnFreeze = el<HTMLButtonElement>('btnFreeze');
const btnUnfreeze = el<HTMLButtonElement>('btnUnfreeze');
const viewsState = el('viewsState');
const sizeOriginal = el('sizeOriginal');
const sizeOutput = el('sizeOutput');
const sizeDelta = el('sizeDelta');
const partsSummary = el('partsSummary');
const partsList = el('partsList');
const emptyState = el('emptyState');
const btnPickEmpty = el<HTMLButtonElement>('btnPickEmpty');
const btnSampleEmpty = el<HTMLButtonElement>('btnSampleEmpty');
const btnUnsafeEmpty = el<HTMLButtonElement>('btnUnsafeEmpty');
const dropOverlay = el('dropOverlay');

// --- state -----------------------------------------------------------------

const state = {
  wb: null as Workbook | null,
  fileName: '',
  originalBytes: null as Uint8Array | null,
  originalParts: new Map<string, Uint8Array>(),
  sheetIndex: 0,
  edits: 0,
  /** `sheetId:col:row` the structure inputs were last filled for. */
  structTarget: '',
  /** `col:row` the grid last selected (the panel follows it until overridden). */
  structSelection: '1:1',
};

/** Merges created through the playground buttons in this session. */
const sessionMerges = new Set<string>();

/**
 * Structure edits applied in this session, per sheet id.
 *
 * Column widths, row heights and views are stored by the library as *pending*
 * patches: their getters only report what the file declared at load time, so
 * the panel mirrors what the playground just wrote (same idea as
 * `sessionMerges`). The grid preview and the inspector read from here.
 */
interface SheetStructure {
  colWidth: Map<number, number | undefined>;
  colHidden: Map<number, boolean>;
  rowHeight: Map<number, number | undefined>;
  views: SheetViewProps[] | null;
}

const sessionStructure = new Map<number, SheetStructure>();

function structureOf(sheet: Worksheet): SheetStructure {
  let entry = sessionStructure.get(sheet.id);
  if (!entry) {
    entry = {
      colWidth: new Map(),
      colHidden: new Map(),
      rowHeight: new Map(),
      views: null,
    };
    sessionStructure.set(sheet.id, entry);
  }
  return entry;
}

/** Effective column width: this session's edit, else what the file declares. */
function columnWidth(sheet: Worksheet, col: number): number | undefined {
  const pending = sessionStructure.get(sheet.id)?.colWidth;
  if (pending?.has(col)) return pending.get(col);
  return sheet.getColumn(col).width;
}

/** Effective hidden flag: this session's edit, else what the file declares. */
function columnHidden(sheet: Worksheet, col: number): boolean {
  const pending = sessionStructure.get(sheet.id)?.colHidden;
  if (pending?.has(col)) return pending.get(col) ?? false;
  return sheet.getColumn(col).hidden;
}

/** Effective row height: this session's edit, else what the file declares. */
function rowHeightOf(sheet: Worksheet, row: number): number | undefined {
  const pending = sessionStructure.get(sheet.id)?.rowHeight;
  if (pending?.has(row)) return pending.get(row);
  return sheet.getRow(row).height;
}

/** Effective views: this session's edit, else what the file declares. */
function viewsOf(sheet: Worksheet): SheetViewProps[] {
  const pending = sessionStructure.get(sheet.id)?.views;
  return pending ?? sheet.views;
}

function activeSheet(): Worksheet {
  const wb = state.wb;
  if (!wb || wb.worksheets.length === 0) {
    throw new SheetLiteError('No sheet available', 'EMPTY_WORKBOOK');
  }
  return wb.worksheets[state.sheetIndex] ?? wb.worksheets[0]!;
}

// --- grid ------------------------------------------------------------------

const grid = new Grid(gridScroll, {
  onSelected: () => {
    syncStructureToSelection();
    updateInspector();
    updateStatus();
    updateActionButtons();
  },
  onEdited: () => {
    state.edits++;
    scheduleMonitor();
    updateInspector();
    updateStatus();
    updateActionButtons();
    renderTabs();
    updateDirtyChip();
  },
  columnStyle: (col) => {
    const sheet = tryActiveSheet();
    return sheet
      ? { width: columnWidth(sheet, col), hidden: columnHidden(sheet, col) }
      : { width: undefined, hidden: false };
  },
  rowHeight: (row) => {
    const sheet = tryActiveSheet();
    return sheet ? rowHeightOf(sheet, row) : undefined;
  },
});

// --- errors ----------------------------------------------------------------

function showError(err: unknown): void {
  errorName.textContent = err instanceof Error ? err.name : 'Error';
  if (err instanceof SheetLiteError) {
    errorCode.hidden = false;
    errorCode.textContent = err.code;
  } else {
    errorCode.hidden = true;
  }
  errorMessage.textContent = err instanceof Error ? err.message : String(err);
  errorBar.hidden = false;
}

function hideError(): void {
  errorBar.hidden = true;
}

// --- loading ---------------------------------------------------------------

function loadFromBytes(bytes: Uint8Array, name: string): void {
  const wb = Workbook.read(bytes);
  if (wb.worksheets.length === 0) {
    throw new SheetLiteError('This workbook contains no sheets.', 'EMPTY_WORKBOOK');
  }

  state.wb = wb;
  state.originalBytes = bytes;
  state.originalParts = snapshotParts(wb);
  state.fileName = name;
  state.sheetIndex = 0;
  state.edits = 0;
  sessionMerges.clear();
  sessionStructure.clear();
  state.structTarget = '';
  state.structSelection = '1:1';

  hideError();
  emptyState.hidden = true;
  btnDownload.disabled = false;
  newSheetName.value = '';

  renderFileChip();
  renderTabs();
  grid.render(activeSheet());
  updateInspector();
  updateStatus();
  updateMonitor();
  updateDirtyChip();
  updateActionButtons();
  updateStructurePanel(true);
}

async function openFile(file: File): Promise<void> {
  try {
    const bytes = new Uint8Array(await file.arrayBuffer());
    loadFromBytes(bytes, file.name);
  } catch (err) {
    showError(err);
  }
}

function loadSample(): void {
  try {
    const wb = createSampleWorkbook();
    const bytes = wb.toBuffer();
    loadFromBytes(bytes, 'playground-sample.xlsx');
  } catch (err) {
    showError(err);
  }
}

function loadUnsafe(): void {
  try {
    loadFromBytes(createUnsafeBytes(), 'unsafe-package.xlsx');
  } catch (err) {
    showError(err);
  }
}

// --- chrome ----------------------------------------------------------------

function renderFileChip(): void {
  const wb = state.wb;
  if (!wb || state.originalBytes === null) return;
  fileChip.textContent = `${state.fileName} · ${formatBytes(state.originalBytes.byteLength)}`;
  fileChip.title = `${wb.sheetNames.length} sheet(s) · ${state.originalParts.size} OPC parts · loaded locally`;
  fileChip.hidden = false;
}

function renderTabs(): void {
  const wb = state.wb;
  if (!wb) return;
  tabs.replaceChildren(
    ...wb.sheetNames.map((name, i) => {
      const button = document.createElement('button');
      button.type = 'button';
      button.className = 'tab';
      button.setAttribute('role', 'tab');
      button.setAttribute('aria-selected', String(i === state.sheetIndex));
      if (wb.worksheets[i]?.isDirty()) button.classList.add('is-dirty');
      button.innerHTML = `${escapeHtml(name)}<span class="tab-dot" title="sheet has unsaved edits"></span>`;
      button.addEventListener('click', () => switchSheet(i));
      return button;
    }),
  );
}

function switchSheet(index: number): void {
  state.sheetIndex = index;
  state.structTarget = '';
  state.structSelection = '1:1';
  renderTabs();
  grid.render(activeSheet());
  updateInspector();
  updateStatus();
  updateActionButtons();
  updateStructurePanel(true);
}

function updateDirtyChip(): void {
  const dirty = state.wb?.worksheets.some((sheet) => sheet.isDirty()) ?? false;
  dirtyChip.hidden = !dirty;
}

function updateStatus(): void {
  const sel = grid.selected;
  const ref = sel ? `${columnLabel(sel.col)}${sel.row}` : '—';
  const truncated = grid.truncated ? ' · preview capped, see notice above' : '';
  statusLine.innerHTML =
    `<span class="ref">${escapeHtml(ref)}</span>` +
    `<span>${escapeHtml(grid.statusSummary())}${truncated}</span>` +
    `<span>double-click or type to edit · Enter commits · Esc cancels</span>`;

  if (!state.wb) {
    gridNotice.hidden = true;
    return;
  }
  if (grid.rowWindow < grid.totalRowCount || grid.colWindow < grid.totalColCount) {
    const parts: string[] = [];
    if (grid.rowWindow < grid.totalRowCount) {
      parts.push(
        `showing the first ${grid.rowWindow.toLocaleString()} of ` +
          `${grid.totalRowCount.toLocaleString()} rows`,
      );
    }
    if (grid.colWindow < grid.totalColCount) {
      parts.push(`${grid.colWindow} of ${grid.totalColCount} columns`);
    }
    noticeText.textContent = `Large sheet — ${parts.join(' and ')} to keep the playground snappy.`;
    btnMoreRows.hidden = grid.rowWindow >= grid.totalRowCount;
    gridNotice.hidden = false;
  } else {
    gridNotice.hidden = true;
  }
}

// --- inspector -------------------------------------------------------------

function updateInspector(): void {
  const sel = grid.selected;
  const sheet = state.wb ? tryActiveSheet() : null;
  if (!sel || !sheet) {
    inspRef.textContent = '—';
    inspValue.textContent = '—';
    inspType.textContent = '—';
    inspNumFmt.textContent = '—';
    inspFont.textContent = '—';
    inspFill.textContent = '—';
    inspBorder.textContent = '—';
    inspAlignment.textContent = '—';
    actionSheet.textContent = '—';
    return;
  }

  const cell = sheet.getCell(sel.row, sel.col);
  actionSheet.textContent = sheet.name;
  inspRef.textContent = cell.address;

  const value = cell.value;
  inspValue.title = '';
  if (value === null || value === undefined) {
    inspValue.textContent = '(empty)';
  } else if (typeof value === 'object' && 'richText' in value) {
    inspValue.replaceChildren(...richTextNodes(value.richText));
    inspValue.title = `${value.richText.length} run(s) — cell.value = { richText: [...] }`;
  } else if (typeof value === 'object' && 'formula' in value) {
    const result = value.result;
    const resultText =
      result === undefined || result === null
        ? 'not cached'
        : result instanceof Date
          ? displayDate(result)
          : String(result);
    inspValue.textContent = `=${value.formula}  →  ${resultText}`;
  } else if (typeof value === 'object' && 'error' in value) {
    inspValue.textContent = `#${value.error}`;
  } else {
    inspValue.textContent = value instanceof Date ? displayDate(value) : String(value);
  }

  inspType.textContent = '';
  inspType.append(typeChip(cell.type));
  inspNumFmt.textContent = cell.numFmt;

  const font = cell.font;
  const fontParts: string[] = [];
  if (font.bold) fontParts.push('bold');
  if (font.italic) fontParts.push('italic');
  if (font.underline) fontParts.push('underline');
  if (font.size !== undefined) fontParts.push(`${font.size}pt`);
  if (font.name) fontParts.push(font.name);
  if (font.color) fontParts.push(`#${font.color}`);
  inspFont.textContent = fontParts.length > 0 ? fontParts.join(' · ') : 'regular';

  const fill = cell.fill;
  if (fill?.pattern && fill.pattern !== 'none') {
    inspFill.textContent = '';
    if (fill.fgColor) {
      const swatch = document.createElement('span');
      swatch.className = 'fill-swatch';
      swatch.style.background = `#${fill.fgColor.slice(-6)}`;
      inspFill.append(swatch);
    }
    inspFill.append(
      document.createTextNode(`${fill.pattern}${fill.fgColor ? ` ${fill.fgColor}` : ''}`),
    );
  } else {
    inspFill.textContent = 'none';
  }

  inspBorder.textContent = describeBorder(cell.border);
  inspAlignment.textContent = describeAlignment(cell.alignment);
}

/** Rich-text runs as DOM nodes, each painted with its own run font. */
function richTextNodes(runs: RichTextRun[]): HTMLElement[] {
  return runs.map((run) => {
    const node = document.createElement('span');
    node.style.cssText = fontStyle(run.font);
    node.textContent = run.text;
    return node;
  });
}

function describeBorder(border: BorderProps): string {
  const sides: string[] = [];
  for (const side of ['top', 'right', 'bottom', 'left', 'diagonal'] as const) {
    const style = border[side]?.style;
    if (style) sides.push(`${side} ${style}`);
  }
  if (sides.length === 0) return 'none';
  if (border.diagonalUp || border.diagonalDown) {
    sides.push(`diagonal ${border.diagonalUp ? '▲' : ''}${border.diagonalDown ? '▼' : ''}`);
  }
  return sides.join(' · ');
}

function describeAlignment(alignment: AlignmentProps): string {
  const parts: string[] = [];
  if (alignment.horizontal) parts.push(`horizontal ${alignment.horizontal}`);
  if (alignment.vertical) parts.push(`vertical ${alignment.vertical}`);
  if (alignment.wrapText) parts.push('wrap');
  if (alignment.shrinkToFit) parts.push('shrink');
  if (alignment.indent !== undefined) parts.push(`indent ${alignment.indent}`);
  if (alignment.textRotation !== undefined) parts.push(`rotate ${alignment.textRotation}`);
  if (alignment.readingOrder !== undefined) parts.push(`reading order ${alignment.readingOrder}`);
  return parts.length > 0 ? parts.join(' · ') : 'default';
}

function tryActiveSheet(): Worksheet | null {
  try {
    return activeSheet();
  } catch {
    return null;
  }
}

function typeChip(type: string): HTMLElement {
  const chip = document.createElement('span');
  chip.className = 'type-chip';
  chip.textContent = type;
  return chip;
}

// --- actions ---------------------------------------------------------------

function updateActionButtons(): void {
  const wbLoaded = state.wb !== null;
  const sel = grid.selected;
  const cell = wbLoaded && sel ? tryActiveSheet()?.getCell(sel.row, sel.col) : undefined;
  const alignment = cell?.alignment;

  btnBold.disabled = !cell;
  btnNumFmt.disabled = !cell;
  btnAddRow.disabled = !wbLoaded;
  btnMerge.disabled = !wbLoaded;
  btnUnmerge.disabled = !wbLoaded || !sessionMerges.has('A1:B1');
  btnAlignLeft.disabled = !cell;
  btnAlignCenter.disabled = !cell;
  btnAlignRight.disabled = !cell;
  btnWrap.disabled = !cell;
  btnBoxBorder.disabled = !cell;
  btnEdgeBorder.disabled = !cell;
  btnRichText.disabled = !cell;
  btnBold.setAttribute('aria-pressed', String(cell?.font.bold === true));
  btnAlignLeft.setAttribute('aria-pressed', String(alignment?.horizontal === 'left'));
  btnAlignCenter.setAttribute('aria-pressed', String(alignment?.horizontal === 'center'));
  btnAlignRight.setAttribute('aria-pressed', String(alignment?.horizontal === 'right'));
  btnWrap.setAttribute('aria-pressed', String(alignment?.wrapText === true));

  updateStructurePanel();
}

function afterEdit(): void {
  state.edits++;
  scheduleMonitor();
  updateInspector();
  updateStatus();
  updateActionButtons();
  renderTabs();
  updateDirtyChip();
}

function afterStructuralChange(): void {
  updateMonitor();
  renderTabs();
  updateDirtyChip();
  updateStatus();
  updateInspector();
  updateActionButtons();
}

btnBold.addEventListener('click', () => {
  const sel = grid.selected;
  if (!sel) return;
  const cell = activeSheet().getCell(sel.row, sel.col);
  cell.font = { bold: !(cell.font.bold ?? false) };
  grid.refresh(sel.row, sel.col);
  afterEdit();
});

btnNumFmt.addEventListener('click', () => {
  const sel = grid.selected;
  if (!sel) return;
  const cell = activeSheet().getCell(sel.row, sel.col);
  cell.numFmt = '#,##0.00';
  grid.refresh(sel.row, sel.col);
  afterEdit();
});

btnAddRow.addEventListener('click', () => {
  const sheet = activeSheet();
  const r = sheet.rowCount + 1;
  sheet.addRow([
    'Added row',
    2,
    3.5,
    { formula: `B${r}*C${r}`, result: 7 },
    new Date(),
    true,
    'via sheet.addRow()',
  ]);
  grid.render(sheet);
  afterStructuralChange();
});

btnMerge.addEventListener('click', () => {
  activeSheet().mergeCells('A1:B1');
  sessionMerges.add('A1:B1');
  afterStructuralChange();
});

btnUnmerge.addEventListener('click', () => {
  activeSheet().unmergeCells('A1:B1');
  sessionMerges.delete('A1:B1');
  afterStructuralChange();
});

/** Apply an edit to the selected cell and repaint that cell only. */
function withSelectedCell(apply: (cell: Cell) => void): void {
  const sel = grid.selected;
  if (!sel) return;
  const sheet = tryActiveSheet();
  if (!sheet) return;
  apply(sheet.getCell(sel.row, sel.col));
  grid.refresh(sel.row, sel.col);
  afterEdit();
}

for (const [button, horizontal] of [
  [btnAlignLeft, 'left'],
  [btnAlignCenter, 'center'],
  [btnAlignRight, 'right'],
] as const) {
  button.addEventListener('click', () => {
    withSelectedCell((cell) => {
      cell.alignment = { horizontal };
    });
  });
}

btnWrap.addEventListener('click', () => {
  withSelectedCell((cell) => {
    cell.alignment = { wrapText: !cell.alignment.wrapText };
  });
});

btnBoxBorder.addEventListener('click', () => {
  withSelectedCell((cell) => {
    const thin = { style: 'thin' as const };
    cell.border = { top: thin, right: thin, bottom: thin, left: thin };
  });
});

btnEdgeBorder.addEventListener('click', () => {
  withSelectedCell((cell) => {
    cell.border = { bottom: { style: 'double' } };
  });
});

btnRichText.addEventListener('click', () => {
  withSelectedCell((cell) => {
    cell.value = {
      richText: [
        { text: 'Total: ' },
        { text: '1 234 €', font: { bold: true, color: 'FFFF0000' } },
        { text: ' · sheet-lite', font: { italic: true, size: 11 } },
      ],
    };
  });
});

btnMoreRows.addEventListener('click', () => {
  grid.showMoreRows();
  updateStatus();
  updateInspector();
});

// --- sheet structure -------------------------------------------------------

/** First free `SheetN` name, used when the name field is left empty. */
function suggestSheetName(wb: Workbook): string {
  const taken = new Set(wb.sheetNames);
  let n = 1;
  while (taken.has(`Sheet${n}`)) n++;
  return `Sheet${n}`;
}

/** Parse an optional numeric input: blank → undefined (clear the property). */
function optionalNumber(input: HTMLInputElement): number | undefined {
  const text = input.value.trim();
  if (text === '') return undefined;
  const value = Number(text);
  return Number.isFinite(value) ? value : undefined;
}

/** Clamp a 1-based column index to the range `getColumn` accepts (1-16384). */
function readColumn(): number {
  const n = Math.floor(Number(colNumber.value) || 1);
  return Math.min(Math.max(n, 1), 16384);
}

/** Clamp a 1-based row index to at least 1. */
function readRow(): number {
  return Math.max(Math.floor(Number(rowNumber.value) || 1), 1);
}

/**
 * Refresh the sheet-structure panel from the workbook. Inputs are only
 * overwritten when the target changed (selection / sheet switch / after an
 * apply), so typing in them is never interrupted.
 */
function updateStructurePanel(force = false): void {
  const sheet = tryActiveSheet();
  const sheetCount = state.wb?.worksheets.length ?? 0;
  const col = readColumn();
  const row = readRow();
  const height = sheet ? rowHeightOf(sheet, row) : undefined;
  const width = sheet ? columnWidth(sheet, col) : undefined;
  const hidden = sheet ? columnHidden(sheet, col) : false;
  const views = sheet ? viewsOf(sheet) : [];
  const pane = views[0];
  const frozen = !!pane && pane.state !== undefined && pane.state !== 'normal';

  structSheet.textContent = sheet?.name ?? '—';
  newSheetName.placeholder = state.wb ? suggestSheetName(state.wb) : 'Sheet1';
  btnAddSheet.disabled = !state.wb;
  btnRemoveSheet.disabled = !sheet || sheetCount < 2;
  btnRemoveSheet.title =
    sheetCount < 2
      ? 'A workbook must keep at least one sheet: removeWorksheet() throws LAST_SHEET here.'
      : `wb.removeWorksheet('${sheet?.name}') — drops the sheet part, its rel, the <sheet> entry and the content-type override`;
  btnColumn.disabled = !sheet;
  btnRowHeight.disabled = !sheet;
  btnFreeze.disabled = !sheet;
  btnUnfreeze.disabled = !sheet || !frozen;

  colState.textContent = sheet
    ? `Column ${columnLabel(col)} · width ${width === undefined ? 'default' : width}${hidden ? ' · hidden' : ''}`
    : '—';
  rowState.textContent = sheet
    ? `Row ${row} · height ${height === undefined ? 'default' : `${height}pt`}`
    : '—';
  viewsState.textContent = !sheet
    ? '—'
    : frozen
      ? `${pane.state} · ${pane.xSplit ?? 0} col(s) · ${pane.ySplit ?? 0} row(s)${pane.topLeftCell ? ` · ${pane.topLeftCell}` : ''}`
      : 'no frozen panes';

  // Never clobber an input the user is typing in: only refill when the panel
  // starts describing a different column/row (selection, sheet switch, apply).
  const target = `${sheet?.id ?? 0}:${col}:${row}`;
  if (force || target !== state.structTarget) {
    state.structTarget = target;
    colWidth.value = width === undefined ? '' : String(width);
    colHidden.checked = hidden;
    rowHeightInput.value = height === undefined ? '' : String(height);
    if (frozen) {
      freezeX.value = String(pane.xSplit ?? 0);
      freezeY.value = String(pane.ySplit ?? 0);
    }
  }
}

/** Point the structure panel at the cell the user selected in the grid. */
function syncStructureToSelection(): void {
  const sel = grid.selected;
  if (!sel) return;
  const next = `${sel.col}:${sel.row}`;
  // Moving the selection always retargets the panel; re-rendering the grid
  // (which re-selects the same cell) only does so while the panel still
  // points at it, so a column/row the user is editing is never overwritten.
  if (next !== state.structSelection || `${readColumn()}:${readRow()}` === next) {
    colNumber.value = String(sel.col);
    rowNumber.value = String(sel.row);
  }
  state.structSelection = next;
}

newSheetName.addEventListener('keydown', (e) => {
  if (e.key === 'Enter') {
    e.preventDefault();
    btnAddSheet.click();
  }
});

btnAddSheet.addEventListener('click', () => {
  const wb = state.wb;
  if (!wb) return;
  const name = newSheetName.value.trim() || suggestSheetName(wb);
  try {
    const sheet = wb.addWorksheet(name);
    state.sheetIndex = wb.worksheets.indexOf(sheet);
    newSheetName.value = '';
    hideError();
    renderFileChip();
    grid.render(sheet);
    updateStructurePanel(true);
    afterStructuralChange();
  } catch (err) {
    showError(err); // SHEET_EXISTS / BAD_SHEET_NAME
  }
});

btnRemoveSheet.addEventListener('click', () => {
  const wb = state.wb;
  const sheet = tryActiveSheet();
  if (!wb || !sheet) return;
  try {
    wb.removeWorksheet(sheet.name);
    sessionStructure.delete(sheet.id);
    state.sheetIndex = Math.min(state.sheetIndex, wb.worksheets.length - 1);
    hideError();
    renderFileChip();
    grid.render(activeSheet());
    updateStructurePanel(true);
    afterStructuralChange();
  } catch (err) {
    showError(err); // LAST_SHEET / SHEET_NOT_FOUND
  }
});

btnColumn.addEventListener('click', () => {
  const sheet = tryActiveSheet();
  if (!sheet) return;
  const col = readColumn();
  const width = optionalNumber(colWidth);
  try {
    sheet.getColumn(col).width = width;
    sheet.getColumn(col).hidden = colHidden.checked;
    const pending = structureOf(sheet);
    pending.colWidth.set(col, width);
    pending.colHidden.set(col, colHidden.checked);
    hideError();
    grid.refreshStructure();
    updateStructurePanel(true);
    afterStructuralChange();
  } catch (err) {
    showError(err); // BAD_COLUMN
  }
});

btnRowHeight.addEventListener('click', () => {
  const sheet = tryActiveSheet();
  if (!sheet) return;
  const row = readRow();
  const height = optionalNumber(rowHeightInput);
  try {
    sheet.getRow(row).height = height;
    structureOf(sheet).rowHeight.set(row, height);
    hideError();
    grid.refreshStructure();
    updateStructurePanel(true);
    afterStructuralChange();
  } catch (err) {
    showError(err); // BAD_ROW_HEIGHT
  }
});

btnFreeze.addEventListener('click', () => {
  const sheet = tryActiveSheet();
  if (!sheet) return;
  const xSplit = Math.max(Math.floor(Number(freezeX.value) || 0), 0);
  const ySplit = Math.max(Math.floor(Number(freezeY.value) || 0), 0);
  try {
    const views: SheetViewProps[] = [{ state: 'frozen', xSplit, ySplit }];
    sheet.views = views;
    structureOf(sheet).views = views;
    hideError();
    updateStructurePanel(true);
    afterStructuralChange();
  } catch (err) {
    showError(err);
  }
});

btnUnfreeze.addEventListener('click', () => {
  const sheet = tryActiveSheet();
  if (!sheet) return;
  try {
    const views: SheetViewProps[] = [{ state: 'normal' }];
    sheet.views = views;
    structureOf(sheet).views = views;
    hideError();
    updateStructurePanel(true);
    afterStructuralChange();
  } catch (err) {
    showError(err);
  }
});

colNumber.addEventListener('input', () => updateStructurePanel());
rowNumber.addEventListener('input', () => updateStructurePanel());

btnDownload.addEventListener('click', () => {
  const wb = state.wb;
  if (!wb) return;
  // `toBuffer()` returns the (patched) package bytes; copy into a fresh
  // ArrayBuffer so the Blob owns its data.
  const bytes = new Uint8Array(wb.toBuffer());
  const blob = new Blob([bytes], {
    type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
  });
  const url = URL.createObjectURL(blob);
  const link = document.createElement('a');
  const base = state.fileName.replace(/\.xlsx$/i, '') || 'workbook';
  link.href = url;
  link.download = `${base}-edited.xlsx`;
  document.body.append(link);
  link.click();
  link.remove();
  setTimeout(() => URL.revokeObjectURL(url), 1_000);
});

// --- patch monitor ---------------------------------------------------------

let monitorTimer: number | undefined;

function scheduleMonitor(): void {
  window.clearTimeout(monitorTimer);
  monitorTimer = window.setTimeout(updateMonitor, 120);
}

function updateMonitor(): void {
  if (!state.wb || state.originalBytes === null) return;
  const report = buildPatchReport(state.originalBytes, state.originalParts, state.wb);

  sizeOriginal.textContent = formatBytes(report.originalSize);
  sizeOutput.textContent = formatBytes(report.outputSize);

  const delta = report.outputSize - report.originalSize;
  sizeDelta.hidden = false;
  if (report.untouched) {
    sizeDelta.textContent = 'byte-identical';
    sizeDelta.className = 'chip chip-delta zero';
  } else if (delta === 0) {
    sizeDelta.textContent = 'same size';
    sizeDelta.className = 'chip chip-delta zero';
  } else {
    sizeDelta.textContent = formatDelta(delta);
    sizeDelta.className = `chip chip-delta ${delta > 0 ? 'up' : 'down'}`;
  }

  if (report.untouched) {
    partsSummary.textContent =
      'No edits yet — toBuffer() returns the original bytes verbatim.';
  } else {
    const summary = document.createElement('span');
    const strong = (text: string): HTMLElement => {
      const node = document.createElement('strong');
      node.textContent = text;
      return node;
    };
    summary.append(
      strong(`${report.identicalCount} of ${report.totalParts}`),
      document.createTextNode(' parts byte-identical · '),
      strong(String(report.changedCount)),
      document.createTextNode(' re-serialized'),
    );
    if (report.addedCount > 0) summary.append(document.createTextNode(` · ${report.addedCount} added`));
    if (report.removedCount > 0) summary.append(document.createTextNode(` · ${report.removedCount} removed`));
    partsSummary.replaceChildren(summary);
  }

  partsList.replaceChildren(
    ...report.parts.map((row) => {
      const li = document.createElement('li');
      li.title = row.status;
      const dot = document.createElement('span');
      dot.className = `dot ${row.status}`;
      const name = document.createElement('span');
      name.className = 'part-name';
      name.textContent = row.name;
      const size = document.createElement('span');
      size.className = 'part-size';
      size.textContent = formatBytes(row.size);
      li.append(dot, name, size);
      return li;
    }),
  );

  if (report.parseError) {
    const note = document.createElement('li');
    note.textContent = `Could not re-read the output: ${report.parseError}`;
    partsList.append(note);
  }
}

// --- wiring ----------------------------------------------------------------

btnOpen.addEventListener('click', () => fileInput.click());
btnPickEmpty.addEventListener('click', () => fileInput.click());
btnSample.addEventListener('click', loadSample);
btnSampleEmpty.addEventListener('click', loadSample);
btnUnsafe.addEventListener('click', loadUnsafe);
btnUnsafeEmpty.addEventListener('click', loadUnsafe);
errorClose.addEventListener('click', hideError);

fileInput.addEventListener('change', () => {
  const file = fileInput.files?.[0];
  if (file) void openFile(file);
  fileInput.value = '';
});

let dragDepth = 0;
window.addEventListener('dragenter', (e) => {
  e.preventDefault();
  dragDepth++;
  dropOverlay.hidden = false;
});
window.addEventListener('dragover', (e) => e.preventDefault());
window.addEventListener('dragleave', () => {
  dragDepth = Math.max(0, dragDepth - 1);
  if (dragDepth === 0) dropOverlay.hidden = true;
});
window.addEventListener('drop', (e) => {
  e.preventDefault();
  dragDepth = 0;
  dropOverlay.hidden = true;
  const file = e.dataTransfer?.files?.[0];
  if (file) void openFile(file);
});

window.addEventListener('beforeunload', (e) => {
  if (state.wb?.worksheets.some((sheet) => sheet.isDirty())) {
    e.preventDefault();
    e.returnValue = '';
  }
});

updateInspector();
updateActionButtons();
updateStructurePanel(true);
