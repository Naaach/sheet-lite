import './style.css';

import { Workbook, SheetLiteError } from '../../src/index';
import type { Worksheet } from '../../src/index';
import { Grid } from './grid';
import { buildPatchReport, snapshotParts } from './patch';
import { createSampleWorkbook, createUnsafeBytes } from './sample';
import { columnLabel, displayDate, escapeHtml, formatBytes, formatDelta } from './format';

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
const actionSheet = el('actionSheet');
const btnBold = el<HTMLButtonElement>('btnBold');
const btnNumFmt = el<HTMLButtonElement>('btnNumFmt');
const btnAddRow = el<HTMLButtonElement>('btnAddRow');
const btnMerge = el<HTMLButtonElement>('btnMerge');
const btnUnmerge = el<HTMLButtonElement>('btnUnmerge');
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
};

/** Merges created through the playground buttons in this session. */
const sessionMerges = new Set<string>();

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

  hideError();
  emptyState.hidden = true;
  btnDownload.disabled = false;

  renderFileChip();
  renderTabs();
  grid.render(activeSheet());
  updateInspector();
  updateStatus();
  updateMonitor();
  updateDirtyChip();
  updateActionButtons();
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
  renderTabs();
  grid.render(activeSheet());
  updateInspector();
  updateStatus();
  updateActionButtons();
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
    actionSheet.textContent = '—';
    return;
  }

  const cell = sheet.getCell(sel.row, sel.col);
  actionSheet.textContent = sheet.name;
  inspRef.textContent = cell.address;

  const value = cell.value;
  if (value === null || value === undefined) {
    inspValue.textContent = '(empty)';
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

  btnBold.disabled = !cell;
  btnNumFmt.disabled = !cell;
  btnAddRow.disabled = !wbLoaded;
  btnMerge.disabled = !wbLoaded;
  btnUnmerge.disabled = !wbLoaded || !sessionMerges.has('A1:B1');
  btnBold.setAttribute('aria-pressed', String(cell?.font.bold === true));
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

btnMoreRows.addEventListener('click', () => {
  grid.showMoreRows();
  updateStatus();
  updateInspector();
});

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
