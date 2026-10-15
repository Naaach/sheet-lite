import type { CellValue, FontProps, RichTextValue } from '../../src/index';

/** Human-readable byte size, e.g. "12.4 kB". */
export function formatBytes(n: number): string {
  if (!Number.isFinite(n)) return '—';
  if (Math.abs(n) < 1024) return `${n} B`;
  if (Math.abs(n) < 1024 * 1024) return `${(n / 1024).toFixed(1)} kB`;
  return `${(n / (1024 * 1024)).toFixed(2)} MB`;
}

/** Signed delta like "+412 B" / "−1.2 kB". */
export function formatDelta(delta: number): string {
  const sign = delta > 0 ? '+' : delta < 0 ? '−' : '±';
  return `${sign}${formatBytes(Math.abs(delta))}`;
}

const pad2 = (n: number): string => String(n).padStart(2, '0');

/** Local-timezone ISO-ish date, e.g. "2026-10-14". */
export function displayDate(d: Date): string {
  if (Number.isNaN(d.getTime())) return 'invalid date';
  return `${d.getFullYear()}-${pad2(d.getMonth() + 1)}-${pad2(d.getDate())}`;
}

/**
 * Flattened text of a rich-text value. Reading a cell with `<r>` runs gives
 * `{ richText }` instead of a plain string, so every value formatter has to
 * collapse the runs before it can print anything.
 */
export function richTextText(value: RichTextValue): string {
  return value.richText.map((run) => run.text).join('');
}

/**
 * Grid text for a cell value. Numbers with a `#,##0.00` format get thousand
 * separators so the grid feels like a spreadsheet, formulas show their
 * computed result (the raw formula lives in the tooltip/inspector).
 */
export function displayValue(value: CellValue, numFmt = 'General'): string {
  if (value === null || value === undefined) return '';
  if (typeof value === 'object' && 'richText' in value) return richTextText(value);
  if (typeof value === 'number') {
    if (numFmt === '#,##0.00') {
      return value.toLocaleString('en-US', {
        minimumFractionDigits: 2,
        maximumFractionDigits: 2,
      });
    }
    return String(value);
  }
  if (typeof value === 'boolean') return value ? 'TRUE' : 'FALSE';
  if (value instanceof Date) return displayDate(value);
  if (typeof value === 'object' && 'error' in value) return `#${value.error}`;
  if (typeof value === 'object' && 'formula' in value) {
    const { result } = value;
    if (result === undefined || result === null) return '';
    return displayValue(result, numFmt);
  }
  return String(value);
}

/** Raw text to seed the cell editor (formulas come back as `=...`). */
export function editSeed(value: CellValue): string {
  if (value === null || value === undefined) return '';
  if (typeof value === 'object' && 'richText' in value) return richTextText(value);
  if (typeof value === 'object' && 'formula' in value) return `=${value.formula}`;
  if (typeof value === 'number') return String(value);
  if (typeof value === 'boolean') return value ? 'TRUE' : 'FALSE';
  if (value instanceof Date) return displayDate(value);
  if (typeof value === 'object' && 'error' in value) return `#${value.error}`;
  return String(value);
}

/**
 * Turn typed grid text back into a `CellValue`:
 * empty → null, numbers → number, TRUE/FALSE → boolean, `=…` → formula,
 * ISO dates → Date, everything else → string.
 */
export function parseCellInput(text: string): CellValue {
  if (text === '') return null;
  const t = text.trim();
  if (t.startsWith('=')) return { formula: t.slice(1) };
  if (/^true$/i.test(t)) return true;
  if (/^false$/i.test(t)) return false;
  if (/^\d{4}-\d{2}-\d{2}$/.test(t)) {
    const d = new Date(`${t}T00:00:00`);
    if (!Number.isNaN(d.getTime())) return d;
  }
  if (/^-?[\d,]*\.?\d+(e[+-]?\d+)?$/i.test(t)) {
    const n = Number(t.replace(/,/g, ''));
    if (Number.isFinite(n)) return n;
  }
  return text;
}

/** Column label for a 1-based index: 1 → "A", 27 → "AA". */
export function columnLabel(col: number): string {
  let label = '';
  let n = col;
  while (n > 0) {
    const rem = (n - 1) % 26;
    label = String.fromCharCode(65 + rem) + label;
    n = Math.floor((n - 1) / 26);
  }
  return label;
}

/**
 * CSS declarations for a font, used to paint rich-text runs. ARGB colors keep
 * their last 6 digits (RGB), exactly like the fill swatch in the inspector.
 */
export function fontStyle(font: FontProps | undefined): string {
  if (!font) return '';
  const declarations: string[] = [];
  if (font.bold) declarations.push('font-weight:650');
  if (font.italic) declarations.push('font-style:italic');
  if (font.underline) declarations.push('text-decoration:underline');
  if (font.size !== undefined) declarations.push(`font-size:${font.size}pt`);
  if (font.color) declarations.push(`color:#${font.color.slice(-6)}`);
  return declarations.join(';');
}

/**
 * Approximate CSS pixels for an Excel column width / row height so the grid
 * previews what the file will look like. Excel sizes columns in characters of
 * the default font (~7 px each, +5 px padding) and rows in points.
 * Both are clamped so a silly value cannot break the preview.
 */
export function excelWidthPx(width: number): number {
  return Math.min(Math.max(Math.round(width * 7 + 5), 24), 520);
}

export function rowHeightPx(points: number): number {
  return Math.min(Math.max(Math.round((points * 96) / 72), 12), 300);
}

/** Minimal HTML escaping for injected cell text. */
const ENTITIES: Array<[string, string]> = [
  ['&', '&' + 'amp;'],
  ['<', '&' + 'lt;'],
  ['>', '&' + 'gt;'],
  ['"', '&' + 'quot;'],
  ["'", '&' + '#39;'],
];

export function escapeHtml(s: string): string {
  let out = s;
  for (const [ch, entity] of ENTITIES) out = out.replaceAll(ch, entity);
  return out;
}
