/**
 * Cell and pane markup — the ONE serializer shared by both write paths.
 *
 * The surgical writer (`writer/sheet-writer.ts`) rewrites cells inside an
 * existing sheet, the incremental writer (`writer/sheet-stream-writer.ts`)
 * creates a sheet from scratch; both must emit byte-identical `<c>` markup for
 * the same value + style, so the rendering lives here rather than twice.
 *
 * Values are expected to be already validated (`assertWritableValue` on the
 * writer side); escaping, `xml:space`, dates, formulas and rich text are
 * handled here.
 */
import { dateToSerial } from '../model/dates';
import type { CellValue, RichTextRun } from '../model/value';
import type { FontProps } from '../model/style';
import type { SheetViewProps } from '../model/view';
import { coordsToRef } from '../model/refs';
import { escapeXml } from '../xml/escape';
import { assertFont } from './style-validate';

/**
 * `<rPr>` children for one rich-text run (run fonts use `rFont`, not `name`).
 * A run font is a cell font: the same size/colour rules apply, so `size` and
 * `color` are validated (an ARGB colour needs no escaping) while the free-form
 * font name is escaped.
 */
function runPrXml(font: FontProps): string {
  assertFont(font);
  let xml = '';
  if (font.bold) xml += '<b/>';
  if (font.italic) xml += '<i/>';
  if (font.underline) xml += '<u/>';
  if (font.size !== undefined) xml += `<sz val="${font.size}"/>`;
  if (font.color) xml += `<color rgb="${font.color}"/>`;
  if (font.name) xml += `<rFont val="${escapeXml(font.name)}"/>`;
  return xml ? `<rPr>${xml}</rPr>` : '';
}

/** `<r>` runs of a rich-text inline string. */
function richTextXml(runs: RichTextRun[]): string {
  if (runs.length === 0) return '<t/>';
  return runs
    .map((run) => {
      const text = run.text ?? '';
      const preserve = /^\s|\s$|[\n\t]/.test(text) ? ' xml:space="preserve"' : '';
      return `<r>${run.font ? runPrXml(run.font) : ''}<t${preserve}>${escapeXml(text)}</t></r>`;
    })
    .join('');
}

/**
 * Serialize a single cell into OOXML `<c>` markup. `''` removes the cell.
 * Strings are written INLINE (`t="inlineStr"`): a shared-string table grows
 * with the number of distinct strings, which would make a streaming writer's
 * memory unbounded by definition.
 */
export function generateCellXml(
  ref: string,
  value: CellValue,
  styleIndex: number,
  date1904: boolean,
): string {
  if (value === null || value === undefined) return '';
  const sAttr = styleIndex ? ` s="${styleIndex}"` : '';

  if (typeof value === 'number') return `<c r="${ref}"${sAttr}><v>${value}</v></c>`;
  if (typeof value === 'boolean') return `<c r="${ref}" t="b"${sAttr}><v>${value ? 1 : 0}</v></c>`;
  if (value instanceof Date) return `<c r="${ref}"${sAttr}><v>${dateToSerial(value, date1904)}</v></c>`;
  if (typeof value === 'object' && 'error' in value) {
    return `<c r="${ref}" t="e"${sAttr}><v>${escapeXml(value.error)}</v></c>`;
  }
  if (typeof value === 'object' && 'formula' in value) {
    const { result } = value;
    let resultText = '';
    if (result instanceof Date) resultText = String(dateToSerial(result, date1904));
    else if (typeof result === 'number' || typeof result === 'string' || typeof result === 'boolean') {
      resultText = String(result);
    }
    return `<c r="${ref}"${sAttr}><f>${escapeXml(value.formula)}</f><v>${escapeXml(resultText)}</v></c>`;
  }
  if (typeof value === 'object' && 'richText' in value) {
    return `<c r="${ref}" t="inlineStr"${sAttr}><is>${richTextXml(value.richText)}</is></c>`;
  }

  const str = String(value);
  const preserve = /^\s|\s$|[\n\t]/.test(str) ? ' xml:space="preserve"' : '';
  return `<c r="${ref}" t="inlineStr"${sAttr}><is><t${preserve}>${escapeXml(str)}</t></is></c>`;
}

/** `<pane>` element for a sheet view, or `null` when the sheet is not frozen. */
export function paneXmlFor(view: SheetViewProps | undefined): string | null {
  if (!view) return null;
  const state = view.state ?? 'normal';
  if (state === 'normal') return null;
  const x = Math.max(0, Math.floor(view.xSplit ?? 0));
  const y = Math.max(0, Math.floor(view.ySplit ?? 0));
  if (x === 0 && y === 0) return null;
  const attrs = [
    x > 0 ? `xSplit="${x}"` : '',
    y > 0 ? `ySplit="${y}"` : '',
    `topLeftCell="${escapeXml(view.topLeftCell ?? coordsToRef(y + 1, x + 1))}"`,
    `activePane="${x > 0 && y > 0 ? 'bottomRight' : x > 0 ? 'topRight' : 'bottomLeft'}"`,
  ].filter(Boolean);
  if (state !== 'split') attrs.push(`state="${state}"`);
  return `<pane ${attrs.join(' ')}/>`;
}