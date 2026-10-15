/**
 * XML entity escaping / decoding.
 *
 * A dependency-free leaf module on purpose: the eager reader/patch writer, the
 * streaming reader and the incremental writer all escape or decode text, and
 * none of them should pull a DOM (or a ZIP layer) in just to share these.
 */

const NAMED: Record<string, string> = {
  amp: '&',
  lt: '<',
  gt: '>',
  quot: '"',
  apos: "'",
};

/** Resolve XML entities (named + numeric); out-of-range code points become U+FFFD. */
export function decodeEntities(value: string): string {
  if (!value.includes('&')) return value;
  return value.replace(/&(#x?[0-9a-fA-F]+|[a-zA-Z]+);/g, (whole, entity: string) => {
    const toChar = (codePoint: number): string =>
      Number.isInteger(codePoint) && codePoint >= 0 && codePoint <= 0x10ffff
        ? String.fromCodePoint(codePoint)
        : '\uFFFD';
    if (entity.startsWith('#x') || entity.startsWith('#X')) {
      return toChar(parseInt(entity.slice(2), 16));
    }
    if (entity.startsWith('#')) {
      return toChar(parseInt(entity.slice(1), 10));
    }
    return NAMED[entity] ?? whole;
  });
}

/** Escape a value for use in XML text or attribute content. */
export function escapeXml(value: string): string {
  return value
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&apos;');
}

/**
 * Characters XML 1.0 forbids outright: the C0 controls except TAB/LF/CR, plus
 * the U+FFFE / U+FFFF non-characters. They are legal JS strings and legal in
 * most other formats, but an .xlsx part containing one is a file Excel offers
 * to repair.
 */
const ILLEGAL_XML_CHARS = /[\u0000-\u0008\u000B\u000C\u000E-\u001F\uFFFE\uFFFF]/;

/**
 * The first character of `value` XML forbids, as a `"U+0001"` code point, or
 * `null` when the string is writable as XML text.
 *
 * One rule for every writer and for the worksheet name: what one write path
 * accepts must never be "repaired" by Excel because the other path let it
 * through.
 */
export function illegalXmlChar(value: string): string | null {
  const match = ILLEGAL_XML_CHARS.exec(value);
  if (!match) return null;
  return `U+${match[0].charCodeAt(0).toString(16).padStart(4, '0').toUpperCase()}`;
}