/**
 * ZIP entry-name safety.
 *
 * Kept in its own tiny module because BOTH the eager container (`zip/raw.ts`)
 * and the streaming container (`zip/zip-stream.ts`) must apply exactly the same
 * rule — and because the writer's chunk must not drag the whole eager ZIP layer
 * in just to reuse one predicate.
 */

/** Signatures of a path-traversal / absolute / NUL-bearing ZIP entry name. */
const UNSAFE = /(^\/)|(^[a-zA-Z]:)|(^|\/)\.\.(\/|$)|(\0)/;

/** ZIP names always use forward slashes. */
export function normalizeName(name: string): string {
  return name.replace(/\\/g, '/');
}

/**
 * True when a name may be written into (or read from) an archive: no absolute
 * paths, no drive letters, no `..` segments, no NUL, no directory markers.
 */
export function isSafeEntryName(name: string): boolean {
  if (name.length === 0) return false;
  const normalized = normalizeName(name);
  if (normalized.endsWith('/')) return false; // directory markers, not needed
  return !UNSAFE.test(normalized);
}