import { Workbook } from '../../src/index';

export type PartStatus = 'identical' | 'changed' | 'added' | 'removed';

export interface PartRow {
  name: string;
  status: PartStatus;
  size: number;
}

export interface PatchReport {
  /** Bytes `toBuffer()` would produce right now (the file that gets downloaded). */
  output: Uint8Array;
  originalSize: number;
  outputSize: number;
  parts: PartRow[];
  identicalCount: number;
  changedCount: number;
  addedCount: number;
  removedCount: number;
  /** Total part count before the edits. */
  totalParts: number;
  /** True when `toBuffer()` returned the original bytes verbatim (no edits at all). */
  untouched: boolean;
  /** Set when re-reading the output failed (should not happen; surfaced honestly). */
  parseError: string | null;
}

/** Copy of every part name → bytes at load time, the "before" snapshot. */
export function snapshotParts(wb: Workbook): Map<string, Uint8Array> {
  const store = wb.parts();
  const snapshot = new Map<string, Uint8Array>();
  for (const name of store.names()) {
    const data = store.get(name);
    if (data) snapshot.set(name, data.slice());
  }
  return snapshot;
}

function sameBytes(a: Uint8Array, b: Uint8Array): boolean {
  if (a === b) return true;
  if (a.byteLength !== b.byteLength) return false;
  for (let i = 0; i < a.length; i++) {
    if (a[i] !== b[i]) return false;
  }
  return true;
}

/**
 * Serialize the workbook and compare every OPC part against the load-time
 * snapshot. The output is re-read through `Workbook.read` so the comparison
 * uses the same public API on both sides — no private internals.
 */
export function buildPatchReport(
  originalBytes: Uint8Array,
  originalParts: Map<string, Uint8Array>,
  wb: Workbook,
): PatchReport {
  const output = wb.toBuffer();
  const untouched = output === originalBytes; // clean workbooks pass bytes through verbatim

  let parsed: Workbook | null = null;
  let parseError: string | null = null;
  try {
    parsed = Workbook.read(output);
  } catch (error) {
    parseError = error instanceof Error ? error.message : String(error);
  }

  const parts: PartRow[] = [];
  let identicalCount = 0;
  let changedCount = 0;
  let addedCount = 0;

  if (parsed) {
    const store = parsed.parts();
    const now = new Set(store.names());

    for (const [name, before] of originalParts) {
      if (now.has(name)) {
        const after = store.get(name)!;
        const identical = sameBytes(before, after);
        parts.push({ name, status: identical ? 'identical' : 'changed', size: after.byteLength });
        if (identical) identicalCount++;
        else changedCount++;
      } else {
        parts.push({ name, status: 'removed', size: 0 });
      }
    }
    for (const name of store.names()) {
      if (!originalParts.has(name)) {
        parts.push({ name, status: 'added', size: store.get(name)!.byteLength });
        addedCount++;
      }
    }
  }

  return {
    output,
    originalSize: originalBytes.byteLength,
    outputSize: output.byteLength,
    parts,
    identicalCount,
    changedCount,
    addedCount,
    removedCount: parts.length - identicalCount - changedCount - addedCount,
    totalParts: originalParts.size,
    untouched,
    parseError,
  };
}
