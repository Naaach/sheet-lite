import { PartStore } from '../opc/part-store';
import { readZip, writeZip } from '../zip/archive';
import { parseContentTypes, type ContentTypes } from '../opc/content-types';

function toUint8(input: Uint8Array | ArrayBuffer | ArrayBufferView): Uint8Array {
  if (input instanceof Uint8Array) return input;
  if (ArrayBuffer.isView(input)) {
    return new Uint8Array(input.buffer, input.byteOffset, input.byteLength);
  }
  return new Uint8Array(input);
}

/**
 * The public entry point. v0.1 covers loading a package and writing it back
 * losslessly (patch-only): if nothing changed, the original bytes are returned.
 */
export class Workbook {
  private constructor(private readonly store: PartStore) {}

  /** Read a workbook from a filesystem path (Node only). */
  static async readFile(path: string): Promise<Workbook> {
    const { readFile } = await import('node:fs/promises');
    return Workbook.read(await readFile(path));
  }

  /** Read a workbook from raw bytes (works in Node, browsers and edge runtimes). */
  static read(input: Uint8Array | ArrayBuffer | ArrayBufferView): Workbook {
    return new Workbook(new PartStore(readZip(toUint8(input))));
  }

  /** Write to a filesystem path (Node only). */
  async writeFile(path: string): Promise<void> {
    const { writeFile } = await import('node:fs/promises');
    await writeFile(path, this.toBuffer());
  }

  /**
   * Serialize the workbook.
   * Patch-only: when no part was modified, the original bytes are returned
   * verbatim, which guarantees a perfect no-op round-trip.
   */
  toBuffer(): Uint8Array {
    if (!this.store.isDirty()) return this.store.raw;
    return writeZip(this.store.toArchive());
  }

  /** Parsed content types of the package. */
  contentTypes(): ContentTypes {
    const part = this.store.get('[Content_Types].xml');
    return part ? parseContentTypes(part) : { defaults: new Map(), overrides: new Map() };
  }

  /** Internal access (tests / advanced use). */
  parts(): PartStore {
    return this.store;
  }
}
