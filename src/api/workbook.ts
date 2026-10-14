import { strToU8, strFromU8 } from 'fflate';
import { PartStore } from '../opc/part-store';
import { readZip, writeZip } from '../zip/archive';
import { SheetLiteError } from '../zip/errors';
import { parseContentTypes, type ContentTypes } from '../opc/content-types';
import { readPackageMeta, type PackageMeta } from '../reader/package-reader';
import { StyleRegistry } from '../writer/style-registry';
import type { StreamingWorkbook, StreamInput } from './streaming';
import type { ReadLimits } from '../zip/raw';
import { Worksheet } from './worksheet';

function toUint8(input: Uint8Array | ArrayBuffer | ArrayBufferView): Uint8Array {
  if (input instanceof Uint8Array) return input;
  if (ArrayBuffer.isView(input)) {
    return new Uint8Array(input.buffer, input.byteOffset, input.byteLength);
  }
  return new Uint8Array(input);
}

const toPartKey = (partName: string): string => partName.replace(/^\//, '');

/**
 * The public entry point. Loads an OPC package and writes it back
 * losslessly: only the parts that actually changed are re-serialized.
 */
export class Workbook {
  private readonly store: PartStore;
  private readonly meta: PackageMeta;
  private readonly styleRegistry: StyleRegistry | null;
  private readonly worksheetList: Worksheet[];

  private constructor(store: PartStore) {
    this.store = store;
    this.meta = readPackageMeta(store);
    const stylesData = store.get('xl/styles.xml');
    this.styleRegistry = stylesData ? new StyleRegistry(strFromU8(stylesData)) : null;
    this.worksheetList = this.meta.sheets.map((sheet) => {
      const key = toPartKey(sheet.partName);
      const xml = strFromU8(store.get(key)!);
      return new Worksheet(sheet.name, Number(sheet.sheetId), key, xml, this.meta, this.styleRegistry);
    });
  }

  static async readFile(path: string): Promise<Workbook> {
    const { readFile } = await import('node:fs/promises');
    return Workbook.read(await readFile(path));
  }

  static read(input: Uint8Array | ArrayBuffer | ArrayBufferView): Workbook {
    return new Workbook(new PartStore(readZip(toUint8(input))));
  }

  /**
   * Read a workbook from a `Blob`/`File` — browser, Deno and edge Workers
   * (also available in Node, which has a global `Blob` since v18).
   */
  static async readBlob(blob: Blob): Promise<Workbook> {
    return Workbook.read(await blob.arrayBuffer());
  }

  /**
   * Streaming, READ-ONLY open (Node file path, or an in-memory buffer).
   *
   * Reads only the small package parts up front and lets you iterate rows one
   * at a time with bounded memory — see `StreamingWorkbook` in the
   * `sheet-lite/stream` entry point, which is loaded on demand here.
   * Editing requires the eager `Workbook.read`/`readFile`/`readBlob` APIs.
   */
  static async openStream(input: StreamInput, limits?: Partial<ReadLimits>): Promise<StreamingWorkbook> {
    const { StreamingWorkbook: SW } = await import('./streaming');
    return SW.open(input, limits);
  }

  async writeFile(path: string): Promise<void> {
    const { writeFile } = await import('node:fs/promises');
    await writeFile(path, this.toBuffer());
  }

  /** Serialize the workbook as a `Blob` (browser / Deno / Workers / Node ≥ 18). */
  toBlob(): Blob {
    return new Blob([this.toBuffer() as Uint8Array<ArrayBuffer>]);
  }

  /**
   * Browser only: trigger a client-side download of the workbook.
   * Throws a `SheetLiteError` (code `ENV_UNSUPPORTED`) outside a browser.
   */
  download(filename: string): void {
    if (typeof document === 'undefined') {
      throw new SheetLiteError(
        'Workbook.download() is only available in a browser runtime (no `document` found). Use toBlob() instead.',
        'ENV_UNSUPPORTED',
      );
    }
    const url = URL.createObjectURL(this.toBlob());
    const anchor = document.createElement('a');
    anchor.href = url;
    anchor.download = filename;
    document.body.appendChild(anchor);
    anchor.click();
    anchor.remove();
    setTimeout(() => URL.revokeObjectURL(url), 0);
  }

  get worksheets(): Worksheet[] {
    return this.worksheetList;
  }

  get sheetNames(): string[] {
    return this.worksheetList.map((sheet) => sheet.name);
  }

  getSheetNames(): string[] {
    return this.sheetNames;
  }

  getWorksheet(nameOrId: string | number): Worksheet | undefined {
    if (typeof nameOrId === 'number') {
      return this.worksheetList.find((sheet) => sheet.id === nameOrId);
    }
    return this.worksheetList.find((sheet) => sheet.name === nameOrId);
  }

  get metadata(): PackageMeta {
    return this.meta;
  }

  contentTypes(): ContentTypes {
    const part = this.store.get('[Content_Types].xml');
    return part ? parseContentTypes(part) : { defaults: new Map(), overrides: new Map() };
  }

  /**
   * Serialize the workbook.
   * Patch-only: unchanged sheets are carried over untouched; if nothing at all
   * changed, the original bytes are returned verbatim.
   */
  toBuffer(): Uint8Array {
    const dirtySheets = this.worksheetList.filter((sheet) => sheet.isDirty());
    if (dirtySheets.length === 0 && !this.store.isDirty()) {
      return this.store.raw;
    }

    const store = this.store.clone();
    for (const sheet of dirtySheets) {
      store.set(sheet.partName, strToU8(sheet.serialize(this.styleRegistry)));
    }
    if (this.styleRegistry?.changed) {
      store.set('xl/styles.xml', strToU8(this.styleRegistry.toXml()));
    }
    return writeZip(store.toArchive());
  }

  /** Low-level part access (advanced use / tests). */
  parts(): PartStore {
    return this.store;
  }
}
