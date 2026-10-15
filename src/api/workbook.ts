import { strToU8, strFromU8 } from 'fflate';
import { PartStore } from '../opc/part-store';
import { readZip, writeZip } from '../zip/archive';
import { SheetLiteError } from '../zip/errors';
import { parseContentTypes, type ContentTypes } from '../opc/content-types';
import { pruneOrphanParts } from '../opc/orphan-prune';
import { readPackageMeta, type PackageMeta, type SheetRef } from '../reader/package-reader';
import { StyleRegistry } from '../writer/style-registry';
import type { StreamingWorkbook, StreamInput } from './streaming';
import type { WorkbookWriter, WorkbookWriterOptions } from './write-workbook';
import type { ReadLimits } from '../zip/raw';
import { assertNotCompoundFile } from '../zip/raw';
import type { SheetViewProps } from '../model/view';
import {
  MAIN_CONTENT_TYPE,
  MINIMAL_STYLES,
  SHEET_CONTENT_TYPE,
  SHEET_NS,
  STYLES_CONTENT_TYPE,
  STYLES_REL,
  WORKSHEET_REL,
  XML_DECL,
  minimalContentTypes,
  minimalWorkbookRels,
} from '../opc/minimal-package';
import { validateSheetName } from '../opc/sheet-name';
import {
  addOverride,
  addRelationship,
  addSheetEntry,
  maxRelationshipId,
  maxSheetId,
  relsPathForPart,
  removeOverride,
  removeRelationship,
  removeRelationshipByTarget,
  removeSheetEntry,
} from '../opc/package-edit';
import { Worksheet } from './worksheet';

function toUint8(input: Uint8Array | ArrayBuffer | ArrayBufferView): Uint8Array {
  if (input instanceof Uint8Array) return input;
  if (ArrayBuffer.isView(input)) {
    return new Uint8Array(input.buffer, input.byteOffset, input.byteLength);
  }
  return new Uint8Array(input);
}

const toPartKey = (partName: string): string => partName.replace(/^\//, '');

/** First sheet of a workbook without any: shared strings are written inline. */
const MINIMAL_SHEET_XML =
  XML_DECL +
  `<worksheet ${SHEET_NS}>` +
  '<dimension ref="A1"/><sheetData/></worksheet>';

/**
 * The public entry point. Loads an OPC package and writes it back
 * losslessly: only the parts that actually changed are re-serialized.
 */
export class Workbook {
  private readonly store: PartStore;
  private readonly meta: PackageMeta;
  private styleRegistry: StyleRegistry | null;
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
    const bytes = toUint8(input);
    // A legacy `.xls` / encrypted OOXML container is not a broken archive.
    assertNotCompoundFile(bytes);
    return new Workbook(new PartStore(readZip(bytes)));
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

  /**
   * Create a NEW workbook that is written incrementally, with bounded memory
   * (ExcelJS-style: build rows one by one, then `commit()`/`toBuffer()`).
   *
   * ```ts
   * const wb = await Workbook.create({ file: './huge.xlsx' });
   * const ws = wb.addWorksheet('Data');
   * for (let i = 0; i < 100_000; i++) ws.addRow([i, `row ${i}`]);
   * await wb.commit();
   * ```
   *
   * Unlike this class, the returned writer never holds the worksheet XML (nor
   * the archive) in memory: rows are deflated and streamed straight to the
   * target. It lives in the separate `sheet-lite/write` entry point and is
   * loaded on demand here, so the eager bundle keeps its size budget.
   *
   * To EDIT an existing workbook use `Workbook.read()` / `readFile()`.
   */
  static async create(options?: WorkbookWriterOptions): Promise<WorkbookWriter> {
    const { WorkbookWriter: Writer } = await import('../write');
    return Writer.create(options);
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

  /**
   * Add a new, empty worksheet to the workbook (ExcelJS-style).
   *
   * Creates the sheet part, its content-type override, a workbook relationship
   * and the `<sheets>` entry — registering only what is missing, so every
   * untouched part stays byte-identical. Strings written into the new sheet
   * are stored inline (`inlineStr`); a missing `styles.xml` is created as a
   * minimal valid style sheet.
   *
   * Throws `SHEET_EXISTS` for a duplicate name and `BAD_SHEET_NAME` when the
   * name is empty, longer than 31 characters or contains `: \ / ? * [ ]`.
   */
  addWorksheet(name: string, options?: { views?: SheetViewProps[] }): Worksheet {
    validateSheetName(name);
    if (this.worksheetList.some((sheet) => sheet.name === name)) {
      throw new SheetLiteError(`A worksheet named "${name}" already exists`, 'SHEET_EXISTS');
    }

    // 1. The sheet part. Pick the first free `sheetN.xml` slot.
    let slot = 1;
    while (this.store.has(`xl/worksheets/sheet${slot}.xml`)) slot += 1;
    const partKey = `xl/worksheets/sheet${slot}.xml`;
    this.store.set(partKey, strToU8(MINIMAL_SHEET_XML));

    // 2. A missing styles.xml is created (Excel requires fonts/fills/cellXfs).
    const createdStyles = !this.store.has('xl/styles.xml');
    if (createdStyles) {
      this.store.set('xl/styles.xml', strToU8(MINIMAL_STYLES));
      this.styleRegistry = new StyleRegistry(MINIMAL_STYLES);
    }

    // 3. [Content_Types].xml override(s).
    const ctKey = '[Content_Types].xml';
    let ct = this.store.has(ctKey) ? strFromU8(this.store.get(ctKey)!) : minimalContentTypes();
    ct = addOverride(ct, `/${partKey}`, SHEET_CONTENT_TYPE);
    if (createdStyles) ct = addOverride(ct, '/xl/styles.xml', STYLES_CONTENT_TYPE);
    this.store.set(ctKey, strToU8(ct));

    // 4. Workbook relationship with a fresh rId.
    const relsKey = 'xl/_rels/workbook.xml.rels';
    let rels = this.store.has(relsKey) ? strFromU8(this.store.get(relsKey)!) : minimalWorkbookRels();
    const rId = `rId${maxRelationshipId(rels) + 1}`;
    rels = addRelationship(rels, rId, WORKSHEET_REL, `worksheets/sheet${slot}.xml`);
    if (createdStyles) {
      rels = addRelationship(rels, `rId${maxRelationshipId(rels) + 1}`, STYLES_REL, 'styles.xml');
    }
    this.store.set(relsKey, strToU8(rels));

    // 5. The <sheets> entry with a unique sheetId.
    const wbKey = 'xl/workbook.xml';
    const wbXml = strFromU8(this.store.get(wbKey)!);
    const sheetId = maxSheetId(wbXml) + 1;
    this.store.set(wbKey, strToU8(addSheetEntry(wbXml, name, sheetId, rId)));

    // 6. Register the sheet in the in-memory model.
    const sheetRef: SheetRef = { name, sheetId: String(sheetId), rId, partName: `/${partKey}` };
    this.meta.sheets.push(sheetRef);
    const worksheet = new Worksheet(name, sheetId, partKey, MINIMAL_SHEET_XML, this.meta, this.styleRegistry);
    this.worksheetList.push(worksheet);
    if (options?.views) worksheet.views = options.views;
    return worksheet;
  }

  /**
   * Remove a worksheet by name or sheetId: the sheet part, its relationships
   * file (if any), the workbook entry + relationship and the content-type
   * override. A stale `calcChain.xml` (it indexes sheets by position) is
   * dropped as well — Excel rebuilds it; defined names bound to the removed
   * sheet are removed with it.
   *
   * By default the parts ONLY that sheet owned are pruned too: the walk
   * follows its rels (drawing, legacy VML, comments, table, printer settings,
   * hyperlinks) and recursively the rels of everything they reference (a
   * drawing owns its charts, a chart its embedded workbook), dropping those
   * parts, their `.rels` files and their content-type entries. A part that any
   * surviving relationship still points at — another sheet's drawing, an image
   * two drawings share, `styles.xml`, `sharedStrings.xml`, a pivot cache — is
   * kept, as is anything ambiguous: an orphan is invisible to Excel, a
   * dangling target is a repair prompt. Parts that survive stay byte-identical.
   *
   * @param options - `prune: false` keeps the orphaned parts in the package
   * (the historical behaviour: nothing is deleted except the sheet itself and
   * what would otherwise go stale). Default `true`.
   *
   * Removing the LAST remaining sheet is rejected with `LAST_SHEET` (a
   * workbook must keep at least one sheet for Excel to open it).
   */
  removeWorksheet(nameOrId: string | number, options?: { prune?: boolean }): void {
    const index = this.worksheetList.findIndex((sheet) =>
      typeof nameOrId === 'number' ? sheet.id === nameOrId : sheet.name === nameOrId,
    );
    const worksheet = this.worksheetList[index];
    if (!worksheet) {
      throw new SheetLiteError(`No worksheet "${String(nameOrId)}" found`, 'SHEET_NOT_FOUND');
    }
    if (this.worksheetList.length === 1) {
      throw new SheetLiteError(
        'Cannot remove the last remaining worksheet: a workbook must keep at least one sheet',
        'LAST_SHEET',
      );
    }

    const partKey = worksheet.partName;
    const refIndex = this.meta.sheets.findIndex((sheet) => toPartKey(sheet.partName) === partKey);
    const ref = this.meta.sheets[refIndex];

    // First, while the sheet and its rels are still in the store: whatever
    // ONLY it could reach dies with it.
    if (options?.prune ?? true) pruneOrphanParts(this.store, partKey);

    this.store.remove(partKey);
    const relsFile = relsPathForPart(partKey);
    if (this.store.has(relsFile)) this.store.remove(relsFile);

    // calcChain references sheets by position and would go stale; Excel and
    // LibreOffice rebuild it on load.
    const calcChainKey = 'xl/calcChain.xml';
    const removedCalcChain = this.store.has(calcChainKey);
    if (removedCalcChain) this.store.remove(calcChainKey);

    const wbKey = 'xl/workbook.xml';
    const wbXml = strFromU8(this.store.get(wbKey)!);
    if (ref) {
      this.store.set(wbKey, strToU8(removeSheetEntry(wbXml, ref.rId, worksheet.name)));
    }

    const relsKey = 'xl/_rels/workbook.xml.rels';
    if (this.store.has(relsKey)) {
      let rels = strFromU8(this.store.get(relsKey)!);
      if (ref) rels = removeRelationship(rels, ref.rId);
      if (removedCalcChain) rels = removeRelationshipByTarget(rels, 'calcChain.xml');
      this.store.set(relsKey, strToU8(rels));
    }

    const ctKey = '[Content_Types].xml';
    if (this.store.has(ctKey)) {
      let ct = strFromU8(this.store.get(ctKey)!);
      ct = removeOverride(ct, `/${partKey}`);
      if (removedCalcChain) ct = removeOverride(ct, `/${calcChainKey}`);
      this.store.set(ctKey, strToU8(ct));
    }

    if (refIndex !== -1) this.meta.sheets.splice(refIndex, 1);
    this.worksheetList.splice(index, 1);
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

