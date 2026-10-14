# sheet-lite

> Read, write and **edit** Excel (`.xlsx`) without breaking your file — lightweight, secure and with **zero transitive dependencies**.

When you open an Excel file, change a value and save, most libraries rebuild the whole package and can silently drop formatting, data validations, conditional formatting or images. `sheet-lite` is built to **patch only what changed** and leave everything else untouched.

📖 **Documentation site:** [`docs-site/`](docs-site/index.html) · background & decisions: [`docs/STATUS.md`](docs/STATUS.md)

## Status

**v0.5 — read, edit, style, sheets — plus streaming reads for large files.**

- [x] `Workbook.read()` / `readFile()` / `toBuffer()` / `writeFile()`
- [x] Value reading: strings, numbers, booleans, dates (1900/1904), formulas (cached result), inline & shared strings
- [x] `Worksheet` / `Cell` / `Row` API: `getWorksheet`, `getCell`, `getRow`, `eachRow`, `addRow`
- [x] **Editing without breaking:** only changed cells/merges are rewritten; other parts stay byte-identical
- [x] **Sheets:** `wb.addWorksheet(name, options?)` / `wb.removeWorksheet(nameOrId)` — full OPC registration, works on files without `styles.xml`
- [x] **Columns & rows:** `ws.getColumn(n).width` / `.hidden`, `row.height` — surgical `<cols>` / `<row ht>` patching
- [x] **Frozen panes:** `ws.views = [{ state: 'frozen', xSplit: 2, ySplit: 3 }]` (ExcelJS-compatible, read back on re-open)
- [x] **Styles:** `cell.numFmt`, `cell.font` (bold/italic/size/name/color), `cell.fill` (solid/pattern), `cell.border` (all 13 OOXML line styles, deduped in `styles.xml`), `cell.alignment` (horizontal/vertical/wrap/indent/rotation…)
- [x] **Rich text:** read **and** write `{ richText: [{ text, font? }, …] }` — inline `<r>` runs on write, `{ richText }` on read from inline and shared strings; plain strings stay plain
- [x] **Merges:** `worksheet.mergeCells('A1:C1')` / `unmergeCells`
- [x] Patch-only write: if nothing changed, the original bytes are returned verbatim
- [x] Security by design: path-traversal, entry-count, size and compression-ratio limits
- [x] **Streaming reads** for large files: `sheet-lite/stream` (read-only) walks rows with a bounded memory ceiling — ~10 MB instead of ~673 MB for a 100k × 8 sheet
- [x] Browser/edge adapters: `readBlob` / `toBlob` / `download`

## Install

```bash
npm install sheet-lite
```

## Usage

```ts
import { Workbook } from 'sheet-lite';

const wb = await Workbook.readFile('./report.xlsx');
const ws = wb.getWorksheet('Sheet1')!;

// read
console.log(ws.getCell('A1').value, ws.getCell('A1').type);

// edit — only this cell is rewritten
ws.getCell('B2').value = 1234.56;
ws.getCell('B2').numFmt = '#,##0.00';
ws.getCell('A1').font = { bold: true };
ws.getCell('A3').value = new Date();
ws.mergeCells('C1:D1');
ws.addRow(['nuevo', 7]);

// borders + alignment + rich text (ExcelJS-compatible shapes)
ws.getCell('C3').border = {
  top: { style: 'thin', color: { argb: 'FF00B0F0' } },
  bottom: { style: 'double' },
};
ws.getCell('C3').alignment = { horizontal: 'center', vertical: 'center', wrapText: true };
ws.getCell('D4').value = {
  richText: [
    { text: 'Total: ' },
    { text: '1 234 €', font: { bold: true, color: 'FFFF0000' } },
  ],
};

// structure — add/remove sheets, columns, rows, frozen panes
const summary = wb.addWorksheet('Summary');
summary.getCell('A1').value = 'Total';
summary.getColumn('A').width = 24;
summary.getRow(1).height = 22;
summary.views = [{ state: 'frozen', xSplit: 0, ySplit: 1 }];
wb.removeWorksheet('OldSheet');

await wb.writeFile('./report-edited.xlsx');
// data validations, conditional formatting, images and other parts stay intact
```

## Streaming large files

For files too big to hold in memory, the **read-only** streaming entry point
walks rows one at a time (it never patches/writes):

```ts
import { StreamingWorkbook } from 'sheet-lite/stream';

const wb = await StreamingWorkbook.open('./huge.xlsx'); // Node file path
// ...or an in-memory buffer: StreamingWorkbook.open(bytes)
for (const sheet of wb.sheets) console.log(sheet.name, sheet.sheetId);

for await (const row of wb.getSheet('Data')!.readRows()) {
  // row.number; row.cells = [{ ref, row, col, value, type }]
  for (const cell of row.cells) console.log(cell.ref, cell.value, cell.type);
}
await wb.close();
```

`Workbook.openStream(path)` (main entry) is a lazy alias that loads the
streaming chunk on demand. Values and types are **identical to the eager
reader**; the difference is memory: a 100 000 × 8 sheet streams in **~10 MB**
of live heap versus **~673 MB** for `Workbook.read` (run `pnpm bench:stream`).

## Benchmark

Generated a 5000 × 8 workbook with ExcelJS and read it back:

| Library | Read (median) | Published size (gzip, incl. deps) | Transitive deps |
|---|---|---|---|
| **sheet-lite** | **149 ms** | **~30 kB** (18.5 kB main entry + fflate 11.8 kB; streaming adds ~6 kB on demand) | **0** |
| ExcelJS | 192 ms | 249.9 kB | 9 |
| SheetJS (`xlsx`) | 218 ms | 136.4 kB | 7 |

Run it yourself: `pnpm bench`

## Limitations (honest list)

`sheet-lite` is narrower than ExcelJS/SheetJS by design:

- `addWorksheet` creates **plain** worksheets only (values, optional frozen panes) — no charts, tables, images or conditional formatting on new sheets.
- `removeWorksheet` rejects removing the **last remaining sheet** (a workbook must keep at least one sheet) with the coded error `LAST_SHEET`. Orphan parts that only the removed sheet referenced (e.g. a drawing) stay in the package unreferenced.
- Only the first `sheetView` of a sheet is editable (`views`); other views/attributes are preserved as-is.
- Styles subset: number formats, bold/italic/size/name/color, solid fills, borders and cell alignment. Border colors are ARGB-only (theme colors preserved but not exposed); rich-text run fonts model bold/italic/underline/size/color/name. Charts, pivot tables, images, comments and conditional formatting are **preserved, not created**.
- No `.xls` / `.xlsb`, no encryption, no CSV.
- No formula evaluation (cached results only — same as ExcelJS and SheetJS).
- **Streaming mode is read-only.** `sheet-lite/stream` (and `Workbook.openStream`) reads rows with bounded memory but cannot patch or write; editing needs the eager `Workbook`. Streaming reads `sharedStrings.xml` into memory (the string table, not a DOM), so sheets with millions of *unique* strings are bounded by that table, not by the row count. `ReadableStream` input is not supported (ZIP has no seeking and the EOCD is at the end) — pass a Node file path or a buffer.
- `Workbook.read` is eager: the whole archive is inflated and the worksheet is parsed into memory. Use streaming for large files.

## Acknowledgements & legal

- The implementation is **original**; the only runtime dependency is [`fflate`](https://github.com/101arrowz/fflate) (MIT).
- [`exceljs`](https://github.com/exceljs/exceljs) and [`xlsx`](https://github.com/SheetJS/sheetjs) are **devDependencies used only to generate test fixtures and cross-check results** — they are never shipped in the published package.
- Published-size figures come from [bundlephobia](https://bundlephobia.com) for the exact versions listed in `devDependencies`; read times are measured by [`scripts/bench.mjs`](scripts/bench.mjs) — reproduce with `pnpm bench`.
- "Excel" is a trademark of Microsoft Corporation. ExcelJS and SheetJS are independent open-source projects; references to them in these docs are factual comparisons, not endorsements.

## Development

```bash
pnpm install
pnpm test        # vitest
pnpm typecheck   # tsc --noEmit
pnpm build       # tsdown -> dist/
pnpm bench       # size + read/write benchmark
pnpm bench:stream # eager vs streaming memory/time benchmark
pnpm size        # bundle size gate (fails over budget)
```

CI runs on every push/PR (typecheck, tests, build, size gate, pack). Releases to npm are manual (`release.yml`, dry-run first) and docs + playground deploy to **Cloudflare Pages** — see [`docs/PUBLISHING.md`](docs/PUBLISHING.md).

## License

MIT
