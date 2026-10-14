# sheet-lite

> Read, write and **edit** Excel (`.xlsx`) without breaking your file — lightweight, secure and with **zero transitive dependencies**.

When you open an Excel file, change a value and save, most libraries rebuild the whole package and can silently drop formatting, data validations, conditional formatting or images. `sheet-lite` is built to **patch only what changed** and leave everything else untouched.

📖 **Documentation site:** [`docs-site/`](docs-site/index.html) · background & decisions: [`docs/STATUS.md`](docs/STATUS.md)

## Status

**v0.3 — read, edit and style.**

- [x] `Workbook.read()` / `readFile()` / `toBuffer()` / `writeFile()`
- [x] Value reading: strings, numbers, booleans, dates (1900/1904), formulas (cached result), inline & shared strings
- [x] `Worksheet` / `Cell` / `Row` API: `getWorksheet`, `getCell`, `getRow`, `eachRow`, `addRow`
- [x] **Editing without breaking:** only changed cells/merges are rewritten; other parts stay byte-identical
- [x] **Styles:** `cell.numFmt`, `cell.font` (bold/italic/size/name/color), `cell.fill` (solid/pattern)
- [x] **Merges:** `worksheet.mergeCells('A1:C1')` / `unmergeCells`
- [x] Patch-only write: if nothing changed, the original bytes are returned verbatim
- [x] Security by design: path-traversal, entry-count, size and compression-ratio limits
- [ ] Streaming + browser/edge adapters (roadmap)

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

await wb.writeFile('./report-edited.xlsx');
// data validations, conditional formatting, images and other parts stay intact
```

## Benchmark

Generated a 5000 × 8 workbook with ExcelJS and read it back:

| Library | Read (median) | Published size (gzip, incl. deps) | Transitive deps |
|---|---|---|---|
| **sheet-lite** | **120 ms** | **~20.3 kB** (8.5 kB + fflate 11.8 kB) | **0** |
| ExcelJS | 192 ms | 249.9 kB | 9 |
| SheetJS (`xlsx`) | 218 ms | 136.4 kB | 7 |

Run it yourself: `pnpm bench`

## Development

```bash
pnpm install
pnpm test        # vitest
pnpm typecheck   # tsc --noEmit
pnpm build       # tsdown -> dist/
pnpm bench       # size + read benchmark
```

CI runs on every push/PR (typecheck, tests, build). The docs site is deployed to GitHub Pages.

## License

MIT
