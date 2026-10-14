# sheet-lite

> Read, write and **edit** Excel (`.xlsx`) without breaking your file — lightweight, secure and with **zero transitive dependencies**.

When you open an Excel file, change a value and save, most libraries rebuild the whole package and can silently drop formatting, data validations, conditional formatting or images. `sheet-lite` is built to **patch only what changed** and leave everything else untouched.

See [`docs/STATUS.md`](docs/STATUS.md) for the full background, decisions and roadmap.

## Status

**v0.2 — read + edit.** You can read values and patch cells/rows while the rest of the file stays intact.

- [x] `Workbook.read()` / `readFile()` — parse a package (hardened ZIP limits)
- [x] Value reading: strings, numbers, booleans, dates (1900/1904), formulas (cached result), inline & shared strings
- [x] `Worksheet` / `Cell` / `Row` API: `getWorksheet`, `getCell('A1')`, `getCell(row, col)`, `eachRow`, `addRow`
- [x] **Editing without breaking:** only changed cells are rewritten; other parts stay byte-identical
- [x] Patch-only write: if nothing changed, the original bytes are returned verbatim
- [x] Security by design: path-traversal, entry-count, size and compression-ratio limits
- [ ] Styles API, streaming, browser/edge adapters (roadmap)

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
ws.getCell('B3').value = new Date();
ws.getCell('B4').value = { formula: 'SUM(B2:B3)', result: 1234.56 };
ws.addRow(['nuevo', 7]);

await wb.writeFile('./report-edited.xlsx');
// data validations, conditional formatting, merges and other parts stay intact
```

## Benchmark

Generated a 5000 × 8 workbook with ExcelJS and read it back:

| Library | Read (median) | Published size (gzip, incl. deps) | Transitive deps |
|---|---|---|---|
| **sheet-lite** | **120 ms** | **~20.3 kB** (8.5 kB + fflate 11.8 kB) | **0** |
| ExcelJS | 192 ms | 249.9 kB | 9 |
| SheetJS (`xlsx`) | 218 ms | 136.4 kB | 7 |

Run it yourself:

```bash
pnpm bench
```

## Development

```bash
pnpm install
pnpm test        # vitest
pnpm typecheck   # tsc --noEmit
pnpm build       # tsdown -> dist/
pnpm bench       # size + read benchmark
```

## License

MIT
