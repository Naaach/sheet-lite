# sheet-lite — Project status & background

> Lightweight and secure library to read, write and **edit** Excel (`.xlsx`) without breaking your file.
> Zero transitive dependencies.

This document consolidates the *why*, the *decisions* and the *current state* of the project.

---

## 1. Why this project exists (market evidence)

The JS spreadsheet ecosystem has an unusual gap: **the demand is huge and the maintenance is stalled.**

| Package | Downloads / month | GitHub stars | Open issues | Last push |
|---|---|---|---|---|
| `exceljs` | **74.9M** | 15.5k | **816** | **Jan 2025** (~21 months) |
| `xlsx` (SheetJS) | 57.9M | 36.3k | 133 | Apr 2024 (npm frozen at 0.18.5 since 2022) |
| `pdf-lib` (for comparison) | 71.7M | 8.7k | 318 | Jul 2024 |

The most-downloaded library in the domain (`exceljs`) has been effectively frozen for ~21 months and, critically, **destroys data when you read + write an existing file**:

- 🚫 **Data validations**: lost (7 → 0 rules).
- 🚫 **Conditional formatting**: cleared ([exceljs #1024](https://github.com/exceljs/exceljs/issues/1024)).
- 🚫 **Drawings / images**: lost.
- 🚫 **Styles**: mostly stripped (10 KB → 1 KB).

The other option (`xlsx`/SheetJS) is insecure and abandoned on npm (high-severity advisories with *"no fix available"*), and **107+ public issues** on GitHub are literally titled `"replace xlsx"`.

And people are asking for this in the open:

- r/node: *"Need a JavaScript Library to Modify Excel Files **Without Losing Formatting**"*.
- xlsx-populate's own README: *"it is extremely difficult for libraries to support the entire specification … any content/styles they don't support are lost."*

**Conclusion:** the winning angle is not "another values reader" — it is **editing without breaking the file.**

---

## 2. Decisions (locked)

| # | Decision | Value |
|---|---|---|
| 1 | Package name | **`sheet-lite`** |
| 2 | Runtime target | **Node + browser from day one** (edge later) |
| 3 | Dependencies | **`fflate` only** (ZIP + DEFLATE, ~8 KB, **0 transitive deps**) |
| 4 | Tooling | **TypeScript + `tsdown` + `pnpm` + `vitest`** |
| 5 | Public promise | **"opens without repair"** (byte-identical for untouched parts is an internal quality bar, not a marketing claim) |
| 6 | License | **MIT** |

**On dependencies:** reading any real `.xlsx` requires DEFLATE. DEFLATE is now native (Node `zlib`, browser `DecompressionStream('deflate-raw')`), but the ZIP container and CRC32 are not. `fflate` solves all three in ~8 KB with no transitive deps. The ZIP layer is isolated in `src/zip/`, so switching to native APIs later (to claim literal "0 dependencies") is a small refactor.

---

## 3. Architecture

```
API (Workbook)                         <- public surface
  |
Model (only what is touched)           <- future: Worksheet / Cell
  |
Part engine (patch-only)               <- PartStore + dirty tracking
  |
OPC layer                              <- ContentTypes, Relationships
  |
ZIP layer (OPC packaging)              <- read/write, hardened
  |
Runtime: Node (fs) + Browser (Blob)    <- TS, ESM + CJS
```

**Core principle — patch-only + copy-verbatim:** only re-serialize parts that actually changed. Everything else is carried over untouched. When nothing changed, the original bytes are returned verbatim.

---

## 4. Current state — v0.3 (read + edit + style)

- [x] `Workbook.read()` / `readFile()` / `toBuffer()` / `writeFile()`.
- [x] **Patch-only** write (no-op returns the original bytes exactly).
- [x] `PartStore` with dirty tracking; hardened ZIP layer.
- [x] OPC helpers: `[Content_Types].xml` parsing, relationship resolution.
- [x] Value reading: number, string, boolean, Date (1900/1904), formula (cached result), inline & shared strings.
- [x] Public `Worksheet` / `Cell` / `Row` API (`getWorksheet`, `getCell`, `getRow`, `eachRow`, `addRow`).
- [x] Editing: surgical patch of only the changed cells/merges; every untouched byte preserved.
- [x] Styles: `cell.numFmt`, `cell.font` (bold/italic/size/name/color), `cell.fill` (solid/pattern) via a deduplicating `StyleRegistry`.
- [x] Merges: `worksheet.mergeCells` / `unmergeCells`.
- [x] TypeScript-first, ESM + CJS, runtime-agnostic core (filesystem via lazy `node:fs/promises`).

### Verification (executed)

| Check | Result |
|---|---|
| `pnpm test` | **29/29 passing** (incl. real ExcelJS-generated files) |
| `pnpm typecheck` (TS 7) | clean |
| `pnpm build` (`tsdown`) | ESM 27.4 kB / **gzip 8.7 kB** (excl. `fflate`) |
| Values vs SheetJS | match on a real file (numbers, strings, booleans, formulas) |
| Preservation | data validations, conditional formatting and merges survive an edit |
| Styles & merges | round-trip through our own reader; the result parses with SheetJS |

Tests cover: reading every part, exact no-op round-trip, byte-identical preservation of unrelated parts, value types, dates 1900/1904, formula round-trip, compatibility with a real ExcelJS file (cross-checked against SheetJS), ZIP hardening (traversal, absolute paths, null bytes, entry-count, size and ratio limits).

### Benchmark (5000 × 8 worksheet generated with ExcelJS)

| Library | Read (median) | Published size (gzip, incl. deps) | Transitive deps |
|---|---|---|---|
| **sheet-lite** | **120 ms** | **~20.3 kB** (8.5 kB + fflate 11.8 kB) | **0** |
| ExcelJS | 192 ms | 249.9 kB | 9 |
| SheetJS (`xlsx`) | 218 ms | 136.4 kB | 7 |

> **Honesty note:** in the specific scenario we tested, **ExcelJS 4.4.0 preserved** the conditional formatting / data validation, so we do not assert the opposite in our own tests. The preservation guarantee is verified on *our* side (real files + byte-identical untouched parts).

---

## 5. Roadmap

| Version | Goal | "Done" criteria |
|---|---|---|
| **v0.1** ✅ | Read + lossless patch-only write | Round-trip tests green |
| **v0.2** ✅ | Values API + editing (preserving) + real-file compatibility + benchmark | 25/25 tests; values match SheetJS; benchmark published |
| **v0.3** ✅ | Styles API (`numFmt`/font/fill) + merges | 29/29 tests; output parses with SheetJS |
| **v0.4** | Browser / edge (Deno, Workers) adapters; reuse original deflate streams | Example running on Workers + Deno; faster writes |
| **v0.5** | Streaming reads for large files | 100k rows under a fixed memory ceiling |

**Explicitly out of scope (for now):** `.xls`/`.xlsb`, encryption, formula evaluation, creating charts/pivots/images/macros (preserved, not created).

---

## 6. Development

```bash
pnpm install
pnpm test        # vitest
pnpm typecheck   # tsc --noEmit
pnpm build       # tsdown -> dist/
```

### Layout

```
src/
  index.ts                # public API
  api/workbook.ts         # Workbook: read / toBuffer / worksheets
  api/worksheet.ts        # Worksheet: getCell / eachRow / addRow
  api/cell.ts             # Cell
  xml/xml.ts              # tiny XML reader + entity helpers
  reader/package-reader.ts# workbook, rels, shared strings, styles
  reader/sheet-parser.ts  # cells, rows, references, value decoding
  writer/sheet-writer.ts  # surgical cell patching
  model/value.ts          # CellValue / ValueType
  model/dates.ts          # Excel serial <-> Date (1900/1904)
  zip/archive.ts          # ZIP read/write over fflate + safety limits
  zip/errors.ts           # SheetLiteError / ZipSecurityError
  opc/part-store.ts       # in-memory parts + dirty tracking
  opc/content-types.ts    # [Content_Types].xml parser
  opc/relationships.ts    # relationship path resolution
test/
  fixtures/make.ts        # external .xlsx fixture
  noop-roundtrip.test.ts  # lossless round-trip
  edit.test.ts            # editing preserves everything else
  values.test.ts          # value reading + types
  compat.test.ts          # real ExcelJS file vs SheetJS
  zip-hardening.test.ts   # security limits
  opc.test.ts             # content-types + paths
scripts/bench.mjs         # size + read benchmark
```
