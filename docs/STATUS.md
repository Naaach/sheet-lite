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

### Runtime milestone — raw ZIP reuse + browser/edge adapters

- [x] **Raw ZIP layer** (`src/zip/raw.ts`): hand-rolled local headers / central directory / EOCD parsing and a writer that reuses each untouched entry's ORIGINAL compressed bytes + CRC verbatim (only the edited part is re-compressed). Dependency-free container logic; CRC-32 (slicing-by-8) implemented in-house.
- [x] **Rejected with a stable error instead of crashing:** zip64 (`ZIP_UNSUPPORTED`), data descriptors (bit 3), encryption (bit 0), unknown compression methods, multi-disk archives. Note: streaming-written files (e.g. `xlsx.writeStream`) that use bit-3 descriptors are therefore rejected on read.
- [x] **All hardening kept, and strengthened:** path-traversal, entry-count, per-entry size, total size and compression-ratio limits are enforced on the ZIP structure BEFORE inflation; inflate output is preallocated to the declared (validated) size, so lying headers cannot over-allocate; every entry is CRC-32 verified.
- [x] **Browser / edge adapters on `Workbook`:** `readBlob(blob)`, `toBlob()`, `download(filename)` (browser only, throws `ENV_UNSUPPORTED` elsewhere). `node:fs` stays lazy; the ESM dist contains no static `node:` references (verified by a test).
- [x] Tests: no-op whole-file byte-identity via the raw path, per-entry verbatim reuse after an edit (independent central-directory walker), stored+deflate archives, zip64 / data-descriptor / encrypted / CRC-mismatch rejection, Blob adapters, dist scan.

### Stream milestone — bounded-memory reads

- [x] **Streaming ZIP layer** (`src/zip/stream.ts`): random-access `ByteSource` (Node file handle, or an in-memory buffer) + `readZipDirectory()` / `streamEntryChunks()`. It reuses the SAME structural parser as the eager path (`readEocd` / `parseCentralDirectory` / `resolveEntryDataOffset` in `zip/raw.ts`), so every structural hardening rule — zip64, encryption, multi-disk, traversal names, entry-count / per-entry / total-size and compression-ratio limits — rejects **identically, before any inflation**, and features `raw.ts` supports (stored + deflate, data-descriptor entries via the central-directory CRC/sizes) behave the same. Entries are inflated with fflate's push-based `Inflate` in 64 KiB compressed chunks; every streamed byte is CRC-32 verified and a stream that produces more than its declared size is aborted.
- [x] **Incremental XML reader** (`src/xml/pull-parser.ts` + `src/reader/stream-sheet.ts`): a push/pull tokenizer that decodes one `<row>` at a time with the SAME value rules as the eager reader (shared / inline / formula strings, numbers, dates 1900/1904, booleans, errors, formulas with cached result, positional cells). Only the current row is ever materialized.
- [x] **Shared strings parsed incrementally too** (`reader/package-reader.ts`): no DOM, so a large `sharedStrings.xml` no longer spikes memory (the string table itself is still held in memory while streaming — see limitations).
- [x] **Public API**: `StreamingWorkbook` / `StreamingSheet` in the separate **`sheet-lite/stream` entry point**, plus `Workbook.openStream(...)` on the main entry (a lazy `import()` of the streaming chunk). `workbook.sheets` lists name/id; `sheet.readRows()` returns an `AsyncGenerator<StreamRow>` yielding `{ number, cells: [{ ref, row, col, value, type }] }`. Inputs: Node file path or in-memory buffer. **Read-only** — streaming has no patch/write path.
- [x] **Kept out of the main bundle**: streaming ships as `dist/stream.mjs` + a lazy chunk; `scripts/size.mjs` now gates the static import graph of `index.mjs` and reports lazy chunks separately.
- [x] Tests (`test/stream.test.ts`): exact value/type parity with the eager reader (generated + ExcelJS files, all value kinds incl. dates, a formula, rich-text shared strings), every structural hardening rejection through the streaming path plus data-descriptor parity, CRC-32 rejection, bounded parser buffer, incremental decompression, and a memory gate with forced GC.

### Stream benchmark — eager vs streaming (100 000 × 8 worksheet, ExcelJS)

Measured with `pnpm bench:stream`; each phase runs in its own fresh node process (`node --expose-gc`), Node 24 / Windows. File: 4.5 MB compressed; worksheet part **30.5 MB** inflated; sharedStrings 4.8 MB inflated.

| Metric | Eager (`Workbook.read`) | Streaming (`StreamingWorkbook`) |
|---|---|---|
| Peak heap (sampled) | 673.0 MB | **19.0 MB** |
| Live heap (after GC) | 673.0 MB (retained) | **10.2 MB** |
| Read / walk time | 2.87 s | 3.21 s |
| Cells streamed | 800 000 | 800 000 |

The streaming reader holds only a 64 KiB compressed chunk, fflate's 32 KiB window, a few output chunks, the current row and the shared-string table, so its live heap stays bounded regardless of file size (**~66× smaller** here). Note: eager `Worksheet.eachRow()` is quadratic today and is excluded from the eager timing — the eager cost measured is `Workbook.read`, which is also where the memory goes.

### Verification (executed)

| Check | Result |
|---|---|
| `pnpm test` | **81/81 passing** (15 files; includes the 56 pre-existing tests and the sibling data-descriptor/corpus suites) |
| `pnpm typecheck` (TS 7) | clean |
| `pnpm build` (`tsdown`, code-split) | ESM main graph **19.5 kB gzip** (`index.mjs` 8.7 + shared chunk 10.8; budget 20 kB); streaming chunk 5.9 kB gzip + `stream.mjs` 0.2 kB, loaded on demand (not gated) |
| `pnpm size` | ok (main entry under the 20 kB gzip budget) |
| `pnpm bench:stream` | 100k × 8: eager 673 MB vs streaming 10.2 MB live heap (see table above) |
| ESM dist static-`node:` scan | clean across every `dist/*.mjs` chunk (only lazy `import("node:fs/promises")`) |
| No-op round-trip | `Workbook.read(f).toBuffer()` returns the original bytes verbatim; `writeZip(readZip(f))` is byte-identical for archives without directory markers |
| Preservation | after a single-cell edit, every untouched ZIP entry is byte-identical (local header + compressed bytes + CRC) |
| Values vs SheetJS | match on a real file (numbers, strings, booleans, formulas) |
| Streaming vs eager | identical cells/types incl. dates and a formula, on generated + ExcelJS files |

Tests cover: reading every part, exact no-op round-trip, byte-identical preservation of unrelated ZIP entries (raw verbatim reuse), value types, dates 1900/1904, formula round-trip, compatibility with a real ExcelJS file (cross-checked against SheetJS), ZIP hardening (traversal, absolute paths, null bytes, entry-count, size and ratio limits, zip64/data-descriptor/encryption rejection, CRC validation), styles, merges, Blob adapters, the audit regression cases, and streaming (value/type parity with the eager reader, the same hardening rejections, incremental decompression and a bounded-memory gate).

### Write benchmark — raw deflate reuse (5000 × 8 worksheet, same file as below)

Same inputs, interleaved A/B: the v0.3 writer (`fflate.zipSync` over ALL entries) vs the runtime writer (verbatim reuse of untouched entries).

| Write scenario | v0.3 writer | runtime writer |
|---|---|---|
| Raw pass-through (`writeZip(readZip(f))`, nothing changed) | ~99 ms | **~0.1 ms** (~900x) |
| Small part edited (e.g. metadata) | ~99 ms | **~0.2 ms** (~450x) |
| One cell edited in the big sheet | ~99 ms | ~93–103 ms* |

\* The edited sheet is ~85% of this file's decompressed bytes and must be re-compressed either way, so the win is bounded here; the reuse pays off whenever untouched parts dominate. Full write pipeline (read + edit one cell + serialize) measures ~250–290 ms, dominated by reading (~150 ms) and re-compressing the edited sheet. No-op writes through `Workbook.toBuffer()` return the original bytes without any ZIP work.

> **Honesty notes:** (1) whole-file byte-identity via the raw path holds for archives without directory markers; directory-marker entries (e.g. `_rels/` in ExcelJS output) were already dropped by the previous reader and are still omitted from re-serialized output, while a no-op `toBuffer()` returns the original file verbatim for any input. (2) Fresh entries get a fixed DOS-epoch timestamp unless an `mtime` is provided (v0.3 stamped them with the current time via fflate). (3) The numbers above are medians on one machine (Node 24, Windows); cross-run variance was ±10–15%.

### Benchmark (5000 × 8 worksheet generated with ExcelJS)

| Library | Read (median) | Published size (gzip, incl. deps) | Transitive deps |
|---|---|---|---|
| **sheet-lite** | **~150 ms** | **~31 kB** (19.5 kB main entry + fflate 11.8 kB; streaming adds ~6 kB on demand) | **0** |
| ExcelJS | ~175–230 ms | 249.9 kB | 9 |
| SheetJS (`xlsx`) | ~210–260 ms | 136.4 kB | 7 |

> An independent audit (see [`reports/audit.md`](../reports/audit.md)) was run against v0.3. Every finding was fixed and is covered by tests: invalid `mergeCells` placement (was emitted inside `<sheetData>`), 1900 leap-year off-by-one, date-format detection (locale prefixes vs accounting formats), positional cells without `r`, empty `<v>` read as `0`, out-of-range XML entities throwing, and ZIP inflation limits.

> **Honesty note:** in the specific scenario we tested, **ExcelJS 4.4.0 preserved** the conditional formatting / data validation, so we do not assert the opposite in our own tests. The preservation guarantee is verified on *our* side (real files + byte-identical untouched parts).

---

## 5. Roadmap

| Version | Goal | "Done" criteria |
|---|---|---|
| **v0.1** ✅ | Read + lossless patch-only write | Round-trip tests green |
| **v0.2** ✅ | Values API + editing (preserving) + real-file compatibility + benchmark | 25/25 tests; values match SheetJS; benchmark published |
| **v0.3** ✅ | Styles API (`numFmt`/font/fill) + merges | 29/29 tests; output parses with SheetJS |
| **runtime** ✅ | Browser / edge (Deno, Workers) adapters; reuse original deflate streams | Raw ZIP layer with verbatim reuse (byte-identical untouched entries; zip64/descriptor rejection); `readBlob`/`toBlob`/`download` + ESM dist verified free of static `node:` imports; 56/56 tests. Workers/Deno example itself not exercised (Blob API smoke-tested on Node) |
| **stream** ✅ | Streaming reads for large files | Bounded-memory read-only API (`sheet-lite/stream`): 100k × 8 rows at **10.2 MB** live heap vs **673 MB** eager, values identical to the eager reader; same hardening rejections through the streaming path; main bundle kept under budget via a lazy chunk. 81/81 tests |
| **real-world** 🟡 | Real-file corpus + LibreOffice gate in CI | Infra ready (`test/fixtures/real/`, `test/corpus.test.ts`, `scripts/validate-corpus.mjs`, `pnpm validate:corpus`); 4 real files pass read/edit/no-op. Pending: LibreOffice CI job + more files |

> Milestones are named, not versioned. Published npm versions continue sequentially: `0.3.0` = styles; **`0.4.0` = runtime + stream + real-world corpus infrastructure + ZIP data-descriptor support**. The streaming API is additive and ships from a separate `sheet-lite/stream` subpath so the eager entry stays small.

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
  stream.ts               # `sheet-lite/stream` entry point (StreamingWorkbook)
  api/workbook.ts         # Workbook: read / readBlob / toBuffer / toBlob / download / openStream(→ stream)
  api/streaming.ts        # StreamingWorkbook / StreamingSheet (read-only, row-by-row)
  api/worksheet.ts        # Worksheet: getCell / eachRow / addRow
  api/cell.ts             # Cell
  xml/xml.ts              # tiny DOM XML reader + entity helpers
  xml/pull-parser.ts      # incremental XML pull-parser (streaming sheets + shared strings)
  reader/package-reader.ts# workbook, rels, shared strings, styles
  reader/sheet-parser.ts  # cells, rows, references, value decoding
  reader/stream-sheet.ts  # streaming worksheet: one row at a time
  writer/sheet-writer.ts  # surgical cell patching
  model/value.ts          # CellValue / ValueType
  model/dates.ts          # Excel serial <-> Date (1900/1904)
  zip/raw.ts              # raw ZIP container: parse/serialize, verbatim entry reuse, CRC-32, hardening
  zip/stream.ts           # streaming ZIP: byte sources, directory parse, incremental inflate
  zip/archive.ts          # ZIP facade: readZip/writeZip on the raw layer + safety limits
  zip/errors.ts           # SheetLiteError / ZipSecurityError / ZipUnsupportedError
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
  raw-zip.test.ts         # raw-layer: verbatim reuse, byte-identity, zip64/descriptor/CRC rejection
  stream.test.ts          # streaming parity + hardening + bounded-memory gate
  browser-adapters.test.ts# readBlob / toBlob / download
  dist.test.ts            # built ESM (every chunk) has no static node: imports
  opc.test.ts             # content-types + paths
scripts/bench.mjs         # size + read + write benchmarks (incl. raw-reuse scenarios)
scripts/bench-stream.mjs  # eager-vs-streaming memory/time benchmark
```
