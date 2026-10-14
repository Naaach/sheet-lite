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

## 4. Current state — v0.1 (done)

- [x] `Workbook.read()` / `Workbook.readFile()` — parse an OPC package.
- [x] `Workbook.toBuffer()` / `writeFile()` — **patch-only** write (no-op returns the original bytes exactly).
- [x] `PartStore` — in-memory parts + dirty tracking.
- [x] OPC helpers: `[Content_Types].xml` parsing, relationship path resolution.
- [x] Security by design: path-traversal, entry-count, per-entry/total size and compression-ratio limits.
- [x] TypeScript-first, ESM + CJS, runtime-agnostic core (filesystem via lazy `node:fs/promises`).

### Verification (executed)

| Check | Result |
|---|---|
| `pnpm test` | **13/13 passing** |
| `pnpm typecheck` (TS 7) | clean |
| `pnpm build` (`tsdown`) | ESM **8.59 kB** / gzip **3.11 kB** · CJS 8.92 kB (excl. `fflate`) |

Tests cover: reading every expected part, exact no-op round-trip, byte-identical preservation of unrelated parts when one part changes, content-types parsing, relationship resolution, and ZIP hardening (traversal, absolute paths, null bytes, entry-count, size and ratio limits).

---

## 5. Roadmap

| Version | Goal | "Done" criteria |
|---|---|---|
| **v0.1** ✅ | Read + lossless patch-only write | Round-trip tests green |
| **v0.2** | Public `Worksheet`/`Cell` API; read values (types, dates 1900/1904, shared/inline strings); edit cells & rows preserving everything else; reuse original deflate streams | Reads 20 real files with values matching SheetJS; edited files open without repair |
| **v0.3** | Basic styles + merges (bold/font/fill/`numFmt`/width) | Visual render matches original; styles de-duplicated |
| **v0.4** | Browser / edge (Deno, Workers) adapters; size benchmark vs ExcelJS | Example running on Workers + Deno; published size numbers |
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
  zip/archive.ts          # ZIP read/write over fflate + safety limits
  zip/errors.ts           # SheetLiteError / ZipSecurityError
  opc/part-store.ts       # in-memory parts + dirty tracking
  opc/content-types.ts    # [Content_Types].xml parser
  opc/relationships.ts    # relationship path resolution
  api/workbook.ts         # Workbook.read / toBuffer (patch-only)
test/
  fixtures/make.ts        # external .xlsx fixture (built with fflate directly)
  noop-roundtrip.test.ts
  zip-hardening.test.ts
  opc.test.ts
```
