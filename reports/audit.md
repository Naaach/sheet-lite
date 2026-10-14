# sheet-lite — independent audit

Auditor: subagent (read-only). Platform: Windows, PowerShell, Node v24.14.1, pnpm v12.11.2.

> **Active development notice.** The working tree moved while this audit was running.
> At the start it was clean at **`a9b91fe`** (v0.2). During the audit a concurrent process
> committed **`f21c688`** ("v0.3 - styles (numFmt/font/fill) + merges; npm publish setup; CI + docs site").
> All findings below were **re-verified against the current HEAD `f21c688`** unless explicitly marked
> otherwise. Line numbers refer to `f21c688`. This is the state I observed; it may change again.

## (a) Repository state

Start of audit:

```
$ git rev-parse HEAD
a9b91fecba0c3dd7bbf83afc67f3b819e5b1a183
$ git status
On branch main
Your branch is up to date with 'origin/main'.
nothing to commit, working tree clean
```

End of audit (another process committed mid-run):

```
$ git rev-parse HEAD
f21c688c116af46eef880e37ec074382e5777a2d
$ git log --oneline -3
f21c688 feat: v0.3 - styles (numFmt/font/fill) + merges; npm publish setup; CI + docs site
a9b91fe feat: v0.2 - value reading, Worksheet/Cell API, lossless editing + benchmark
f32d648 chore: sheet-lite v0.1 — read + lossless patch-only write
```

The concurrent commit also swept up four throw-away repro scripts I had written under `reports/`; I
deleted them afterwards, so only this file remains. I did not author or alter any `src/`, `test/` or
`package.json` content.

## (b) Command results

Run against `f21c688` (fresh `pnpm install` first; earlier `a9b91fe` runs are noted where interesting):

| Command | Result |
|---|---|
| `pnpm install` | exit 0 — "Already up to date" |
| `pnpm typecheck` (`tsc --noEmit`) | exit 0, clean |
| `pnpm test` (`vitest run`) | exit 0 — **7 files, 29/29 tests passed** (on `a9b91fe` it was 6 files / 25 passing) |
| `pnpm build` (`tsdown`) | exit 0 — ESM `dist/index.mjs` **40.30 kB / gzip 11.90 kB**; CJS 40.97 kB / gzip 12.05 kB; dts 11.62 kB. (on `a9b91fe`: ESM 27.41 kB / gzip 8.73 kB) |
| `pnpm bench` | exit 0 — see below |

`pnpm bench` output (`f21c688`):

```
Generated file: 228.3 kB · 5000 rows x 8 cols
Read time (median of 5):
  sheet-lite   147.4 ms
  SheetJS      245.0 ms
  ExcelJS      206.0 ms
Published size (gzipped, whole package incl. deps):
  sheet-lite   11.6 kB  +  fflate 11.8 kB  =  ~23.4 kB   (0 transitive deps)
  ExcelJS      249.9 kB   (9 deps)
  SheetJS      136.4 kB   (7 deps)
```

The benchmark script runs and the numbers are plausible: sheet-lite is faster than both baselines and
the reported sizes match the freshly built artifact. `node_modules/fflate/package.json` (0.8.3) has no
runtime `dependencies`, so the "0 transitive deps" claim holds.

## (c) Findings

Severities: **critical** = silently corrupts files / core promise broken; **major** = wrong data or
crash on realistic input; **minor** = edge/robustness/quality.

### CRITICAL

**C1. Adding a merge to a sheet that has none writes `<mergeCells>` *inside* `<sheetData>` (invalid OOXML).**
`src/writer/sheet-writer.ts:147-150`. `applyMergeEdits` inserts the new container at `sheet.sheetDataClose`,
which is the index of the opening `<` of `</sheetData>` — i.e. it lands *before* the closing tag, inside
`sheetData`. In ECMA-376 `mergeCells` is a sibling of `sheetData` and must come after it. Repro:

```xml
<sheetData><row r="1"><c r="A1"><v>1</v></c></row><mergeCells count="1"><mergeCell ref="C1:D1"/></mergeCells></sheetData><pageMargins .../>
```

The existing test only asserts `.toContain('<mergeCells ...')` and that SheetJS does not throw
(`test/styles.test.ts:48`), so it passes despite the invalid document. This directly contradicts the
headline promise "edit … without breaking your file". Test idea: assert `indexOf('</sheetData>') <
indexOf('<mergeCells')` and validate against the schema; merging is a very common operation on files
that never had merges.

### MAJOR

**M1. 1900 date system is off by one for serials before 1900-03-01 (the classic Excel leap-year bug).**
`src/model/dates.ts:4,8-16`. `EPOCH_1900 = 25569` assumes the modern mapping and ignores the fictitious
1900-02-29. Repro (`serialToDate`): serial `1` → `1899-12-31` (Excel: `1900-01-01`); serial `59` →
`1900-02-27` (Excel: `1900-02-28`); serial `60` → `1900-02-28` (Excel: `1900-02-29`); `dateToSerial`
(`1900-01-01`) → `2` (expected `1`). Only serials ≥ 61 are correct, which is why the compat test (a 2026
date) passes. Test idea: table of serials {1,59,60,61,25569} and the inverse for the same dates.

**M2. Date detection misses locale-prefixed date formats → dates read back as numbers.**
`src/model/dates.ts:24-27`: `isDateFormatCode` strips quoted literals and escapes, then requires
`/[ymdhs]/i` **and** `!/^\[/`. The `!^[` guard rejects the extremely common Excel locale prefix, e.g.
`[$-409]d-mmm-yy;@` → `false`, `[$-F800]dddd, mmmm dd, yyyy` → `false`. A cell with such a custom
format is returned as a raw number. Repro: end-to-end read returned `45000`/`Number` where a `Date` is
expected. Test idea: `expect(isDateFormatCode('[$-409]d-mmm-yy')).toBe(true)`.

**M3. Date detection false positive: formats containing `[Red]` are treated as dates → numbers read back as `Date`.**
`src/model/dates.ts:24-27`. The `d` in the literal `[Red]` matches `/[ymdhs]/i`, and the format does not
start with `[`, so `0.00;[Red]-0.00`, `$#,##0.00;[Red]($#,##0.00)` and `#,##0.00_);[Red](#,##0.00)` all
return `true`. End-to-end: a numeric cell (3.14) styled with a custom `0.00;[Red]-0.00` format was read
back as `Date` `1900-01-02T03:21:36.000Z`. Test idea: `expect(isDateFormatCode('0.00;[Red]-0.00')).toBe(false)`.

**M4. Cells without an `r` attribute are silently dropped, and editing such a sheet creates duplicate refs.**
`src/reader/sheet-parser.ts:119-120` (`if (!ref) continue;`). Per OOXML, `r` is optional and omitted cells
are positional (next column after the previous cell). Repro: a row `<c><v>5</v></c><c r="B1"><v>9</v></c>`
reads `A1 = null`, and after `getCell('A1').value = 111` the output keeps the positional `<c>` **and**
appends `<c r="A1">` after `<c r="B1">`, producing two cells for A1 and out-of-order columns. Test idea:
parse positional cells and assert `A1 === 5`.

**M5. Inserting a cell into an existing self-closing row writes the `<c>` outside the `<row>`.**
`src/writer/sheet-writer.ts:88-90`. For `<row r="1"/>`, `rowNode.closeStart === end`, so the inserted cell
is emitted after `/>`: `<row r="1"/><c r="A1"><v>42</v></c>`, a sibling of the row. Re-reading gives
`A1 = null` (value lost). Test idea: fixture with `<row r="1"/>` and set a cell in row 1.

**M6. XML entity decoder throws `RangeError` on out-of-range numeric entities → uncaught crash during read (DoS).**
`src/xml/xml.ts:154-165`: `String.fromCodePoint(parseInt(...))` throws for values > 0x10FFFF. Repro:
`parseXml('<a b="&#x110000;"/>')` throws `RangeError: Invalid code point 1114112`; likewise a shared
string `<t>&#x110000;</t>` decoded via `textOf` throws while `Workbook.read` is parsing. A malformed or
hostile `.xlsx` therefore aborts the read with a non-`SheetLiteError`. Also `&#0;` yields `\u0000` and
`&#xD800;` yields a lone surrogate. Test idea: `expect(() => Workbook.read(badEntityXlsx)).not.toThrow()`
(or a controlled error).

**M7. ZIP "bomb"/size limits are enforced only *after* `unzipSync` has fully decompressed the archive.**
`src/zip/archive.ts:74` calls `unzipSync(raw)` with no options; entry-count (`:80`), per-entry size
(`:98`), total size (`:104`) and ratio (`:115`) checks all run on the already-decompressed map. The limits
therefore cannot prevent the memory blow-up they are advertised to stop (README: "Security by design:
path-traversal, entry-count, size and compression-ratio limits"; STATUS: "hardened ZIP layer"). fflate
supports an `UnzipOptions.filter` callback that can reject entries before inflating; it is not used.
Evidence is by code reading — I did not attempt to OOM the process. Finding stands regardless of the
concurrent refactor.

### MINOR

- **m1. Empty `<v>` handling.** `src/reader/sheet-parser.ts:81-85`: `<v></v>` → `Number('') === 0`, so an
  empty numeric cell reads as `0`; and `t="s"` with an empty `<v>` returns `sharedStrings[0]` (`:70-71`).
- **m2. Formula without a cached result gains `result: 0` after a write/read cycle.**
  `src/writer/sheet-writer.ts:40` always emits `<v></v>`; the parser then yields `result: 0`. Repro:
  write `{ formula: '2+2' }`, re-read → `{ formula: '2+2', result: 0 }` (initially `undefined`).
- **m3. New cells appended to an existing row are not sorted against existing cells.**
  `src/writer/sheet-writer.ts:88,98`: setting `B1` on a row that has `A1` and `C1` yields `A1,C1,B1`.
- **m4. Adding a row when `<sheetData/>` is self-closing puts `<row>` after trailing elements.**
  `src/writer/sheet-writer.ts:110` + `src/reader/sheet-parser.ts:147-152`: falls back to
  `lastIndexOf('</worksheet>')`, producing `…<pageMargins/>…<row r="1">…</row></worksheet>` (row after
  pageMargins — wrong child order).
- **m5. Self-closing container tags break insertion.** `<mergeCells/>` produces
  `<mergeCells count="1"/><mergeCell ref="C1:D1"/>` (mergeCell outside the element)
  (`src/writer/sheet-writer.ts:135,145`). The same `closeStart === openEnd` pattern would affect a
  self-closing `<numFmts/>`, `<fonts/>`, `<fills/>`, `<cellXfs/>` in `style-registry.ts:203-224`.
- **m6. Merge count is not updated when a merge is removed.** `src/writer/sheet-writer.ts:135` returns
  before the count update when there are only removals, so removing one of two merges leaves
  `<mergeCells count="2">` with a single `<mergeCell>`.
- **m7. `StyleRegistry` crashes on a styles.xml that omits optional sections.** `src/writer/style-registry.ts:86,89,92`
  use non-null assertions on `fonts`/`fills`/`cellXfs`. A schema-minimal styles.xml (sections optional)
  makes `Workbook.read` throw `TypeError: Cannot read properties of undefined (reading 'children')`
  instead of a `SheetLiteError`.
- **m8. XML attribute values containing `>` are mis-parsed (legal XML).** `src/xml/xml.ts:25` uses
  `<[^>]*>`, so `parseXml('<a b="x>y"><c/></a>')` yields `<a>` with no attrs and stray text `y">`.
  Newlines inside tags are fine; `>` is the only break.
- **m9. `refToCoords` silently maps malformed refs to `{row:0,col:0}`** (`src/reader/sheet-parser.ts:52-56`);
  e.g. `getCell('A1:B2')` becomes row/col 0 and a subsequent write would emit `<c r="A1:B2">` / `<row r="0">`.
- **m10. CDATA content is entity-decoded.** `src/xml/xml.ts:71-72` appends CDATA to `node.text`, which
  `textOf` then runs through `decodeEntities` — XML says CDATA is literal. Repro: shared string
  `<t><![CDATA[a &amp; b]]></t>` reads as `"a & b"` instead of `"a &amp; b"`.
- **m11. `[Content_Types].xml` parser is order/entity-sensitive.** `src/opc/content-types.ts:19-20`
  requires `Extension` before `ContentType` (and `PartName` before `ContentType`) and does not
  entity-decode both attributes or accept closed-with-`/>` variations beyond the regex.

### Honesty check (claims vs. code)

- **H1 (with C1): "edit … without breaking your file" / "opens without repair"** (`README.md:3,5`,
  `STATUS.md:46`) is not met today: C1 produces schema-invalid sheet XML, and M5/m4 produce invalid
  element placement. SheetJS tolerates all three, which is why the tests pass.
- **H2: benchmark/verification numbers are stale versus the shipped v0.3 code.** `README.md:58` and
  `STATUS.md:103` still say **120 ms / 8.5 kB / ~20.3 kB**, and `STATUS.md:92` still says "ESM 27.4 kB /
  gzip 8.7 kB". The measured current build is **147 ms / 11.6 kB / ~23.4 kB** and ESM **40.30 kB /
  gzip 11.90 kB**. (The bench *script* measures `dist` live, so running it reveals the discrepancy; the
  hardcoded table in the docs does not.)
- **H3: "other parts stay byte-identical"** is accurate at *part-content* level (tests verify it), but the
  whole ZIP is re-deflated on every write; `STATUS.md:118` acknowledges reusing raw deflate streams is
  still roadmap. Not a defect, just worth stating precisely.
- `version` is now correctly `0.3.0`; the README/STATUS v0.3 status matches the code. The v0.2 "25/25"
  claim was updated to "29/29", matching the run.

### Fixed since `a9b91fe`

- The v0.2 `toBuffer()` bug where a low-level edit of **only** `[Content_Types].xml` was silently
  discarded (the old `hasLowLevelEdits`/`return this.store.raw` branch) is gone in `f21c688`; the edit now
  persists (verified). Had it still existed it would have been a major silent-data-loss finding.

## (d) What I could NOT verify

- **Real Excel / LibreOffice interoperability.** Validity judgments are based on ECMA-376 child ordering
  and on SheetJS being tolerant; I could not open the produced files in Excel to see whether it prompts
  to repair C1/M5/m4 outputs.
- **Actual memory behaviour of the ZIP limits (M7).** I deliberately did not feed a real zip bomb; M7 is a
  code-reading conclusion (no `UnzipOptions.filter`, all checks after `unzipSync`).
- **Broad real-world file corpus.** Styles/merges were checked with hand-built minimal parts and the
  ExcelJS fixtures in `test/`; unusual producers (prefixed namespaces, positional cells, split styles,
  many shared-string encodings) were only spot-checked.
- **Browser/edge/Deno runtimes** (STATUS targets them) — only Node 24 on Windows was exercised.
- **1904 date system** correctness beyond spot values (serial 0 → 1904-01-01, inverse → 0, serial 60 →
  1904-03-01) — those looked correct, but I did not test the full pre-1904 history.
- **Concurrency.** Another process was committing during the audit; if it keeps editing, the specific
  findings and line numbers may drift. Everything above reflects `f21c688` at the time of writing.
- **The `xlsx` (SheetJS) devDependency** is only used by tests/bench; I did not audit it.
