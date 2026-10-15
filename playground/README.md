# sheet-lite · playground

An interactive, in-browser playground for [`sheet-lite`](../README.md) — open an
`.xlsx`, edit it live, and watch the patch-only writer do its thing.

Vite + vanilla TypeScript. No frameworks. The library is imported **live from
source** (`import { Workbook } from '../src/index'`), so any edit you make in
`src/` shows up on the next reload/HMR tick without a library build step.

## Run it

```sh
pnpm --dir playground dev        # or: cd playground && pnpm install && pnpm dev
```

Then open the printed URL (default `http://localhost:5173`).

Other scripts:

```sh
pnpm --dir playground build      # production bundle into playground/dist
pnpm --dir playground preview    # serve the production build
pnpm --dir playground typecheck  # tsc --noEmit (covers playground + ../src)
```

## What it demonstrates

- **Reading** — `Workbook.read(bytes)` on a file picked or dropped in the
  browser (plus a **Generate sample** button that builds a workbook entirely
  in-browser with the library itself: values, a real date, formulas, number
  formats, bold headers and merges across two sheets).
- **The public API surface** — sheet tabs (`wb.sheetNames`), an editable grid
  backed by `getCell`/`getRow`/`eachRow`, `cell.value` edits (numbers, text,
  booleans, dates, and `=…` typed formulas), `cell.font`, `cell.numFmt`,
  `sheet.addRow(values)`, `sheet.mergeCells`/`unmergeCells`.
- **Cell styling (styles2 milestone)** — the **Cell actions** panel writes
  `cell.border` (*Box border* = four thin sides, *Bottom border* = one side
  merging over whatever the cell already had), `cell.alignment`
  (left/center/right, wrap) and a **rich-text value**
  (`cell.value = { richText: [{ text, font }, …] }`), which is painted run by
  run in the grid and in the inspector (which also lists border + alignment).
- **Sheet structure (sheets milestone)** — the **Sheet structure** panel calls
  `wb.addWorksheet(name)` / `wb.removeWorksheet(name)`, `ws.getColumn(n).width`
  / `.hidden`, `ws.getRow(n).height` and `ws.views = [{ state: 'frozen',
  xSplit, ySplit }]` (plus `state: 'normal'` to unfreeze). Column and row
  targets follow the grid selection unless you point them somewhere else.
- **Patch-only writing** — the **Patch monitor** snapshots every OPC part at
  load time (`Workbook.parts()`), then after each edit compares parts
  byte-for-byte between the original and a fresh `Workbook.read(wb.toBuffer())`:
  - *Original vs Output* file sizes with the delta,
  - a checklist of every part (`✓ identical` / `● re-serialized` / `+ added`),
    so you can see that only dirty sheets — and `xl/styles.xml` when new
    formats were introduced — are rewritten.
- **Safety** — `SheetLiteError` and `ZipSecurityError` surface in an error bar
  with their machine-readable `code` (try the *Unsafe demo* button, which
  feeds the library a package with a `../` path-traversal entry, or duplicate /
  illegal sheet names in the **Sheet structure** panel, which report
  `SHEET_EXISTS` / `BAD_SHEET_NAME`). Removing the last remaining sheet is
  reported the same way (`LAST_SHEET`) and the button is disabled up front.
- **Downloading** — `wb.toBuffer()` → `Blob` → temporary `<a download>` link.
  (Node-only APIs like `Workbook.readFile`/`writeFile` are never used here.)

## Keyboard

| Key            | Action                                    |
| -------------- | ----------------------------------------- |
| Arrows         | Move the selected cell                    |
| `Enter` / `F2` | Start editing the selected cell           |
| Any character  | Start editing with that character         |
| `Enter`        | Commit and move down                      |
| `Tab`          | Commit and move right (`Shift+Tab` left)  |
| `Esc`          | Cancel the edit                           |
| `Delete`       | Clear the selected cell                   |

## Notes & limitations

- **Not published to npm** — the root package's `files` only ships `dist/`,
  `README.md` and `LICENSE`; `playground/` never leaves the repo. It has its
  own `package.json` (devDependency: `vite`) and lockfile, and adds nothing to
  the root package.
- The grid is a preview, not a full spreadsheet engine: rows are capped at 500
  (grow with *Show 500 more rows* — the cap exists so huge files stay snappy)
  and columns at 40. Merged ranges are written correctly but are not visually
  spanned in the grid; the inspector shows the cell's merge/number-format data.
- Column widths and row heights set through the panel are *approximated* in the
  grid (Excel character units → px, points → px, both clamped); a hidden column
  collapses to a thin gutter instead of disappearing, so it stays selectable.
  Frozen panes are reported in the panel only — the grid header/row-number
  columns are always sticky.
- Cell borders can be *added* and merged side-by-side, but not cleared: the
  library merges `cell.border` over the cell's current border and an empty
  `{}` is a deliberate no-op, so the panel offers no "remove border" button.
- Cell values written as rich text read back as `{ richText }`; the grid and
  the inspector flatten the runs for display, and editing such a cell replaces
  it with the typed plain text.
- Dates typed into the grid use `YYYY-MM-DD`; other date input formats are
  stored as text.
- Formula cells show their cached result (with an `f` badge and the formula in
  the tooltip); the library does not evaluate formulas, so results refresh when
  Excel/LibreOffice recalculates the downloaded file.
- The ZIP container itself is always re-assembled on write, so the *file* bytes
  differ even when every part is byte-identical — the patch monitor measures
  part-level identity, which is what "patch-only" guarantees.
