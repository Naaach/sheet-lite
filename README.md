# sheet-lite

> Read, write and **edit** Excel (`.xlsx`) without breaking your file — lightweight, secure and with **zero transitive dependencies**.

`sheet-lite` is being built around one promise: **when you open an Excel file and change a value, everything else stays intact.** Most libraries re-serialize the whole package and silently drop formatting, data validations, charts and images. `sheet-lite` is designed to *patch* only what changed.

See [`docs/STATUS.md`](docs/STATUS.md) for the full background, locked decisions and roadmap.

## Status

**v0.1 — foundation.** Reading a package and writing it back losslessly (patch-only) already works and is tested.

- [x] `Workbook.read()` / `Workbook.readFile()` (Node) — parse an OPC package with hardened ZIP limits
- [x] `Workbook.toBuffer()` / `writeFile()` — **patch-only** write: if nothing changed, the original bytes are returned verbatim
- [x] `PartStore` — in-memory parts + dirty tracking
- [x] Minimal OPC helpers: `[Content_Types].xml` parsing, relationship path resolution
- [x] Security by design: path-traversal, entry-count, size and compression-ratio limits
- [ ] Cell values / styles (v0.2)
- [ ] Editing cells & rows (v0.2)
- [ ] Browser / edge adapters (v0.4)

## Design

- **Patch-only + copy-verbatim.** Only re-serialize the parts that actually changed; everything else is carried over untouched.
- **Zero transitive dependencies.** The only direct dependency is `fflate` (~8 KB, no transitive deps), used for the ZIP/DEFLATE layer.
- **Runtime-agnostic core.** The core works on `Uint8Array` / `ArrayBuffer`. Filesystem helpers (`readFile`/`writeFile`) load `node:fs/promises` lazily, so the browser bundle stays clean.
- **Security first.** ZIP bombs, path traversal and absurd ratios are rejected before any parsing.

## Usage

```ts
import { Workbook } from 'sheet-lite';

const wb = await Workbook.readFile('./report.xlsx');

// v0.1: low-level part access (a friendly Worksheet/Cell API lands in v0.2)
const parts = wb.parts();
const sheet = parts.get('xl/worksheets/sheet1.xml')!;
parts.set('xl/worksheets/sheet1.xml', new TextEncoder().encode(
  new TextDecoder().decode(sheet).replace('Hello', 'Hola'),
));

await wb.writeFile('./report-edited.xlsx');
// every other part is byte-for-byte identical
```

## Development

```bash
pnpm install
pnpm test        # vitest
pnpm typecheck   # tsc --noEmit
pnpm build       # tsdown -> dist/
```

## License

MIT
