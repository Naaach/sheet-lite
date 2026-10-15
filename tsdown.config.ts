import { defineConfig } from 'tsdown';

export default defineConfig({
  // Three entry points, three concerns: the eager reader/editor, the streaming
  // reader (`sheet-lite/stream`) and the incremental writer (`sheet-lite/write`).
  // The last two are code-split so the main entry keeps its 20 kB gzip budget
  // (`scripts/size.mjs` gates the static import graph of `dist/index.mjs`).
  entry: ['src/index.ts', 'src/stream.ts', 'src/write.ts'],
  format: ['esm', 'cjs'],
  dts: true,
  clean: true,
  target: 'es2022',
  sourcemap: true,
  treeshake: true,
  // The 20 kB gzip budget for the main entry cannot absorb new features
  // (sheet structure editing, the incremental writer) without minification.
  minify: true,
});
