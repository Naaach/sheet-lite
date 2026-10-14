import { defineConfig } from 'tsdown';

export default defineConfig({
  entry: ['src/index.ts', 'src/stream.ts'],
  format: ['esm', 'cjs'],
  dts: true,
  clean: true,
  target: 'es2022',
  sourcemap: true,
  treeshake: true,
  // The 20 kB gzip budget for the main entry cannot absorb new features
  // (sheet structure editing) without minification.
  minify: true,
});
