import { defineConfig } from 'vite';
import { fileURLToPath } from 'node:url';

// The playground imports sheet-live live from source: `import { Workbook } from '../src/index'`.
// Those files live outside this package's root, so the dev server must be
// explicitly allowed to serve (and HMR) them.
const playgroundRoot = fileURLToPath(new URL('.', import.meta.url));
const repoRoot = fileURLToPath(new URL('..', import.meta.url));

export default defineConfig({
  server: {
    fs: {
      allow: [playgroundRoot, repoRoot],
    },
  },
});
