import { describe, it, expect, beforeAll } from 'vitest';
import { execSync } from 'node:child_process';
import { existsSync, readdirSync, readFileSync, statSync } from 'node:fs';
import { fileURLToPath } from 'node:url';

const root = fileURLToPath(new URL('..', import.meta.url));
const distPath = `${root}dist/index.mjs`;

/** Newest mtime of any source file, to detect a stale dist. */
function newestSrcMtime(): number {
  let newest = 0;
  const visit = (dir: string): void => {
    for (const name of readdirSync(dir, { withFileTypes: true })) {
      const full = `${dir}/${name.name}`;
      if (name.isDirectory()) visit(full);
      else newest = Math.max(newest, statSync(full).mtimeMs);
    }
  };
  visit(`${root}src`);
  return newest;
}

beforeAll(() => {
  const stale = !existsSync(distPath) || statSync(distPath).mtimeMs < newestSrcMtime();
  if (stale) {
    execSync('pnpm exec tsdown', { cwd: root, stdio: 'inherit' });
  }
}, 120_000);

describe('runtime — ESM dist is browser/edge safe', () => {
  it('contains no static `node:` imports (filesystem stays lazy)', () => {
    const content = readFileSync(distPath, 'utf8');
    expect(content).not.toMatch(/from\s*["']node:/);
    expect(content).not.toMatch(/require\(\s*["']node:/);
    // The filesystem helpers must be reachable only through dynamic import.
    expect(content).toMatch(/import\(\s*["']node:fs\/promises["']\s*\)/);
  });
});
