import { defineConfig } from 'vitest/config';

export default defineConfig({
  test: {
    include: ['test/**/*.test.ts'],
    environment: 'node',
    // Lets the memory-ceiling test collect garbage between phases for an
    // honest live-heap measurement.
    execArgv: ['--expose-gc'],
  },
});
