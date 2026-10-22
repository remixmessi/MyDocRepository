import { defineConfig } from 'vitest/config';
export default defineConfig({
  test: {
    include: ['tests/**/*.test.ts'],
    // Integration tests launch isolated Office parsers and make multiple HTTP requests.
    testTimeout: 30000,
    // Avoid starving parser subprocesses on the shared desktop CI runners.
    ...(process.env.CI ? {maxWorkers: 2} : {}),
  },
});
