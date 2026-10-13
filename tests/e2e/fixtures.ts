import { test as base } from '@playwright/test';

const runId = crypto.randomUUID();

export * from '@playwright/test';

export const test = base.extend<{ projectScope: string; projectIsolation: void }>({
  projectScope: async ({}, use, info) => {
    await use(runId + '-' + info.testId + '-' + info.retry);
  },
  projectIsolation: [
    async ({ context, projectScope }, use) => {
      await context.setExtraHTTPHeaders({ 'X-Mydoc-Test-Scope': projectScope });
      await use();
    },
    { auto: true },
  ],
});
