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

/** Confirms an open 新建文件 dialog with its defaults (type 笔记, name 未命名). */
export async function confirmNewFile(page: import('@playwright/test').Page) {
  const dialog = page.getByRole('dialog', { name: '新建文件', exact: true });
  await dialog.getByRole('button', { name: '创建', exact: true }).click();
  await base.expect(dialog).toBeHidden();
}

/** 新建文件 opens a type dialog; confirm it with the default type and name, as a user creating a note would. */
export async function newNote(page: import('@playwright/test').Page) {
  await page.getByRole('button', { name: '新建文件', exact: true }).first().click();
  await confirmNewFile(page);
}
