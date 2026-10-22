import {expect, type Page} from '@playwright/test';

/** Create a note through the unified file dialog before using the note editor. */
export async function createNote(page: Page) {
  await page.getByRole('button', {name: '新建文件', exact: true}).first().click();
  const dialog = page.getByRole('dialog', {name: '新建文件', exact: true});
  await dialog.getByLabel('文件类型', {exact: true}).selectOption('note');
  await dialog.getByRole('button', {name: '创建', exact: true}).click();
  await expect(dialog).toHaveCount(0);
  await expect(page.getByRole('textbox', {name: '文档正文', exact: true})).toBeVisible();
}
