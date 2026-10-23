import { test as base, expect, type Page } from '@playwright/test';
import {newNote,confirmNewFile} from './fixtures';

const runId = crypto.randomUUID();
const test = base.extend<{ projectIsolation: void }>({
  projectIsolation: [async ({ context }, use, info) => {
    await context.setExtraHTTPHeaders({ 'X-Mydoc-Test-Scope': `${runId}-${info.testId}-${info.retry}` });
    await use();
  }, { auto: true }],
});

async function folderActions(page: Page, name: string, action: string) {
  await page.getByRole('button', { name: `管理文件夹 ${name}`, exact: true }).click();
  await page.getByRole('menu').getByRole('menuitem', { name: action, exact: true }).click();
}
async function documentActions(page: Page, name: string, action: string) {
  await page.getByRole('button', { name: `文档操作 ${name}`, exact: true }).click();
  await page.getByRole('menu').getByRole('menuitem', { name: action, exact: true }).click();
}

test('文件夹内新建、导入，文档重命名、移动和回收均保留内容', async ({ page }) => {
  await page.goto('/');
  await page.getByRole('button', { name: '新建文件夹', exact: true }).click();
  await page.getByRole('textbox', { name: '文件夹名称' }).fill('项目');
  await page.getByRole('button', { name: '保存', exact: true }).click();
  await folderActions(page, '项目', '新建子文件夹');
  await expect(page.getByLabel('上级文件夹').locator('option:checked')).toHaveText('项目');
  await page.getByRole('textbox', { name: '文件夹名称' }).fill('资料');
  await page.getByRole('button', { name: '保存', exact: true }).click();
  await folderActions(page, '资料', '新建文档');
  await confirmNewFile(page);
  await expect(page.getByLabel('文档所在文件夹').locator('option:checked')).toHaveText('项目 / 资料');
  await page.getByRole('textbox', { name: '文档正文' }).fill('移动和重命名后保留正文');
  await documentActions(page, '未命名', '重命名文档');
  await page.getByRole('textbox', { name: '文档名称' }).fill('项目笔记');
  await page.getByRole('button', { name: '保存', exact: true }).click();
  await expect(page.getByRole('textbox', { name: '文档标题' })).toHaveValue('项目笔记');
  await page.getByRole('button', { name: '项目', exact: true }).click();
  // Moving the selected document must reopen a collapsed destination.
  await newNote(page);
  await documentActions(page, '未命名', '移动文档');
  await page.getByLabel('目标文件夹').selectOption({ label: '项目 / 资料' });
  await page.getByRole('button', { name: '保存', exact: true }).click();
  await expect(page.getByRole('navigation').getByRole('button', { name: '未命名', exact: true })).toBeVisible();
  await documentActions(page, '项目笔记', '移动文档');
  await page.getByLabel('目标文件夹').selectOption('');
  await page.getByRole('button', { name: '保存', exact: true }).click();
  await folderActions(page, '资料', '导入 Markdown / 文件');
  await expect(page.getByLabel('导入目标文件夹').locator('option:checked')).toHaveText('项目 / 资料');
  await page.getByLabel('选择导入文件', { exact: true }).setInputFiles({ name: '导入笔记.md', mimeType: 'text/markdown', buffer: Buffer.from('# 导入正文\n\n文件夹内导入') });
  await page.getByRole('button', { name: '开始导入', exact: true }).click();
  await expect(page.getByRole('status').filter({ hasText: '已导入 1 篇' })).toBeVisible();
  await page.getByRole('button', { name: '完成', exact: true }).click();
  await page.reload();
  await expect(page.getByLabel('文档所在文件夹').locator('option:checked')).toHaveText('项目 / 资料');
  await expect(page.getByRole('textbox', { name: '文档正文' })).toContainText('文件夹内导入');
  await page.getByRole('navigation').getByRole('button', { name: '项目笔记', exact: true }).click();
  await expect(page.getByRole('textbox', { name: '文档正文' })).toHaveText('移动和重命名后保留正文');
  await expect(page.getByLabel('文档所在文件夹').locator('option:checked')).toHaveText('未分类');
  await documentActions(page, '项目笔记', '删除文档');
  await page.getByRole('button', { name: '确认删除', exact: true }).click();
  await expect(page.getByRole('navigation').getByRole('button', { name: '项目笔记', exact: true })).toHaveCount(0);
  await page.getByRole('button', { name: /回收站/ }).click();
  await expect(page.getByRole('dialog')).toContainText('项目笔记');
});

test('移动文件夹排除自身及子目录，手机可以打开操作入口', async ({ page }) => {
  await page.setViewportSize({ width: 390, height: 844 });
  await page.goto('/');
  await page.getByRole('button', { name: '文档导航', exact: true }).click();
  await page.getByRole('button', { name: '新建文件夹', exact: true }).click();
  await page.getByRole('textbox', { name: '文件夹名称' }).fill('父目录');
  await page.getByRole('button', { name: '保存', exact: true }).click();
  await folderActions(page, '父目录', '新建子文件夹');
  await page.getByRole('textbox', { name: '文件夹名称' }).fill('子目录');
  await page.getByRole('button', { name: '保存', exact: true }).click();
  await folderActions(page, '父目录', '重命名 / 移动 / 删除文件夹');
  await expect(page.getByLabel('上级文件夹').locator('option')).toHaveCount(1);
  await page.getByRole('button', { name: '取消', exact: true }).click();
  const trigger = page.getByRole('button', { name: '管理文件夹 父目录', exact: true });
  await trigger.click();
  await expect(page.getByRole('dialog')).toHaveCount(0);
  await expect(page.getByRole('menu')).toBeVisible();
  await page.keyboard.press('ArrowDown');
  await expect(page.getByRole('menuitem', { name: '新建子文件夹', exact: true })).toBeFocused();
  await page.keyboard.press('Escape');
  await expect(trigger).toBeFocused();
  await expect(page.getByRole('menu')).toHaveCount(0);
  await trigger.click();
  await trigger.click();
  await expect(page.getByRole('menu')).toHaveCount(0);
  await trigger.click();
  await page.screenshot({ path: 'artifacts/folder-actions-mobile.png', fullPage: true });
  const bounds = await page.getByRole('menu').boundingBox();
  expect(bounds!.x).toBeGreaterThanOrEqual(0);
  expect(bounds!.x + bounds!.width).toBeLessThanOrEqual(390);
  await page.getByRole('button', { name: '搜索文档', exact: true }).click();
  await expect(page.getByRole('menu')).toHaveCount(0);
});
