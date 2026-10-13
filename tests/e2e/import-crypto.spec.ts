import {test,expect} from '@playwright/test';

test('缺少 crypto.subtle 时导入、重复识别与刷新保存正常',async({page,context})=>{
 await context.setExtraHTTPHeaders({'X-Mydoc-Test-Scope':crypto.randomUUID()});
 await page.goto('/');
 await expect(page.getByRole('button',{name:'新建文档'}).first()).toBeEnabled();
 const open=async()=>{
  await page.getByLabel('工作区选项').click();
  await page.getByRole('button',{name:'导入 Markdown',exact:true}).click();
 };
 const file={name:'兼容导入.md',mimeType:'text/markdown',buffer:Buffer.from('# 中文标题\n\n兼容导入正文')};
 await open();
 await page.getByLabel('选择导入文件',{exact:true}).setInputFiles(file);
 await expect(page.getByLabel('导入预览')).toBeVisible();
 await page.getByRole('button',{name:'开始导入',exact:true}).click();
 await expect(page.getByRole('status').filter({hasText:'已导入 1 篇'})).toBeVisible();
 await page.getByRole('button',{name:'完成',exact:true}).click();

 await page.addInitScript(()=>Object.defineProperty(globalThis.crypto,'subtle',{configurable:true,value:undefined}));
 await page.reload();
 await expect(page.getByRole('textbox',{name:'文档正文'})).toContainText('兼容导入正文');
 expect(await page.evaluate(()=>globalThis.crypto.subtle)).toBeUndefined();
 await open();
 await page.getByLabel('选择导入文件',{exact:true}).setInputFiles(file);
 await expect(page.getByLabel('导入预览')).toContainText('相同内容已导入');
 await page.getByRole('button',{name:'开始导入',exact:true}).click();
 await expect(page.getByRole('status').filter({hasText:'已导入 0 篇，跳过 1 篇'})).toBeVisible();
 await page.getByRole('button',{name:'完成',exact:true}).click();

 await open();
 await page.getByLabel('选择导入文件',{exact:true}).setInputFiles({...file,name:'新文件.md',buffer:Buffer.from('无 subtle 新导入正文')});
 await expect(page.getByLabel('导入预览')).toBeVisible();
 await page.getByRole('button',{name:'开始导入',exact:true}).click();
 await expect(page.getByRole('status').filter({hasText:'已导入 1 篇'})).toBeVisible();
 await page.getByRole('button',{name:'完成',exact:true}).click();
 await page.reload();
 await expect(page.getByRole('textbox',{name:'文档正文'})).toHaveText('无 subtle 新导入正文');
});
