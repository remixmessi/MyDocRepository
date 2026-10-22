import {test,expect,type Page,type BrowserContext,type Locator} from '@playwright/test';
import {readFile} from 'node:fs/promises';
import {zipSync,strToU8} from 'fflate';

// Runs with playwright.office.config.ts: office import needs /api/cloud and an account.
const suffix=crypto.randomUUID();let cookies:Awaited<ReturnType<BrowserContext['cookies']>>|undefined;
async function login(page:Page,baseURL:string){
 if(!cookies){const response=await page.context().request.post(baseURL+'/api/cloud/register',{data:{name:'drop',email:'drop'+suffix+'@example.com',password:'test-password-12345'}});expect(response.ok()).toBe(true);cookies=await page.context().cookies();}
 else await page.context().addCookies(cookies);
 await page.goto('/#workspace=local');await expect(page.getByRole('button',{name:'导入文件',exact:true})).toBeEnabled();
}
type Input={name:string;bytes:Buffer;type?:string};
async function drop(target:Locator,files:Input[]){await target.first().evaluate((element,inputs)=>{const data=new DataTransfer();for(const file of inputs)data.items.add(new File([Uint8Array.from(file.bytes)],file.name,{type:file.type??'application/octet-stream'}));for(const type of ['dragenter','dragover','drop'])element.dispatchEvent(new DragEvent(type,{bubbles:true,cancelable:true,dataTransfer:data}));},files.map(file=>({...file,bytes:[...file.bytes]})));}
const docx=(text:string)=>Buffer.from(zipSync({'word/document.xml':strToU8(`<w:document><w:body><w:p><w:r><w:t>${text}</w:t></w:r></w:p></w:body></w:document>`)}));
const xlsx=(text:string)=>Buffer.from(zipSync({'xl/worksheets/sheet1.xml':strToU8(`<worksheet><sheetData><row><c t="inlineStr"><is><t>${text}</t></is></c></row></sheetData></worksheet>`)}));
const pptx=(text:string)=>Buffer.from(zipSync({'ppt/slides/slide1.xml':strToU8(`<p:sld><p:sp><a:t>${text}</a:t></p:sp></p:sld>`)}));
const list=(page:Page)=>page.getByRole('navigation',{name:'文档列表'});

test('应用任意位置拖入 Word / Excel 直接导入，无需再点击；旧版 .doc 给出原因',async({page,baseURL})=>{
 await login(page,baseURL!);
 await drop(page.locator('.sidebar-heading'),[{name:'拖入报告.docx',bytes:docx('拖入的Word正文')},{name:'拖入表格.xlsx',bytes:xlsx('拖入的表格')}]);
 await expect(page.getByRole('dialog',{name:'导入办公文件',exact:true})).toHaveCount(0);
 await expect(list(page).getByRole('button',{name:'拖入报告.docx',exact:true})).toBeVisible();await expect(list(page).getByRole('button',{name:'拖入表格.xlsx',exact:true})).toBeVisible();
 await expect(page.getByLabel('第 1 行第 1 列',{exact:true})).toHaveValue('拖入的表格');
 await list(page).getByRole('button',{name:'拖入报告.docx',exact:true}).click();await expect(page.getByLabel('第 1 段',{exact:true})).toHaveText('拖入的Word正文');
 await drop(page.locator('main'),[{name:'旧格式.doc',bytes:Buffer.from('old')}]);await expect(page.locator('.error-banner')).toContainText('另存为 .docx');await expect(page.getByRole('dialog')).toHaveCount(0);
 await page.getByRole('button',{name:'导入文件',exact:true}).click();const dialog=page.getByRole('dialog',{name:'导入办公文件',exact:true});
 await drop(dialog,[{name:'旧格式.doc',bytes:Buffer.from('old')},{name:'演示.pptx',bytes:pptx('x')}]);await expect(dialog.getByRole('alert')).toContainText('旧格式.doc：暂不支持旧版');await expect(dialog.getByRole('button',{name:'导入 2 个文件',exact:true})).toBeDisabled();
 await drop(dialog.locator('.dialog-heading'),[{name:'演示.pptx',bytes:pptx('弹窗拖入演示')}]);await dialog.getByRole('button',{name:'导入',exact:true}).click();await expect(dialog).toHaveCount(0);
 await expect(page.locator('.presentation-object textarea')).toHaveValue('弹窗拖入演示');
 await page.reload();await expect(list(page).getByRole('button',{name:'拖入表格.xlsx',exact:true})).toBeVisible();
});

test('CSV 拖入先预览选项，导入 Markdown 窗口拖入 Word 转为办公导入，批量失败列出原因',async({page,baseURL})=>{
 await login(page,baseURL!);
 await drop(page.locator('body'),[{name:'拖入数据.csv',bytes:Buffer.from('编号,名称\n001,甲\n')}]);
 let dialog=page.getByRole('dialog',{name:'导入办公文件',exact:true});await expect(dialog.getByLabel('导入内容预览')).toContainText('2 行');await dialog.getByRole('button',{name:'导入',exact:true}).click();await expect(dialog).toHaveCount(0);
 await expect(page.getByLabel('第 2 行第 1 列',{exact:true})).toHaveValue('001');
 await page.getByLabel('工作区选项',{exact:true}).click();await page.locator('.workspace-menu-content').getByRole('button',{name:'导入 Markdown',exact:true}).click();
 const markdown=page.getByRole('dialog',{name:'导入文档',exact:true});await drop(markdown,[{name:'转交.docx',bytes:docx('转交正文')}]);await expect(markdown).toHaveCount(0);
 await expect(list(page).getByRole('button',{name:'转交.docx',exact:true})).toBeVisible();
 await page.getByRole('button',{name:'导入文件',exact:true}).click();dialog=page.getByRole('dialog',{name:'导入办公文件',exact:true});
 await drop(dialog,[{name:'好文件.docx',bytes:docx('好')},{name:'坏文件.docx',bytes:Buffer.from('not zip')}]);await dialog.getByRole('button',{name:'导入 2 个文件',exact:true}).click();
 await expect(dialog.getByLabel('办公文件导入结果')).toContainText('坏文件.docx');await expect(dialog.getByRole('status')).toContainText('已导入 1 个，未导入 1 个');await dialog.getByRole('button',{name:'完成',exact:true}).click();await expect(dialog).toHaveCount(0);
 await expect(list(page).getByRole('button',{name:'好文件.docx',exact:true})).toBeVisible();
});

test('TXT、办公备份、草稿、文档模板弹窗支持拖入，备份 JSON 拖到应用即恢复',async({page,baseURL})=>{
 await login(page,baseURL!);
 await page.getByLabel('工作区选项',{exact:true}).click();await page.locator('.workspace-menu-content').getByRole('button',{name:'导入 TXT',exact:true}).click();
 let dialog=page.getByRole('dialog',{name:'导入 TXT',exact:true});await drop(dialog,[{name:'拖入文本.txt',bytes:Buffer.from('拖入的文本')}]);await expect(dialog.getByLabel('文本导入预览')).toHaveText('拖入的文本');await dialog.getByRole('button',{name:'导入并保留原件',exact:true}).click();await expect(dialog).toHaveCount(0);
 await expect(list(page).getByRole('button',{name:'拖入文本.txt',exact:true})).toBeVisible();
 // Office backup round trip by drop.
 await page.getByRole('button',{name:'文档操作 拖入文本.txt',exact:true}).click();await page.getByRole('menuitem',{name:'备份文件',exact:true}).click();
 const downloaded=page.waitForEvent('download');await page.getByRole('dialog',{name:'备份办公文件',exact:true}).getByRole('button',{name:'下载备份',exact:true}).click();const officeBackup=await readFile((await(await downloaded).path())!);
 await page.getByLabel('工作区选项',{exact:true}).click();await page.locator('.workspace-menu-content').getByRole('button',{name:'从文件备份恢复',exact:true}).click();
 dialog=page.getByRole('dialog',{name:'从备份恢复办公文件',exact:true});await drop(dialog,[{name:'文本.mydoc-office.json',bytes:officeBackup,type:'application/json'}]);await expect(dialog).toContainText('文本.mydoc-office.json');await dialog.getByRole('button',{name:'恢复为新文件',exact:true}).click();await expect(dialog).toHaveCount(0);
 await expect(list(page).getByRole('button',{name:'拖入文本.txt',exact:true})).toHaveCount(2);
 // Dropping the office backup on the app opens the same restore dialog with the file chosen.
 await drop(page.locator('.sidebar-heading'),[{name:'文本.mydoc-office.json',bytes:officeBackup,type:'application/json'}]);dialog=page.getByRole('dialog',{name:'从备份恢复办公文件',exact:true});await expect(dialog).toContainText('文本.mydoc-office.json');await dialog.getByRole('button',{name:'取消',exact:true}).click();
 // Document template: export a built-in one, then drop it into the dialog and onto the app.
 await page.getByRole('button',{name:'文档模板',exact:true}).click();dialog=page.getByRole('dialog',{name:'文档模板',exact:true});const pending=page.waitForEvent('download');await dialog.getByRole('button',{name:'导出文档模板',exact:true}).click();const template=await readFile((await(await pending).path())!);
 await drop(dialog,[{name:'模板.json',bytes:template,type:'application/json'}]);await expect(dialog.getByRole('button',{name:/自定义模板$/})).toHaveCount(1);await page.keyboard.press('Escape');
 await drop(page.locator('.sidebar-heading'),[{name:'模板.json',bytes:template,type:'application/json'}]);dialog=page.getByRole('dialog',{name:'文档模板',exact:true});await expect(dialog.getByRole('button',{name:/自定义模板$/})).toHaveCount(2);await page.keyboard.press('Escape');
 // Draft dialog accepts a dropped draft; non-JSON gives the parse error.
 await page.getByLabel('工作区选项',{exact:true}).click();await page.locator('.workspace-menu-content').getByRole('button',{name:'从草稿恢复',exact:true}).click();dialog=page.getByRole('dialog',{name:'从草稿恢复',exact:true});
 await drop(dialog,[{name:'草稿.json',bytes:Buffer.from('{}'),type:'application/json'}]);await expect(dialog.getByRole('alert')).toContainText('不是支持的草稿文件');await page.keyboard.press('Escape');
 // Whole workspace backup dropped on the app restores to a new workspace.
 const backup=page.waitForEvent('download');await page.getByLabel('工作区选项',{exact:true}).click();await page.getByRole('button',{name:'备份工作区',exact:true}).click();const bytes=await readFile((await(await backup).path())!);
 await drop(page.locator('.sidebar-heading'),[{name:'备份.json',bytes,type:'application/json'}]);await expect(page.getByLabel('当前工作区',{exact:true})).toContainText('（恢复）');
 await expect(list(page).getByRole('button',{name:'拖入文本.txt',exact:true})).toHaveCount(2);
});

test('弹窗打开时背景拖放被拦截，网页剪藏拖入 HTML 需先填网址',async({page,baseURL})=>{
 await login(page,baseURL!);
 await page.getByRole('button',{name:'导入文件',exact:true}).click();const dialog=page.getByRole('dialog',{name:'导入办公文件',exact:true});
 await drop(page.locator('.sidebar-heading'),[{name:'背景.docx',bytes:docx('x')}]);await expect(dialog.getByText('背景.docx')).toHaveCount(0);await expect(list(page).getByRole('button',{name:'背景.docx',exact:true})).toHaveCount(0);await dialog.getByRole('button',{name:'取消',exact:true}).click();
 await page.getByLabel('工作区选项',{exact:true}).click();await page.locator('.workspace-menu-content').getByRole('button',{name:'网页剪藏',exact:true}).click();const clip=page.getByRole('dialog',{name:'网页剪藏',exact:true});const html=Buffer.from('<html><head><title>拖入网页</title></head><body><article><h1>拖入网页</h1><p>拖入的剪藏正文内容足够长，可以被识别为文章主体。拖入的剪藏正文内容足够长。</p></article></body></html>');
 await drop(clip,[{name:'page.html',bytes:html,type:'text/html'}]);await expect(clip.getByRole('alert')).toContainText('请先填写网页原始网址');
 await clip.getByLabel('剪藏网页网址').fill('https://example.com/drop');await drop(clip,[{name:'page.html',bytes:html,type:'text/html'}]);await expect(clip.getByLabel('剪藏标题')).toHaveValue('拖入网页');
});
