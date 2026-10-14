import {test,expect,_electron as electron,type ElectronApplication} from '@playwright/test';
import {mkdtemp,rm,writeFile,truncate} from 'node:fs/promises';
import {createRequire} from 'node:module';import {execFile} from 'node:child_process';import {promisify} from 'node:util';
import {tmpdir} from 'node:os';import {join,resolve} from 'node:path';
let directory:string;
const project=resolve('packages/frontend/apps/desktop');
async function launch(files:string[]=[]){return electron.launch({args:[project,...files],env:{...process.env,MYDOC_DESKTOP_DB_PATH:join(directory,'data','mydoc.sqlite'),MYDOC_DESKTOP_PROFILE:join(directory,'profile')},timeout:30000});}
async function finish(page:import('@playwright/test').Page){const dialog=page.getByRole('dialog',{name:'导入文档',exact:true});await dialog.getByRole('button',{name:'开始导入',exact:true}).click();await dialog.getByRole('button',{name:'完成',exact:true}).click();}
test.beforeEach(async()=>{directory=await mkdtemp(join(tmpdir(),'mydoc-files-'));});test.afterEach(async()=>{await rm(directory,{recursive:true,force:true});});
test('冷启动多文件预览、确认导入、菜单打开、重复导入跳过与刷新',async()=>{
 const a=join(directory,'第一篇 中文.md'),b=join(directory,'第二篇.markdown');await writeFile(a,'启动正文 A');await writeFile(b,'启动正文 B');let app:ElectronApplication|undefined;
 try{app=await launch([a,b]);const page=await app.firstWindow();const preview=page.getByLabel('导入预览');await expect(preview).toContainText('第一篇 中文');await expect(preview).toContainText('第二篇');await expect(page.getByRole('navigation',{name:'文档列表'}).locator('button.doc-link')).toHaveCount(0);await finish(page);await expect(page.getByRole('textbox',{name:'文档正文'})).toHaveText('启动正文 B');
 await app.evaluate(({dialog},path)=>{dialog.showOpenDialog=async()=>({canceled:false,filePaths:[path]});},a);
 await app.evaluate(({Menu,BrowserWindow})=>{const item=Menu.getApplicationMenu()!.items.find(item=>item.label==='文件')!.submenu!.items.find(item=>item.label==='打开文件…')!;item.click(item,BrowserWindow.getAllWindows()[0],{ctrlKey:false,metaKey:false,shiftKey:false,altKey:false});});
 await expect(preview).toContainText('相同内容已导入');await finish(page);await expect(page.getByRole('navigation',{name:'文档列表'}).locator('button.doc-link')).toHaveCount(2);await page.reload();await expect(page.getByRole('navigation',{name:'文档列表'}).locator('button.doc-link')).toHaveCount(2);
 }finally{await app?.close();}
});
test('open-file 和 second-instance 排队、不覆盖弹窗、重载恢复待处理请求、取消不写入',async()=>{
 const a=join(directory,'等待甲.md'),b=join(directory,'等待乙.md');await writeFile(a,'A');await writeFile(b,'B');const app=await launch();try{const page=await app.firstWindow();await page.getByRole('button',{name:'设置',exact:true}).click();await page.getByRole('dialog',{name:'设置',exact:true}).getByRole('button',{name:'外观',exact:true}).click();
 await app.evaluate(({app},paths)=>{app.emit('open-file',{preventDefault(){}},paths[0]);},[a,b,project]);
 await promisify(execFile)(createRequire(import.meta.url)('electron'),[project,b],{env:{...process.env,MYDOC_DESKTOP_DB_PATH:join(directory,'data','mydoc.sqlite'),MYDOC_DESKTOP_PROFILE:join(directory,'profile')},timeout:15000});
 await app.evaluate(({app},path)=>{app.emit('open-file',{preventDefault(){}},path);},b);
 await expect(page.getByRole('dialog',{name:'设置',exact:true})).toBeVisible();await expect(page.getByRole('dialog',{name:'导入文档',exact:true})).toHaveCount(0);
 await page.reload();await expect(page.getByLabel('导入预览')).toContainText('等待甲');await page.getByRole('dialog',{name:'导入文档',exact:true}).getByRole('button',{name:'取消',exact:true}).click();await expect(page.getByLabel('导入预览')).toContainText('等待乙');await finish(page);await expect(page.getByRole('textbox',{name:'文档标题'})).toHaveValue('等待乙');await expect(page.getByRole('navigation',{name:'文档列表'})).not.toContainText('等待甲');expect(app.windows()).toHaveLength(1);
 }finally{await app.close();}
});
test('无效与超限文件反馈、沙箱不暴露任意路径读取',async()=>{
 const huge=join(directory,'huge.zip');await writeFile(huge,'');await truncate(huge,51*1024*1024);const app=await launch([join(directory,'missing.md'),huge]);try{const page=await app.firstWindow();await expect(page.locator('.error-banner')).toContainText('文件不存在');await expect(page.locator('.error-banner')).toContainText('50 MB');expect(await page.evaluate(()=>Object.keys(window.mydocDesktop!))).not.toContain('readFile');
 await expect(page.getByRole('dialog',{name:'导入文档',exact:true})).toHaveCount(0);await expect(page.getByRole('button',{name:'新建文档',exact:true}).first()).toBeEnabled();
 }finally{await app.close();}
});
test('导入写入进行中阻止退出，完成后正常保存关闭',async()=>{
 const path=join(directory,'正在导入.md');await writeFile(path,'导入结束后保存正文');const app=await launch([path]);try{const page=await app.firstWindow();await expect(page.getByLabel('导入预览')).toBeVisible();
 await app.evaluate(({dialog})=>{
  const state=globalThis as unknown as {releaseImport?:()=>void;closeMessages:number;importStarted:boolean};state.closeMessages=0;state.importStarted=false;
  dialog.showMessageBox=async()=>{state.closeMessages++;return {response:0,checkboxChecked:false};};
  const original=globalThis.fetch;globalThis.fetch=async(...args:Parameters<typeof fetch>)=>{if(args[1]?.method==='POST'&&String(args[0]).includes('/api/storage/document')){globalThis.fetch=original;state.importStarted=true;await new Promise<void>(resolve=>{state.releaseImport=resolve;});}return original(...args);};
 });
 // Wait for the held storage write before requesting quit, so the import cannot finish first.
 await page.getByRole('button',{name:'开始导入',exact:true}).click();await expect(page.getByRole('button',{name:'停止导入',exact:true})).toBeVisible();await expect.poll(()=>app.evaluate(()=>((globalThis as unknown as {importStarted:boolean}).importStarted))).toBe(true);await app.evaluate(({app})=>app.quit());
 await expect.poll(()=>app.evaluate(()=>((globalThis as unknown as {closeMessages:number}).closeMessages))).toBeGreaterThan(0);expect(app.windows()).toHaveLength(1);await expect(page.getByRole('dialog',{name:'导入文档',exact:true})).toBeVisible();
 await app.evaluate(()=>(globalThis as unknown as {releaseImport?:()=>void}).releaseImport?.());await page.getByRole('button',{name:'完成',exact:true}).click();await expect(page.getByRole('textbox',{name:'文档正文'})).toHaveText('导入结束后保存正文');
 }finally{await app.evaluate(()=>(globalThis as unknown as {releaseImport?:()=>void}).releaseImport?.()).catch(()=>{});await app.close();}
});
