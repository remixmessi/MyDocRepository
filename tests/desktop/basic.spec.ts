import {test,expect,_electron as electron,type ElectronApplication} from '@playwright/test';
import {mkdtemp,rm,readFile,open} from 'node:fs/promises';
import {tmpdir} from 'node:os';
import {join,resolve} from 'node:path';
import {Y} from '@mydoc/crdt';
import {execFile} from 'node:child_process';
import {promisify} from 'node:util';
let directory:string;
async function launch(){return electron.launch({executablePath:process.env.MYDOC_TEST_ELECTRON_PATH,args:[resolve('packages/frontend/apps/desktop')],env:{...process.env,MYDOC_DESKTOP_DB_PATH:join(directory,'data','mydoc.sqlite'),MYDOC_DESKTOP_PROFILE:join(directory,'profile')},timeout:30000});}
test.beforeEach(async()=>{directory=await mkdtemp(join(tmpdir(),'mydoc-electron-'));});
test.afterEach(async()=>{await rm(directory,{recursive:true,force:true,maxRetries:10,retryDelay:100});});
test('真实桌面首次空库、沙箱与 IPC、编辑附件主题、关闭保存及重启恢复',async()=>{
 let app:ElectronApplication|undefined;
 try{app=await launch();expect(await app.evaluate(()=>process.arch)).toBe(process.env.MYDOC_TEST_ELECTRON_ARCH??process.arch);const page=await app.firstWindow();await expect(page.getByRole('button',{name:'新建文档',exact:true}).first()).toBeVisible();const security=await app.evaluate(({BrowserWindow})=>(BrowserWindow.getAllWindows()[0].webContents as unknown as {getLastWebPreferences:()=>{contextIsolation:boolean;nodeIntegration:boolean;sandbox:boolean}}).getLastWebPreferences());expect(security.contextIsolation).toBe(true);expect(security.nodeIntegration).toBe(false);expect(security.sandbox).toBe(true);expect(await page.evaluate(()=>({bridge:!!window.mydocDesktop,node:typeof (window as unknown as {require?:unknown}).require}))).toEqual({bridge:true,node:'undefined'});
 await page.getByRole('button',{name:'新建文档',exact:true}).first().click();await page.getByLabel('文档标题',{exact:true}).fill('桌面 SQLite 文档');await page.getByRole('textbox',{name:'文档正文'}).fill('macOS / Windows 共用编辑器');await page.getByLabel('添加本地附件').setInputFiles({name:'桌面.txt',mimeType:'text/plain',buffer:Buffer.from('原始附件字节')});await expect(page.getByRole('button',{name:'下载 桌面.txt'})).toBeVisible();await page.getByRole('button',{name:'设置',exact:true}).click();await page.getByRole('dialog',{name:'设置',exact:true}).getByRole('button',{name:'外观',exact:true}).click();await page.getByRole('dialog',{name:'设置',exact:true}).getByRole('button',{name:'深色',exact:true}).click();await expect(page.getByRole('status').filter({hasText:'外观已保存'})).toBeVisible();await page.getByRole('dialog',{name:'设置',exact:true}).getByRole('button',{name:'关闭弹窗',exact:true}).click();await expect(page.getByRole('dialog')).toHaveCount(0);
 await page.getByRole('textbox',{name:'文档正文'}).press('ControlOrMeta+End');await page.getByRole('textbox',{name:'文档正文'}).press('Enter');await page.getByRole('textbox',{name:'文档正文'}).pressSequentially('关闭前的最终正文');await app.close();app=undefined;expect((await readFile(join(directory,'data','mydoc.sqlite'))).subarray(0,16).toString()).toBe('SQLite format 3\0');
 app=await launch();const restored=await app.firstWindow();await expect(restored.getByLabel('文档标题',{exact:true})).toHaveValue('桌面 SQLite 文档');await expect(restored.getByRole('textbox',{name:'文档正文'})).toContainText('关闭前的最终正文');await expect(restored.locator('html')).toHaveAttribute('data-theme','dark');await expect(restored.getByRole('button',{name:'下载 桌面.txt'})).toBeVisible();
 const rejected=await restored.evaluate(async()=>{try{await window.mydocDesktop!.storage({path:'/test-clear-assets',method:'DELETE'});return false;}catch{return true;}});expect(rejected).toBe(true);
 }finally{await app?.close();}
});
test('保存失败阻止退出，恢复后可正常关闭；macOS 关闭窗口后可重新打开',async()=>{const app=await launch();try{const page=await app.firstWindow();await expect(page.getByRole('button',{name:'新建文档',exact:true}).first()).toBeVisible();await page.getByRole('button',{name:'新建文档',exact:true}).first().click();await expect(page.getByRole('textbox',{name:'文档正文'})).toBeVisible();await app.evaluate(({dialog})=>{dialog.showMessageBox=async()=>({response:0,checkboxChecked:false});});await page.evaluate(()=>{(window as unknown as {cancelClose:()=>void}).cancelClose=window.mydocDesktop!.onPrepareClose(async()=>{throw new Error('模拟保存失败');});});await app.evaluate(({app})=>app.quit());await page.waitForTimeout(300);expect(app.windows()).toHaveLength(1);await expect(page.getByRole('textbox',{name:'文档正文'})).toBeVisible();await page.evaluate(()=>(window as unknown as {cancelClose:()=>void}).cancelClose());if(process.platform==='darwin'){await app.evaluate(({BrowserWindow})=>BrowserWindow.getAllWindows()[0].close());await expect.poll(()=>app.windows().length).toBe(0);await app.evaluate(({app})=>app.emit('activate'));const reopened=await app.firstWindow();await expect(reopened.getByRole('textbox',{name:'文档正文'})).toBeVisible();}}finally{await app.close();}});

test('桌面 SQLite 待提交队列在真实进程重启后保留并按序确认清理',async()=>{
 let app:ElectronApplication|undefined;
 const stop=async()=>{
  if(!app)return;
  const child=app.process();
  // Windows kill() only terminates the main process. Chromium children can keep
  // profile files and inherited stdio open, so kill the tree and wait for exit.
  const exited=()=>child.exitCode!==null||child.signalCode!==null;
  const exitResult=new Promise<Error|undefined>(resolve=>{
   if(exited()){resolve(undefined);return;}
   const finished=()=>{clearTimeout(timer);resolve(undefined);};
   const timer=setTimeout(()=>{child.removeListener('exit',finished);resolve(new Error('Electron 进程未在 10 秒内退出'));},10000);
   child.once('exit',finished);
  });
  if(!exited()){
   if(process.platform==='win32')await promisify(execFile)('taskkill',['/PID',String(child.pid),'/T','/F'],{timeout:10000});
   else child.kill('SIGKILL');
  }
  const error=await exitResult;
  if(error)throw error;
  app=undefined;
 };
 const waitForDatabaseUnlock=async()=>{const path=join(directory,'data','mydoc.sqlite');const deadline=Date.now()+15000;for(;;){try{const handle=await open(path,'r+');await handle.close();return;}catch(error){if(Date.now()>=deadline)throw error;await new Promise(resolve=>setTimeout(resolve,100));}}};
 const doc=new Y.Doc();doc.getText('body').insert(0,'进程重启后仍待提交');
 const update=Buffer.from(Y.encodeStateAsUpdate(doc));doc.destroy();
 const base64=update.toString('base64');
 const request=async(page:import('@playwright/test').Page,path:string,method='GET',body?:unknown)=>page.evaluate(async({path,method,body})=>{
   const reply=await window.mydocDesktop!.storage({path,method,contentType:body===undefined?undefined:'application/json',body:body===undefined?undefined:new TextEncoder().encode(JSON.stringify(body))});
   return {status:reply.status,text:new TextDecoder().decode(reply.body)};
 },{path,method,body});
 try{
  app=await launch();let page=await app.firstWindow();await expect(page.getByRole('button',{name:'新建文档',exact:true}).first()).toBeVisible();
  const path='/sync-queue?'+new URLSearchParams({server:'http://127.0.0.1:4180',workspace:'desktop-restart',id:'doc:restart'}).toString();
  const queued=await request(page,path,'POST',{update:base64,snapshot:base64});expect(queued.status).toBe(200);const seq=JSON.parse(queued.text).seq;expect(seq).toBeGreaterThan(0);
  await stop();await waitForDatabaseUnlock();app=await launch();page=await app.firstWindow();const restored=await request(page,path);expect(restored.status).toBe(200);const snapshot=JSON.parse(restored.text);expect(snapshot.rows).toHaveLength(1);expect(snapshot.rows[0].seq).toBe(seq);expect(snapshot.rows[0].attempts).toBe(0);
  const ack=await request(page,path,'DELETE',{seq});expect(ack.status).toBe(200);await stop();await waitForDatabaseUnlock();app=await launch();page=await app.firstWindow();const empty=await request(page,path);expect(JSON.parse(empty.text).rows).toEqual([]);
 }finally{await stop();}
});

test('桌面 PDF worker 和独立无权限预览、白板编辑与重新载入恢复',async()=>{const app=await launch();try{const page=await app.firstWindow();await expect(page.getByRole('button',{name:'新建文档',exact:true}).first()).toBeVisible();await page.getByRole('button',{name:'新建文档',exact:true}).first().click();await page.getByLabel('添加本地附件').setInputFiles({name:'报告.pdf',mimeType:'application/pdf',buffer:await readFile('tests/fixtures/pdf-preview.pdf')});await page.getByRole('button',{name:'预览 PDF 报告.pdf',exact:true}).click();const preview=page.getByRole('dialog',{name:'PDF 附件预览'});await expect(preview.getByRole('img',{name:'PDF 第 1 页',exact:true})).toBeVisible();const viewerEvent=app.waitForEvent('window');await preview.getByRole('button',{name:'新窗口打开',exact:true}).click();const viewer=await viewerEvent;expect(await viewer.evaluate(()=>({bridge:!!window.mydocDesktop,node:typeof (window as unknown as {require?:unknown}).require}))).toEqual({bridge:false,node:'undefined'});await viewer.close();await page.getByLabel('关闭 PDF 预览').click();await page.getByRole('button',{name:'白板',exact:true}).click();await page.getByRole('button',{name:'矩形',exact:true}).click();const canvas=page.getByLabel('白板画布',{exact:true}),box=(await canvas.boundingBox())!;await page.mouse.move(box.x+120,box.y+120);await page.mouse.down();await page.mouse.move(box.x+320,box.y+240,{steps:5});await page.mouse.up();await page.getByLabel('元素文字').fill('桌面白板');await expect(page.getByText('正在保存…')).toHaveCount(0);await page.reload();await page.getByRole('button',{name:'白板',exact:true}).click();await page.getByLabel('白板元素',{exact:true}).selectOption({label:'1. 桌面白板'});await expect(page.getByLabel('元素文字')).toHaveValue('桌面白板');}finally{await app.close();}});
