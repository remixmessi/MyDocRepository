import {test,expect,_electron as electron} from '@playwright/test';
import {mkdtemp,mkdir,rm} from 'node:fs/promises';
import {tmpdir} from 'node:os';
import {join,resolve} from 'node:path';
import {releaseTarget} from '../../packages/common/nbstore/src/releases';

test('服务器应用内更新：地址保存、下载进度、互斥、状态恢复与安装失败重试',async()=>{
 const dir=await mkdtemp(join(tmpdir(),'mydoc-server-update-'));
 const app=await electron.launch({args:[resolve('packages/frontend/apps/desktop')],env:{...process.env,MYDOC_DESKTOP_DB_PATH:join(dir,'data','mydoc.sqlite'),MYDOC_DESKTOP_PROFILE:join(dir,'profile')}});
 try{
  const page=await app.firstWindow();
  // Stub native installation only, never launch a real installer from a UI test.
  await app.evaluate(({ipcMain,BrowserWindow})=>{
   const state={version:'0.1.1',source:'server',serverUrl:'',channel:'服务器更新',enabled:true,automaticChecks:true,phase:'idle',latestVersion:undefined as string|undefined,releaseNotes:'修复服务器更新',percent:0,transferred:0,total:0,message:undefined as string|undefined,retry:undefined as string|undefined};
   const testState={addresses:[] as string[],finish(){state.phase='downloaded';state.percent=100;BrowserWindow.getAllWindows()[0].webContents.send('mydoc:update-state',state);}};
   let installs=0;
   (globalThis as unknown as {serverUpdateTest:typeof testState}).serverUpdateTest=testState;
   ipcMain.removeHandler('mydoc:update-status');ipcMain.handle('mydoc:update-status',()=>state);
   ipcMain.removeHandler('mydoc:update-server');ipcMain.handle('mydoc:update-server',(_event,address:string)=>{testState.addresses.push(address);state.serverUrl=address;state.phase='available';state.latestVersion='9.0.0';BrowserWindow.getAllWindows()[0].webContents.send('mydoc:update-state',state);return state;});
   ipcMain.removeHandler('mydoc:update-action');ipcMain.handle('mydoc:update-action',(_event,action)=>{
    if(action==='download'){state.phase='downloading';state.percent=46;state.transferred=46*1024*1024;state.total=100*1024*1024;}
    if(action==='install'){installs++;state.phase=installs===1?'downloaded':'installing';state.retry=installs===1?'install':undefined;state.message=installs===1?'保护备份失败，请重试。':undefined;}
    BrowserWindow.getAllWindows()[0].webContents.send('mydoc:update-state',state);return state;
   });
  });
  await page.getByRole('button',{name:'版本与更新',exact:true}).click();
  const dialog=page.getByRole('dialog',{name:'版本与更新'}),panel=dialog.getByRole('region',{name:'服务器版本检查',exact:true});
  const device=await page.evaluate(()=>({arch:window.mydocDesktop!.arch,platform:window.mydocDesktop!.platform}));
  await expect(panel.locator('.client-target')).toContainText(releaseTarget(device.platform,device.arch));
  const input=panel.getByRole('textbox',{name:'更新服务器地址'}),check=panel.getByRole('button',{name:'检查服务器版本',exact:true});
  await expect(check).toBeDisabled();await input.fill('docs.example.com/mydoc');await panel.getByRole('button',{name:'保存地址',exact:true}).click();await expect(input).toHaveValue('http://docs.example.com/mydoc');
  await check.click();await expect(panel.getByRole('status')).toHaveText('发现新版本');expect(await app.evaluate(()=>(globalThis as unknown as {serverUpdateTest:{addresses:string[]}}).serverUpdateTest.addresses)).toEqual(['http://docs.example.com/mydoc']);
  await expect(panel.getByRole('link',{name:'下载安装包'})).toHaveCount(0);await panel.getByRole('button',{name:'下载并更新'}).click();await expect(panel.getByRole('progressbar',{name:'服务器更新下载进度'})).toHaveAttribute('value','46');await expect(input).toBeDisabled();await expect(dialog.getByRole('button',{name:'检查 GitHub 版本',exact:true})).toBeDisabled();
  await dialog.getByRole('button',{name:'关闭',exact:true}).click();await page.getByRole('button',{name:'版本与更新',exact:true}).click();await expect(panel.getByRole('progressbar')).toHaveAttribute('value','46');await expect(input).toHaveValue('http://docs.example.com/mydoc');
  await app.evaluate(()=>(globalThis as unknown as {serverUpdateTest:{finish:()=>void}}).serverUpdateTest.finish());await expect(input).toBeDisabled();await expect(dialog.getByRole('button',{name:'检查 GitHub 版本',exact:true})).toBeDisabled();
  await panel.getByRole('button',{name:'重启并安装',exact:true}).click();await expect(panel.getByRole('alert')).toContainText('保护备份失败');await expect(dialog.getByRole('button',{name:'稍后安装'})).toBeEnabled();
  await page.evaluate(()=>document.documentElement.dataset.theme='dark');await app.evaluate(({BrowserWindow})=>{const window=BrowserWindow.getAllWindows()[0];window.setMinimumSize(360,600);window.setSize(420,800);});await panel.getByRole('button',{name:'重试安装',exact:true}).scrollIntoViewIfNeeded();await mkdir('artifacts',{recursive:true});await page.screenshot({path:'artifacts/server-in-app-update-dark-narrow.png'});expect(await page.evaluate(()=>document.documentElement.scrollWidth<=innerWidth)).toBe(true);
  await panel.getByRole('button',{name:'重试安装',exact:true}).click();await expect(panel.getByRole('status')).toHaveText('正在安装更新，请稍候…');await expect(dialog.getByRole('button',{name:'关闭',exact:true})).toBeDisabled();
 }finally{await app.close();await rm(dir,{recursive:true,force:true});}
});
