import {test,expect,_electron as electron} from '@playwright/test';
import {createServer} from 'node:http';
import {mkdtemp,mkdir,rm} from 'node:fs/promises';
import {tmpdir} from 'node:os';
import {join,resolve} from 'node:path';
import {releaseTarget} from '../../packages/common/nbstore/src/releases';

test('服务器检查使用真实运行架构，保存地址并拒绝不匹配的安装包',async()=>{
 const dir=await mkdtemp(join(tmpdir(),'mydoc-releases-'));
 const requests:URL[]=[];let mode:'latest'|'mismatch'|'empty'|'error'='latest';
 const server=createServer((req,res)=>{
  res.setHeader('Access-Control-Allow-Origin','*');res.setHeader('Content-Type','application/json');
  const url=new URL(req.url!,'http://localhost');requests.push(url);
  if(mode==='error'){res.statusCode=503;res.end('{}');return;}
  if(mode==='empty'){res.end('null');return;}
  const arch=mode==='mismatch'?(url.searchParams.get('arch')==='arm64'?'x64':'arm64'):url.searchParams.get('arch');
  res.end(JSON.stringify({id:'12345678-1234-1234-1234-123456789abc',version:arch==='arm64'?'0.2.0':'0.3.0',platform:url.searchParams.get('platform'),arch,filename:`MyDoc-${arch}.dmg`,notes:'改善编辑体验\n修复导入问题',bytes:2048,sha256:'a'.repeat(64),created:Date.UTC(2026,9,13),downloadPath:'/api/cloud/releases/download?id=12345678-1234-1234-1234-123456789abc'}));
 });
 await new Promise<void>(resolve=>server.listen(0,'127.0.0.1',resolve));
 const address=server.address() as {port:number};const host=`127.0.0.1:${address.port}`;
 const app=await electron.launch({args:[resolve('packages/frontend/apps/desktop')],env:{...process.env,MYDOC_DESKTOP_DB_PATH:join(dir,'data','mydoc.sqlite'),MYDOC_DESKTOP_PROFILE:join(dir,'profile')}});
 try{
  const page=await app.firstWindow();await page.getByRole('button',{name:'版本与更新',exact:true}).click();
  const dialog=page.getByRole('dialog',{name:'版本与更新'}),check=dialog.getByRole('button',{name:'检查服务器版本',exact:true});
  const runtime=await app.evaluate(()=>({arch:process.arch,platform:process.platform}));
  expect(await page.evaluate(()=>({arch:window.mydocDesktop!.arch,platform:window.mydocDesktop!.platform}))).toEqual(runtime);
  await expect(dialog.locator('.client-target')).toContainText(releaseTarget(runtime.platform,runtime.arch));
  await expect(check).toBeDisabled();expect(requests).toHaveLength(0);
  await dialog.getByRole('textbox',{name:'更新服务器地址'}).fill(host);
  await dialog.getByRole('button',{name:'保存地址',exact:true}).click();await expect(check).toBeEnabled();
  await expect(dialog.getByRole('textbox',{name:'更新服务器地址'})).toHaveValue('http://'+host);
  await check.click();await expect(dialog.getByText('发现新版本',{exact:true})).toBeVisible();
  expect(requests[0].pathname).toBe('/api/cloud/releases/latest');
  expect(requests[0].searchParams.get('arch')).toBe(runtime.arch);expect(requests[0].searchParams.get('platform')).toBe(runtime.platform);
  await expect(dialog.locator('.client-release-result h4')).toHaveText(runtime.arch==='arm64'?'0.2.0':'0.3.0');
  await expect(dialog.getByRole('link',{name:'下载安装包'})).toHaveAttribute('href',`http://${host}/api/cloud/releases/download?id=12345678-1234-1234-1234-123456789abc`);
  await mkdir('artifacts',{recursive:true});await page.screenshot({path:'artifacts/client-release-check.png'});
  await dialog.getByRole('button',{name:'关闭',exact:true}).click();
  await page.getByRole('button',{name:'版本与更新',exact:true}).click();await expect(dialog.getByRole('textbox',{name:'更新服务器地址'})).toHaveValue('http://'+host);
  mode='mismatch';await check.click();await expect(dialog.getByRole('alert')).toHaveText('服务器返回的安装包与当前系统或架构不匹配');await expect(dialog.getByRole('link',{name:'下载安装包'})).toHaveCount(0);
  mode='empty';await dialog.getByRole('button',{name:'重新检查',exact:true}).click();await expect(dialog.getByText('暂无匹配的安装包',{exact:true})).toBeVisible();
  mode='error';await check.click();await expect(dialog.getByRole('alert')).toHaveText('更新服务器返回错误（503）');
  await page.evaluate(()=>document.documentElement.dataset.theme='dark');
  await app.evaluate(({BrowserWindow})=>BrowserWindow.getAllWindows()[0].setSize(420,800));
  await page.screenshot({path:'artifacts/client-release-dark-narrow.png'});
  expect(await page.evaluate(()=>document.documentElement.scrollWidth<=innerWidth)).toBe(true);
 }finally{await app.close();await new Promise<void>(resolve=>server.close(()=>resolve()));await rm(dir,{recursive:true,force:true});}
});
