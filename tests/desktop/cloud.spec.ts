import {test,expect,_electron as electron,type ElectronApplication,type Page} from '@playwright/test';
import {createServer} from 'node:http';
import {mkdtemp,rm} from 'node:fs/promises';
import {tmpdir,networkInterfaces} from 'node:os';
import {join,resolve} from 'node:path';
import {CloudServer} from '../../packages/frontend/apps/web/server/cloud';

async function openCloud(page:Page){
  await page.getByLabel('工作区选项').click();
  await page.getByRole('button',{name:'团队工作空间',exact:true}).click();
  return page.getByRole('dialog',{name:'团队工作空间',exact:true});
}

test('桌面端通过 HTTP IP 登录，重新载入与进程重启保持会话，退出后恢复登录入口',async()=>{
  const directory=await mkdtemp(join(tmpdir(),'mydoc-desktop-cloud-'));
  const cloud=new CloudServer(join(directory,'cloud.sqlite'));
  const server=createServer((req,res)=>{req.url=req.url?.replace(/^\/api\/cloud/,'');void cloud.handler(req,res);});
  await new Promise<void>(done=>server.listen(0,'0.0.0.0',done));
  const ip=Object.values(networkInterfaces()).flat().find(address=>address?.family==='IPv4'&&!address.internal)?.address??'127.0.0.2';
  const root='http://'+ip+':'+(server.address() as {port:number}).port;
  const launch=()=>electron.launch({executablePath:process.env.MYDOC_TEST_ELECTRON_PATH,args:[resolve('packages/frontend/apps/desktop')],env:{...process.env,MYDOC_DESKTOP_DB_PATH:join(directory,'data','mydoc.sqlite'),MYDOC_DESKTOP_PROFILE:join(directory,'profile')},timeout:30000});
  let app:ElectronApplication|undefined;
  try{
    const registration=await fetch(root+'/api/cloud/register',{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify({email:'desktop-ip@example.com',name:'IP 登录账户',password:'desktop-password-1'})});
    expect(registration.ok).toBe(true);
    const cookie=registration.headers.get('set-cookie')!.split(';')[0];
    await fetch(root+'/api/cloud/workspaces',{method:'POST',headers:{'Content-Type':'application/json',Cookie:cookie},body:JSON.stringify({name:'IP 团队空间'})});
    app=await launch();let page=await app.firstWindow();
    await expect(page.getByRole('button',{name:'新建文档',exact:true}).first()).toBeVisible();
    let dialog=await openCloud(page);
    await dialog.getByLabel('服务器地址').fill(root);
    await dialog.getByRole('button',{name:'添加服务器',exact:true}).click();
    await dialog.getByRole('listitem').filter({hasText:root}).getByRole('button',{name:'连接',exact:true}).click();
    await expect(dialog).toBeVisible();
    await expect(dialog.getByText('当前应用服务器',{exact:true})).toHaveCount(0);
    await expect(dialog.getByRole('heading',{name:'团队工作空间',exact:true}).last()).toBeVisible();
    await expect(dialog.getByRole('button',{name:'登录',exact:true})).toBeEnabled();
    await dialog.getByLabel('邮箱',{exact:true}).fill('desktop-ip@example.com');
    await dialog.getByLabel('密码',{exact:true}).fill('desktop-password-1');
    await dialog.getByRole('button',{name:'登录',exact:true}).click();
    await expect(dialog).toBeVisible();
    await expect(dialog.getByText('IP 团队空间',{exact:true})).toBeVisible();
    await dialog.getByRole('button',{name:'关闭弹窗',exact:true}).click();
    await page.getByRole('button',{name:'设置',exact:true}).click();
    const settings=page.getByRole('dialog',{name:'设置',exact:true});
    await settings.getByRole('button',{name:'服务器账号',exact:true}).click();
    await expect(settings.locator('dd').filter({hasText:'desktop-ip@example.com'})).toBeVisible();
    await expect(settings.getByRole('button',{name:'退出服务器账号',exact:true})).toBeVisible();
    await settings.getByRole('button',{name:'关闭弹窗',exact:true}).click();
    dialog=await openCloud(page);
    await expect(dialog.getByText('IP 团队空间',{exact:true})).toBeVisible();
    const sessionCookies=await app.context().cookies();
    const session=sessionCookies.find(value=>value.name.startsWith('mydoc-cloud-'))!;
    expect(session.httpOnly).toBe(true);expect(session.secure).toBe(false);expect(session.sameSite).toBe('Strict');
    await dialog.getByRole('button',{name:'打开',exact:true}).click();
    await expect(dialog).toHaveCount(0);
    await page.getByRole('button',{name:'新建文档',exact:true}).first().click();
    await page.getByLabel('文档标题',{exact:true}).fill('IP 同步文档');
    await page.getByRole('textbox',{name:'文档正文'}).fill('桌面 IP 连接后的同步正文');
    await page.getByLabel('添加本地附件').setInputFiles({name:'IP附件.txt',mimeType:'text/plain',buffer:Buffer.from('IP附件字节')});
    await expect(page.getByRole('button',{name:'下载 IP附件.txt'})).toBeVisible();
    await expect(page.getByText('正在保存…',{exact:true})).toHaveCount(0);
    await app.close();app=await launch();page=await app.firstWindow();
    await expect(page.getByRole('button',{name:'新建文档',exact:true}).first()).toBeVisible();
    await expect(page.getByLabel('文档标题',{exact:true})).toHaveValue('IP 同步文档');
    await expect(page.getByRole('textbox',{name:'文档正文'})).toContainText('桌面 IP 连接后的同步正文');
    await expect(page.getByRole('button',{name:'下载 IP附件.txt'})).toBeVisible();
    await page.getByLabel('当前工作区',{exact:true}).selectOption('local');
    dialog=await openCloud(page);
    await expect(dialog.getByText('IP 团队空间',{exact:true})).toBeVisible();
    await dialog.getByRole('button',{name:'关闭弹窗',exact:true}).click();
    await page.getByRole('button',{name:'设置',exact:true}).click();
    await page.getByRole('dialog',{name:'设置',exact:true}).getByRole('button',{name:'服务器账号',exact:true}).click();
    await expect(page.getByRole('dialog',{name:'设置',exact:true}).locator('dd').filter({hasText:'desktop-ip@example.com'})).toBeVisible();
    await page.getByRole('dialog',{name:'设置',exact:true}).getByRole('button',{name:'退出服务器账号',exact:true}).click();
    await expect(page.getByRole('dialog',{name:'设置',exact:true})).toHaveCount(0);
    dialog=await openCloud(page);
    await expect(dialog.getByRole('button',{name:'登录',exact:true})).toBeVisible();
  }finally{
    await app?.close();server.closeAllConnections();await new Promise<void>(done=>server.close(()=>done()));cloud.close();await rm(directory,{recursive:true,force:true});
  }
});
