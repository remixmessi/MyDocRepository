import {test,expect,type Page} from '@playwright/test';
import {mergeUpdates} from 'yjs';
import {SCHEMA_VERSION} from '../../packages/common/nbstore/src/schema';

// Use isolated in-memory storage so desktop startup can be exercised in a browser
// without touching the user's SQLite database or requiring an Electron build.
async function fixture(page:Page,desktop:boolean,unavailable=false,admin=true){
 const documents=new Map<string,Uint8Array>();
 await page.route('**/api/hosting',route=>route.fulfill({json:{hosted:false}}));
 await page.route('**/api/storage/**',async route=>{
  const request=route.request(),url=new URL(request.url()),path=url.pathname.slice('/api/storage'.length);
  if(path==='/status')return route.fulfill({json:{initialized:true,schemaVersion:SCHEMA_VERSION,generation:'test',preferences:{'cloud-server-url':'https://team.example','cloud-servers':JSON.stringify([{url:'https://team.example',name:'team.example'}])}}});
  if(path==='/document'){
   const id=url.searchParams.get('id')!;
   if(request.method()==='POST'){
    const update=new Uint8Array(request.postDataBuffer()!),previous=documents.get(id);
    documents.set(id,previous?mergeUpdates([previous,update]):update);
    return route.fulfill({json:{}});
   }
   const data=documents.get(id);
   return data?route.fulfill({body:Buffer.from(data),contentType:'application/octet-stream'}):route.fulfill({status:204});
  }
  if(path==='/sync-queue')return route.fulfill({json:{generation:'test',rows:[],lastSuccess:0}});
  return route.fulfill({json:{}});
 });
 await page.route('**/api/cloud/**',route=>{
  const pathname=new URL(route.request().url()).pathname,path=pathname.slice(pathname.lastIndexOf('/api/cloud/')+'/api/cloud'.length);
  if(unavailable)return route.fulfill({status:503,json:{error:'测试服务器暂不可用'}});
  if(path==='/session')return route.fulfill({json:{user:{id:'member',name:'团队成员',email:'member@example.com',serverAdmin:admin}}});
  if(path==='/workspaces')return route.fulfill({json:[{id:'joined-team',name:'已加入的团队空间',role:'member',createdAt:1},{id:'other-team',name:'另一个团队空间',role:'owner',createdAt:2}]});
  if(path==='/admin/users')return route.fulfill({json:{users:[{id:'test-user',name:new URL(route.request().url()).searchParams.get('workspace')==='joined-team'?'所选空间用户':'另一个空间用户',email:'user@example.com',verified:1,createdAt:1,serverAdmin:0,defaultAdmin:0}],total:1,offset:0,limit:100}});
  if(path==='/document')return route.fulfill({status:204});
  if(path==='/releases/latest')return route.fulfill({json:null});
  return route.fulfill({json:[]});
 });
 if(desktop)await page.addInitScript(()=>{
  const idle=()=>()=>{};
  window.mydocDesktop={checkServerUpdates:async server=>{(window as Window&{serverUpdateAddress?:string}).serverUpdateAddress=server;return {version:'1.0.0',source:'server',serverUrl:server,enabled:true,phase:'up-to-date',automaticChecks:false,channel:'测试'};},platform:'darwin',arch:'arm64',
   storage:async request=>{
    const response=await fetch('/api/storage'+request.path,{method:request.method,body:request.body?new Uint8Array(request.body):undefined});
    return {status:response.status,contentType:response.headers.get('Content-Type')??'',body:new Uint8Array(await response.arrayBuffer())};
   },updateStatus:async()=>({version:'1.0.0',enabled:false,phase:'disabled',automaticChecks:false,channel:'test'}),updateAction:async()=>{throw new Error('Disabled in fixture');},setAutomaticUpdateChecks:async()=>{throw new Error('Disabled in fixture');},onUpdate:idle,onShowUpdates:idle,
   chooseBackupDirectory:async()=>undefined,openBackupDirectory:async()=>{},acknowledgeDocumentLink:async()=>{},acknowledgeOpenFiles:async()=>{},
   onPrepareClose:idle,onDocumentLink:idle,onOpenFiles:idle,takeDocumentLink:async()=>undefined,takeOpenFiles:async()=>undefined,
  };
 });
}

for(const desktop of [false,true])test(`${desktop?'桌面':'Web'}连接后自动发现已加入的空间，刷新后仍可选择`,async({page})=>{
 await fixture(page,desktop);
 await page.goto('/');
 const picker=page.getByLabel('当前工作区',{exact:true});
 await expect(picker.locator('option[value="joined-team"]')).toHaveText('☁ 已加入的团队空间');
 await picker.selectOption('joined-team');
 await expect(picker).toHaveValue('joined-team');
 await page.reload();
 await expect(picker.locator('option[value="joined-team"]')).toHaveText('☁ 已加入的团队空间');
 await expect(picker).toHaveValue('joined-team');
});

test('团队服务器不可用时保留本地工作区并显示加载失败提示',async({page})=>{
 await fixture(page,true,true);
 await page.goto('/');
 await expect(page.getByLabel('当前工作区',{exact:true})).toHaveValue('local');
 await expect(page.getByRole('status').filter({hasText:'团队工作空间加载失败：测试服务器暂不可用'})).toBeVisible();
});

for(const selected of [false,true])test(`用户管理${selected?'默认选择当前团队空间':'在本地工作区时先选择团队空间'}`,async({page})=>{
 await fixture(page,true);
 await page.goto('/');
 const picker=page.getByLabel('当前工作区',{exact:true});
 await expect(picker.locator('option[value="joined-team"]')).toHaveCount(1);
 if(selected)await picker.selectOption('joined-team');
 await page.getByRole('button',{name:'设置',exact:true}).click();
 await page.getByRole('button',{name:'用户管理',exact:true}).click();
 const selector=page.getByLabel('管理工作空间',{exact:true});
 await expect(selector).toHaveValue(selected?'joined-team':'');
 if(!selected){await expect(page.getByLabel('服务器用户列表')).toHaveCount(0);await selector.selectOption('joined-team');}
 await expect(page.getByLabel('服务器用户列表')).toContainText('所选空间用户');
 await selector.selectOption('other-team');
 await expect(page.getByLabel('服务器用户列表')).toContainText('另一个空间用户');
 await expect(page.getByLabel('服务器用户列表')).not.toContainText('所选空间用户');
});

test('已登录的普通账号显示权限提示，不要求重新连接或登录',async({page})=>{
 await fixture(page,true,false,false);await page.goto('/');
 await page.getByRole('button',{name:'设置',exact:true}).click();await page.getByRole('button',{name:'用户管理',exact:true}).click();
 await expect(page.getByRole('status').filter({hasText:'已登录 member@example.com，当前账号没有服务器用户管理权限'})).toBeVisible();
 await expect(page.getByRole('button',{name:'登录服务器账号',exact:true})).toHaveCount(0);
});

test('会话读取失败显示错误和重试，不伪装成未连接',async({page})=>{
 await fixture(page,true,true);await page.goto('/');
 await page.getByRole('button',{name:'设置',exact:true}).click();await page.getByRole('button',{name:'用户管理',exact:true}).click();
 await expect(page.getByRole('alert').filter({hasText:'账号读取失败：测试服务器暂不可用'})).toBeVisible();
 await expect(page.getByRole('button',{name:'重试读取账号',exact:true})).toBeVisible();
 await expect(page.getByRole('button',{name:'登录服务器账号',exact:true})).toHaveCount(0);
});

test('Web版本维护直接使用当前Web服务器，无连接地址表单',async({page})=>{
 await fixture(page,false);await page.goto('/');
 await page.getByRole('button',{name:'设置',exact:true}).click();
 const request=page.waitForRequest(value=>new URL(value.url()).pathname==='/api/cloud/admin/releases');
 await page.getByRole('button',{name:'版本维护',exact:true}).click();await request;
 await expect(page.getByRole('button',{name:'发布版本',exact:true})).toBeVisible();
 await expect(page.getByLabel('版本服务器地址',{exact:true})).toHaveCount(0);
 await expect(page.getByLabel('更新服务器地址',{exact:true})).toHaveCount(0);
});

test('桌面版本与更新使用独立更新地址，不提供版本发布管理',async({page})=>{
 await fixture(page,true);await page.goto('/');
 await page.getByRole('button',{name:'设置',exact:true}).click();await page.getByRole('button',{name:'版本与更新',exact:true}).first().click();
 await expect(page.getByLabel('更新服务器地址',{exact:true})).toHaveValue('');
 await page.getByLabel('更新服务器地址',{exact:true}).fill('https://releases.example');
 await page.getByRole('button',{name:'保存地址',exact:true}).click();
 await page.getByRole('button',{name:'检查服务器版本',exact:true}).click();
 await expect.poll(()=>page.evaluate(()=>(window as Window&{serverUpdateAddress?:string}).serverUpdateAddress)).toBe('https://releases.example');
 await expect(page.getByText('当前已是最新版本',{exact:true})).toBeVisible();
 await expect(page.getByRole('button',{name:'发布版本',exact:true})).toHaveCount(0);
 await page.getByRole('button',{name:'用户管理',exact:true}).click();
 await expect(page.getByText('当前团队服务器：https://team.example',{exact:true})).toBeVisible();
 await expect(page.getByLabel('管理工作空间',{exact:true})).toBeVisible();
});
