import {test,expect} from '@playwright/test';
import {newNote} from './fixtures';

test('文档权限管理、在线降权和恢复默认权限',async({browser})=>{
 const root=process.env.MYDOC_PERMISSION_TEST_URL??'http://127.0.0.1:4181';
 const first=await browser.newContext(),second=await browser.newContext(),a=await first.newPage(),b=await second.newPage();
 try{
  const suffix=Date.now(),email='member'+suffix+'@example.com';
  for(const [context,name,address] of [[first,'所有者','owner'+suffix+'@example.com'],[second,'普通成员',email]] as const){
   expect((await context.request.post(root+'/api/cloud/register',{data:{name,email:address,password:'test-password-12345'}})).ok()).toBe(true);
  }
  const workspace=await (await first.request.post(root+'/api/cloud/workspaces',{data:{name:'权限验证空间'}})).json();
  const invite=await (await first.request.post(root+'/api/cloud/invite?workspace='+workspace.id,{data:{email}})).json();
  expect((await second.request.post(root+'/api/cloud/accept',{data:{token:invite.token}})).ok()).toBe(true);
  async function openTeam(page:typeof a){await page.goto(root);await page.getByLabel('工作区选项',{exact:true}).click();await page.getByRole('button',{name:'团队工作空间',exact:true}).click();await page.getByRole('dialog',{name:'团队工作空间'}).getByRole('button',{name:'打开',exact:true}).click();}
  await openTeam(a);
  await newNote(a);await a.getByLabel('文档标题',{exact:true}).fill('授权文档');await a.getByRole('textbox',{name:'文档正文'}).fill('权限验证正文');
  await expect(a.getByRole('status').filter({hasText:'正在保存'})).toHaveCount(0);
  await newNote(a);await a.getByLabel('文档标题',{exact:true}).fill('自由文档');
  await openTeam(b);await b.getByRole('button',{name:'授权文档',exact:true}).first().click();
  await expect(b.getByRole('textbox',{name:'文档正文'})).toContainText('权限验证正文');
  await expect(b.getByLabel('文档标题',{exact:true})).toBeEnabled();await expect(b.getByRole('button',{name:'管理文档权限',exact:true})).toHaveCount(0);
  await a.getByRole('button',{name:'授权文档',exact:true}).first().click();await a.getByRole('button',{name:'管理文档权限',exact:true}).click();
  const dialog=a.getByRole('dialog',{name:'文档权限 · 授权文档'});
  await expect(dialog.getByText('所有成员使用空间默认权限。')).toBeVisible();
  await dialog.getByRole('combobox',{name:'成员',exact:true}).selectOption({label:'普通成员 · '+email});await dialog.getByRole('combobox',{name:'权限',exact:true}).selectOption('viewer');await dialog.getByRole('button',{name:'保存文档权限'}).click();
  await expect(dialog.getByText('文档权限已更新')).toBeVisible();await expect(b.getByLabel('文档标题',{exact:true})).toBeDisabled();await expect(b.getByRole('textbox',{name:'文档正文'})).toHaveAttribute('contenteditable','false');
  await b.getByRole('button',{name:'自由文档',exact:true}).first().click();await expect(b.getByLabel('文档标题',{exact:true})).toBeEnabled();
  await b.getByRole('textbox',{name:'文档正文'}).fill('仍可编辑');
  await a.screenshot({path:'artifacts/document-permissions-desktop.png'});
  await a.setViewportSize({width:390,height:844});await a.screenshot({path:'artifacts/document-permissions-mobile.png'});
  expect(await a.evaluate(()=>document.documentElement.scrollWidth<=document.documentElement.clientWidth)).toBe(true);
  await a.setViewportSize({width:1280,height:720});
  a.once('dialog',value=>value.accept());await dialog.getByRole('button',{name:/恢复默认/}).click();await expect(dialog.getByText('已恢复空间默认权限',{exact:true})).toBeVisible();
  await b.getByRole('button',{name:'授权文档',exact:true}).first().click();await expect(b.getByLabel('文档标题',{exact:true})).toBeEnabled();
  await b.getByRole('textbox',{name:'文档正文'}).fill('恢复后编辑');await dialog.getByRole('button',{name:'完成',exact:true}).click();await expect(a.getByRole('textbox',{name:'文档正文'})).toContainText('恢复后编辑');
 }finally{await first.close();await second.close();}
});
