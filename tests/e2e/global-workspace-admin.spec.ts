import {test,expect} from '@playwright/test';

test('全局空间管理、成员与所有权操作、删除审计及移动端',async({browser})=>{
 const root='http://127.0.0.1:4182',admin=await browser.newContext(),owner=await browser.newContext();
 try{
  expect((await admin.request.post(root+'/api/cloud/login',{data:{email:'admin@example.com',password:'admin-test-password-123'}})).ok()).toBe(true);
  expect((await owner.request.post(root+'/api/cloud/register',{data:{email:'owner@example.com',name:'原所有者',password:'owner-password-123'}})).ok()).toBe(true);
  const space=await(await owner.request.post(root+'/api/cloud/workspaces',{data:{name:'全局管理验证空间'}})).json();
  const outsider=await browser.newContext();await outsider.request.post(root+'/api/cloud/register',{data:{email:'new-owner@example.com',name:'新成员',password:'member-password-123'}});await outsider.close();
  const a=await admin.newPage();await a.goto(root);await a.getByRole('button',{name:'设置',exact:true}).click();
  const settings=a.getByRole('dialog',{name:'设置',exact:true});await settings.getByRole('button',{name:'所有团队空间',exact:true}).click();
  await expect(settings.getByRole('list',{name:'所有团队空间'})).toContainText('全局管理验证空间');
  await settings.getByRole('button',{name:'管理 全局管理验证空间',exact:true}).click();
  await settings.getByRole('textbox',{name:'空间名称',exact:true}).fill('全局改名后的空间');await settings.getByRole('button',{name:'保存名称'}).click();
  await expect(settings.getByRole('heading',{name:'全局改名后的空间',exact:true})).toBeVisible();
  await settings.getByRole('textbox',{name:'搜索服务器用户'}).fill('new-owner@example.com');await settings.getByRole('button',{name:'搜索用户',exact:true}).click();
  await settings.getByRole('button',{name:'添加为成员',exact:true}).click();let confirmation=settings.getByRole('alertdialog',{name:'确认全局空间操作'});await confirmation.getByRole('button',{name:'确认执行'}).click();
  const members=settings.getByRole('list',{name:'全局管理空间成员'});await expect(members).toContainText('新成员');
  await members.getByRole('combobox',{name:'新成员的角色'}).selectOption('admin');await settings.getByRole('alertdialog').getByRole('button',{name:'确认执行'}).click();
  await expect(members.getByRole('combobox',{name:'新成员的角色'})).toHaveValue('admin');
  await members.getByRole('button',{name:'设为所有者',exact:true}).click();confirmation=settings.getByRole('alertdialog');
  await expect(confirmation.getByRole('button',{name:'确认执行'})).toBeDisabled();await confirmation.getByRole('textbox').fill('全局改名后的空间');await confirmation.getByRole('button',{name:'确认执行'}).click();
  await expect(members.getByRole('combobox',{name:'原所有者的角色'})).toHaveValue('admin');
  await members.getByRole('button',{name:'移除',exact:true}).click();await settings.getByRole('alertdialog').getByRole('button',{name:'确认执行'}).click();await expect(members).not.toContainText('原所有者');
  await a.screenshot({path:'artifacts/global-space-admin-desktop.png'});await a.setViewportSize({width:390,height:844});await a.screenshot({path:'artifacts/global-space-admin-mobile.png'});
  expect(await a.evaluate(()=>document.documentElement.scrollWidth<=document.documentElement.clientWidth)).toBe(true);
  expect((await admin.request.get(root+'/api/cloud/document?workspace='+space.id+'&id=doc:test')).status()).toBe(403);
  await settings.getByRole('button',{name:'删除工作空间',exact:true}).click();confirmation=settings.getByRole('alertdialog');await confirmation.getByRole('textbox').fill('全局改名后的空间');await confirmation.getByRole('button',{name:'确认执行'}).click();
  await expect(settings.getByRole('heading',{name:'所有团队空间',exact:true})).toBeVisible();await expect(settings.getByText('永久删除空间：全局改名后的空间')).toBeVisible();
  const b=await owner.newPage();await b.goto(root);await b.getByRole('button',{name:'设置',exact:true}).click();await expect(b.getByRole('button',{name:'所有团队空间',exact:true})).toHaveCount(0);
 }finally{await admin.close();await owner.close();}
});
