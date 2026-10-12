import {test,expect} from '@playwright/test';
const root='http://127.0.0.1:4180';
const png=Buffer.concat([Buffer.from([0x89,0x50,0x4e,0x47,0x0d,0x0a,0x1a,0x0a]),Buffer.alloc(64,9)]);
test('账户设置：改名、头像、邮箱验证、改密与登录设备注销',async({browser})=>{
 const suffix=Date.now(),email=`account-${suffix}@example.com`;
 const first=await browser.newContext(),second=await browser.newContext();
 try{
  const page=await first.newPage();
  expect((await page.request.post(root+'/api/cloud/register',{data:{email,name:'账户测试',password:'account-password-1'}})).ok()).toBe(true);
  // 建一个团队工作空间：会话失效提示只在服务端同步的工作空间里出现。
  const space=await (await page.request.post(root+'/api/cloud/workspaces',{data:{name:'账户测试空间'}})).json() as {id:string};
  // 第二个浏览器上下文登录同一账户，用来验证“登录设备”列表与注销。
  const other=await second.newPage();
  expect((await other.request.post(root+'/api/cloud/login',{data:{email,password:'account-password-1'}})).ok()).toBe(true);
  await other.goto(root);await expect(other.getByRole('button',{name:'新建文档',exact:true}).first()).toBeVisible();
  await other.getByLabel('当前工作区',{exact:true}).selectOption(space.id);
  await other.getByRole('button',{name:'新建文档',exact:true}).first().click();
  await other.getByLabel('文档标题',{exact:true}).fill('会话失效前的文档');
  await expect(other.getByRole('status').filter({hasText:'已同步到工作空间'}).first()).toBeVisible();

  await page.goto(root);
  await page.getByLabel('工作区选项',{exact:true}).click();
  await page.getByRole('button',{name:'团队工作空间',exact:true}).click();
  const dialog=page.getByRole('dialog',{name:'团队工作空间'});
  await dialog.getByRole('button',{name:'账户设置'}).click();
  // 未配置邮件服务时如实说明口令的获取方式。
  await expect(dialog.getByText(/邮件服务未配置/)).toBeVisible();
  await expect(dialog.getByLabel('账户概览')).toContainText('未验证');

  await dialog.getByLabel('显示名称').fill('改过的名字');
  await dialog.getByRole('button',{name:'保存显示名称'}).click();
  await expect(dialog.getByText('显示名称已更新')).toBeVisible();
  await expect(dialog.getByText('改过的名字 · '+email)).toBeVisible();

  await dialog.getByLabel(/上传头像/).setInputFiles({name:'avatar.png',mimeType:'image/png',buffer:png});
  await expect(dialog.getByText('头像已更新')).toBeVisible();
  await expect(dialog.getByRole('img',{name:'改过的名字 的头像'})).toBeVisible();

  // 邮箱验证：未配置邮件服务时口令直接显示，填入后完成验证。
  await dialog.getByRole('button',{name:'发送验证口令'}).click();
  const token=await dialog.getByLabel('验证口令（仅本次显示）').inputValue();
  expect(token).toMatch(/^[0-9a-f]{48}$/);
  await dialog.getByLabel('输入收到的验证口令').fill(token);
  await dialog.getByRole('button',{name:'完成验证'}).click();
  await expect(dialog.getByText('邮箱已验证。')).toBeVisible();
  await expect(dialog.getByLabel('账户概览')).toContainText('已验证');

  // 登录设备：两个上下文各一条，注销另一台后对方失去访问权限。
  await expect(dialog.getByLabel('登录设备列表').getByRole('listitem')).toHaveCount(2);
  await dialog.getByRole('button',{name:/^注销/}).first().click();
  await dialog.getByRole('button',{name:'确认注销设备'}).click();
  await expect(dialog.getByLabel('登录设备列表').getByRole('listitem')).toHaveCount(1);
  // 被注销的设备在 SSE 断开后察觉会话失效：暂停编辑，但页面内容仍保留且可导出草稿。
  await expect(other.getByText('登录已过期。页面内容与已暂存的草稿仍保留，重新登录后会继续提交。')).toBeVisible({timeout:30000});
  await expect(other.getByRole('button',{name:'重新登录'})).toBeVisible();
  await expect(other.getByLabel('文档标题',{exact:true})).toBeDisabled();
  await expect(other.getByLabel('文档标题',{exact:true})).toHaveValue('会话失效前的文档');
  const draftDownload=other.waitForEvent('download');
  await other.getByRole('button',{name:'导出草稿',exact:true}).click();
  expect((await draftDownload).suggestedFilename()).toContain('mydoc-draft.json');

  // 改密：当前密码错误时拒绝，正确时当前会话继续可用。
  await dialog.getByLabel('当前密码',{exact:true}).fill('wrong-password-1');
  await dialog.getByLabel('新密码（至少 10 位）').fill('rotated-password-1');
  await dialog.getByRole('button',{name:'保存新密码'}).click();
  await expect(dialog.getByText('当前密码不正确')).toBeVisible();
  await dialog.getByLabel('当前密码',{exact:true}).fill('account-password-1');
  await dialog.getByLabel('新密码（至少 10 位）').fill('rotated-password-1');
  await dialog.getByRole('button',{name:'保存新密码'}).click();
  await expect(dialog.getByText(/密码已更新/)).toBeVisible();
  await page.screenshot({path:'artifacts/account-settings.png'});
  // 新密码生效、旧密码失效。
  expect((await page.request.post(root+'/api/cloud/login',{data:{email,password:'account-password-1'}})).status()).toBe(401);
  expect((await page.request.post(root+'/api/cloud/login',{data:{email,password:'rotated-password-1'}})).ok()).toBe(true);

  // 找回密码入口不暴露邮箱是否存在。
  await page.goto(root+'/?reset-check=1');
  await page.getByLabel('工作区选项',{exact:true}).click();
  await page.getByRole('button',{name:'团队工作空间',exact:true}).click();
  // 退出登录后独立服务显示登录页，找回密码入口就在这里。
  await page.getByRole('dialog',{name:'团队工作空间'}).getByRole('button',{name:'退出登录'}).click();
  await expect.poll(async()=>(await (await page.request.get(root+'/api/cloud/session')).json()).user,{timeout:15000}).toBe(null);
  await page.goto(root);
  await expect(page.getByRole('heading',{name:'MyDoc 工作空间'})).toBeVisible();
  await page.getByRole('button',{name:'忘记密码'}).click();
  // 不存在的邮箱也返回相同提示，不暴露账户是否存在。
  await page.getByLabel('账户邮箱').fill('definitely-missing-'+suffix+'@example.com');
  await page.getByRole('button',{name:'获取重置口令'}).click();
  await expect(page.getByText(/邮件服务未配置|重置口令已发送/)).toBeVisible();
  await expect(page.getByLabel('重置口令')).toBeVisible();
  // 无效口令被拒绝，不会放行。
  await page.getByLabel('重置口令').fill('0'.repeat(48));
  await page.getByLabel('新密码（至少 10 位）').fill('should-not-apply-1');
  await page.getByRole('button',{name:'重置密码'}).click();
  await expect(page.getByRole('alert')).toContainText('口令无效或已被使用');
  await page.screenshot({path:'artifacts/account-password-reset.png'});
 }finally{await first.close();await second.close();}
});
