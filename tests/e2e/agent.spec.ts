import {test,expect} from '@playwright/test';
test('Agent floating window: history, drag, settings page and responsive layout',async({page})=>{
 let sent='';
 await page.route('**/api/*agent/**',async route=>{
  const url=new URL(route.request().url());const json=(value:unknown)=>route.fulfill({json:value});
  if(url.pathname.endsWith('/status'))return json({enabled:true,userConfig:true,managed:false,secureStorage:true,disclosure:'提问会发送到你配置的模型地址。'});
  if(url.pathname.endsWith('/sessions'))return json([{id:'s',title:'资料讨论'}]);
  if(url.pathname.endsWith('/subscribe')){
   const rows=[{runId:'r',payload:{type:'message',message:{role:'user',text:'如何整理资料？'}}},{runId:'r',payload:{type:'message',message:{role:'assistant',text:'先按主题建立目录，再补充来源。'}}},{runId:'r',payload:{type:'status',status:'completed'}}];
   return route.fulfill({contentType:'text/event-stream',body:rows.map((event,i)=>'id: '+(i+1)+'\ndata: '+JSON.stringify({cursor:i+1,payload:JSON.stringify(event)})+'\n\n').join('')});
  }
  if(url.pathname.endsWith('/runs')&&route.request().method()==='GET')return json({status:'cancelled'});
  if(url.pathname.endsWith('/runs')){sent=route.request().postDataJSON().prompt;return json({runId:'next',status:'queued'});}
  return json({});
 });
 await page.goto('/');const originalWidth=await page.locator('main').evaluate(el=>el.getBoundingClientRect().width);await page.getByRole('button',{name:'Agent 助手',exact:true}).click();
 const panel=page.getByRole('complementary',{name:'MyDoc Agent'});await expect(panel).toBeVisible();
 await panel.getByLabel('Agent 会话').selectOption('s');await expect(panel.getByText('先按主题建立目录，再补充来源。')).toBeVisible();
 await page.waitForTimeout(2200);await expect(panel.getByText('先按主题建立目录，再补充来源。')).toHaveCount(1);
 await panel.getByLabel('发送给 Agent').fill('继续解释');await panel.getByRole('button',{name:'发送问题'}).click();await expect(panel.getByLabel('发送给 Agent')).toHaveValue('');expect(sent).toBe('继续解释');await panel.getByRole('button',{name:'停止',exact:true}).click();await expect(panel.getByRole('button',{name:'发送问题'})).toBeVisible();
 expect(await page.locator('main').evaluate(el=>el.getBoundingClientRect().width)).toBe(originalWidth);
 const before=await panel.boundingBox();const handle=panel.getByRole('group',{name:'移动 Agent 窗口，使用方向键'});await handle.focus();await page.keyboard.press('ArrowLeft');expect((await panel.boundingBox())!.x).toBeLessThan(before!.x);
 const bounds=await handle.boundingBox();await page.mouse.move(bounds!.x+40,bounds!.y+20);await page.mouse.down();await page.mouse.move(10,20,{steps:5});await page.mouse.up();expect((await panel.boundingBox())!.x).toBeGreaterThanOrEqual(8);
 await expect(panel.getByLabel('接口地址')).toHaveCount(0);
 await panel.getByRole('button',{name:'打开 Agent 设置'}).click();const settings=page.getByRole('dialog',{name:'设置',exact:true});await expect(settings.getByRole('heading',{name:'Agent 设置'})).toBeVisible();await expect(settings.getByLabel('接口地址')).toBeVisible();await settings.getByLabel('接口地址').fill('https://example.com/v1/chat/completions');await settings.getByLabel('模型名称',{exact:true}).fill('test-model');await settings.getByRole('button',{name:'保存模型连接'}).click();await expect(settings.getByText('模型连接已保存')).toBeVisible();await page.screenshot({path:'artifacts/agent-settings-page.png',fullPage:true});await settings.getByRole('button',{name:'关闭弹窗',exact:true}).click();
 await page.screenshot({path:'artifacts/agent-panel-light.png',fullPage:true});
 await page.evaluate(()=>document.documentElement.dataset.theme='dark');await page.screenshot({path:'artifacts/agent-panel-dark.png',fullPage:true});
 await page.setViewportSize({width:390,height:844});await expect(panel).toBeVisible();expect(await panel.evaluate(el=>el.getBoundingClientRect().width)).toBeLessThanOrEqual(390);expect((await panel.boundingBox())!.height).toBeLessThanOrEqual(844);await expect(panel.getByRole('button',{name:'发送问题'})).toBeInViewport({ratio:1});await page.screenshot({path:'artifacts/agent-panel-mobile.png'});
 await panel.getByRole('button',{name:'关闭 Agent'}).click();await expect(panel).toBeHidden();await expect(page.getByRole('button',{name:'Agent 助手',exact:true})).toBeFocused();
});
test('unconfigured model and failed submission preserve user input',async({page})=>{
 let enabled=false;
 await page.route('**/api/*agent/**',async route=>{const path=new URL(route.request().url()).pathname;
  if(path.endsWith('/status'))return route.fulfill({json:{enabled,userConfig:false,managed:true,disclosure:'由管理员配置'}});
  if(path.endsWith('/sessions'))return route.fulfill({json:[{id:'s',title:'问题'}]});
  if(path.endsWith('/subscribe'))return route.fulfill({contentType:'text/event-stream',body:': connected\n\n'});
  return route.fulfill({status:503,json:{error:'模型暂不可用'}});
 });
 await page.goto('/');await page.getByRole('button',{name:'Agent 助手',exact:true}).click();const panel=page.getByRole('complementary',{name:'MyDoc Agent'});await panel.getByLabel('发送给 Agent').fill('保留这段提问');await expect(panel.getByRole('button',{name:'发送问题'})).toBeDisabled();
 enabled=true;await panel.getByRole('button',{name:'关闭 Agent'}).click();await page.getByRole('button',{name:'Agent 助手',exact:true}).click();await panel.getByLabel('Agent 会话').selectOption('s');await panel.getByLabel('发送给 Agent').fill('保留这段提问');await panel.getByRole('button',{name:'发送问题'}).click();await expect(panel.getByRole('alert')).toContainText('模型暂不可用');await expect(panel.getByLabel('发送给 Agent')).toHaveValue('保留这段提问');
});

test('Agent settings is directly reachable through the application settings',async({page})=>{
 await page.route('**/api/*agent/status',route=>route.fulfill({json:{enabled:false,managed:true,userConfig:false,secureStorage:false,disclosure:'由管理员管理'}}));
 await page.goto('/');await page.getByRole('button',{name:'设置',exact:true}).click();const dialog=page.getByRole('dialog',{name:'设置',exact:true});await dialog.getByRole('button',{name:'Agent',exact:true}).click();await expect(dialog.getByRole('heading',{name:'Agent 设置'})).toBeVisible();await expect(dialog.getByRole('heading',{name:'由服务器管理员管理'})).toBeVisible();await expect(dialog.getByLabel('API 密钥')).toHaveCount(0);
});
test('Web local account config works without a team login and survives reload',async({page,context,browser})=>{
 await page.goto('/');await page.getByRole('button',{name:'设置',exact:true}).click();let dialog=page.getByRole('dialog',{name:'设置',exact:true});
 await dialog.getByRole('button',{name:'创建本地账号',exact:true}).click();await dialog.getByLabel('显示名称',{exact:true}).fill('本地独立账号');await dialog.getByLabel('邮箱',{exact:true}).fill('agent-local-'+crypto.randomUUID()+'@example.com');await dialog.getByLabel('密码',{exact:true}).fill('local-account-test-123');await dialog.getByRole('button',{name:'创建并登录',exact:true}).click();await expect(dialog.getByRole('heading',{name:'本地独立账号'})).toBeVisible();
 const cookies=await context.cookies();expect(cookies.some(cookie=>cookie.name==='mydoc_local_session')).toBe(true);expect(cookies.some(cookie=>cookie.name==='mydoc_session')).toBe(false);
 const root=new URL(page.url()).origin,anonymous=await browser.newContext();
 try{
  expect(cookies.find(cookie=>cookie.name==='mydoc_local_session')?.httpOnly).toBe(true);
  const localStatus=await (await page.request.get(root+'/api/local-agent/status')).json();
  const created=await page.request.post(root+'/api/local-agent/sessions',{headers:{'X-Mydoc-Generation':localStatus.generation},data:{title:'仅当前账号可见的历史'}});expect(created.status()).toBe(201);const {id}=await created.json();
  const state=await (await anonymous.request.get(root+'/api/storage/status')).json();expect(state.preferences).not.toHaveProperty('local-account-session');
  expect(JSON.stringify(state)).not.toContain(cookies.find(cookie=>cookie.name==='mydoc_local_session')!.value);
  expect((await anonymous.request.get(root+'/api/local-agent/sessions')).status()).toBe(401);
  expect((await anonymous.request.get(root+'/api/local-agent/events?session='+id)).status()).toBe(401);
 }finally{await anonymous.close();}
 await dialog.getByRole('button',{name:'Agent',exact:true}).click();await expect(dialog.getByLabel('接口地址')).toBeVisible();await expect(dialog.getByLabel('API 类型')).toHaveValue('openai-chat-completions');await dialog.getByLabel('API 类型').selectOption('anthropic-messages');await expect(dialog.getByLabel('接口地址')).toHaveAttribute('placeholder','https://…/v1/messages');await dialog.getByLabel('API 类型').selectOption('openai-responses');await expect(dialog.getByLabel('接口地址')).toHaveAttribute('placeholder','https://…/v1/responses');await dialog.getByLabel('接口地址').fill('https://example.com/v1/responses');await dialog.getByLabel('模型名称',{exact:true}).fill('local-only-model');await dialog.getByRole('button',{name:'保存模型连接'}).click();await expect(dialog.getByText('模型连接已保存')).toBeVisible();
 await page.reload();await page.getByRole('button',{name:'设置',exact:true}).click();dialog=page.getByRole('dialog',{name:'设置',exact:true});await dialog.getByRole('button',{name:'Agent',exact:true}).click();await expect(dialog.getByLabel('模型名称',{exact:true})).toHaveValue('local-only-model');await expect(dialog.getByLabel('API 类型')).toHaveValue('openai-responses');await expect(dialog.getByLabel('接口地址')).toHaveValue('https://example.com/v1/responses');await page.screenshot({path:'artifacts/agent-api-types.png',fullPage:true});
 await dialog.getByRole('button',{name:'添加服务连接'}).click();await dialog.getByLabel('连接名称',{exact:true}).fill('备用服务');await dialog.getByLabel('接口地址').fill('https://backup.example.com/v1/chat/completions');await dialog.getByLabel('模型名称',{exact:true}).fill('backup-model');await dialog.getByLabel('其他模型',{exact:true}).fill('extra-model');await dialog.getByRole('button',{name:'保存模型连接'}).click();await expect(dialog.getByText('模型连接已保存')).toBeVisible();await expect(dialog.getByRole('button',{name:'编辑 备用服务',exact:true})).toBeVisible();
 const option=dialog.getByLabel('默认模型',{exact:true}).locator('option').filter({hasText:'备用服务 / extra-model'});await dialog.getByLabel('默认模型',{exact:true}).selectOption((await option.getAttribute('value'))!);await dialog.getByRole('button',{name:'保存默认模型'}).click();await expect(dialog.getByText('默认模型已更新')).toBeVisible();
 await page.reload();await page.getByRole('button',{name:'设置',exact:true}).click();dialog=page.getByRole('dialog',{name:'设置',exact:true});await dialog.getByRole('button',{name:'Agent',exact:true}).click();await expect(dialog.getByLabel('默认模型',{exact:true}).locator('option:checked')).toContainText('备用服务 / extra-model');await expect(dialog.getByRole('button',{name:'删除 备用服务',exact:true})).toBeDisabled();
 await dialog.getByRole('button',{name:'删除 模型服务',exact:true}).click();await dialog.getByRole('button',{name:'确认删除',exact:true}).click();await expect(dialog.getByRole('button',{name:'编辑 模型服务',exact:true})).toHaveCount(0);await page.screenshot({path:'artifacts/agent-multi-models.png',fullPage:true});
 await page.setViewportSize({width:390,height:844});await expect(dialog).toBeVisible();expect(await dialog.evaluate(el=>el.scrollWidth<=el.clientWidth+1)).toBe(true);await page.screenshot({path:'artifacts/agent-multi-models-mobile.png'});await page.setViewportSize({width:1280,height:720});

 await dialog.getByRole('button',{name:'本地账号',exact:true}).click();await dialog.getByRole('button',{name:'退出本地账号'}).click();await dialog.getByRole('button',{name:'Agent',exact:true}).click();await expect(dialog.getByRole('alert')).toContainText('登录本地账号');
});
test('launcher drag is bounded, persists without opening chat, and hidden state is reversible',async({page})=>{
 await page.route('**/api/*agent/status',route=>route.fulfill({json:{enabled:false,userConfig:false,managed:true,disclosure:'由管理员配置'}}));
 await page.goto('/');const launcher=page.getByRole('button',{name:'Agent 助手',exact:true});await expect(launcher).toBeVisible();
 const before=(await launcher.boundingBox())!;const saved=page.waitForResponse(response=>response.url().includes('/api/storage/preference?key=mydoc%3Aagent-launcher-position')&&response.ok());
 await page.mouse.move(before.x+25,before.y+24);await page.mouse.down();await page.mouse.move(130,170,{steps:8});await page.mouse.up();await saved;
 await expect(page.getByRole('complementary',{name:'MyDoc Agent'})).toHaveCount(0);const moved=(await launcher.boundingBox())!;expect(moved.x).toBeLessThan(before.x);
 await page.reload();await expect(launcher).toBeVisible();expect(Math.abs((await launcher.boundingBox())!.x-moved.x)).toBeLessThan(2);
 await launcher.focus();await page.keyboard.press('ArrowLeft');expect((await launcher.boundingBox())!.x).toBeLessThan(moved.x);await page.keyboard.press('Enter');await expect(page.getByRole('complementary',{name:'MyDoc Agent'})).toBeVisible();await page.getByRole('button',{name:'关闭 Agent',exact:true}).click();
 await page.getByRole('button',{name:'设置',exact:true}).click();let dialog=page.getByRole('dialog',{name:'设置',exact:true});await dialog.getByRole('button',{name:'Agent',exact:true}).click();await dialog.getByLabel('显示悬浮图标').uncheck();await expect(dialog.getByLabel('显示悬浮图标')).toBeEnabled();await dialog.getByRole('button',{name:'关闭弹窗'}).click();await expect(launcher).toHaveCount(0);
 await page.reload();await expect(page.getByRole('button',{name:'搜索文档',exact:true})).toBeEnabled();await page.keyboard.press('ControlOrMeta+k');const palette=page.getByRole('dialog',{name:'统一命令面板'});await palette.getByRole('searchbox').fill('> 打开 Agent 助手');await palette.getByRole('option').filter({has:page.getByText('打开 Agent 助手',{exact:true})}).click();await expect(page.getByRole('complementary',{name:'MyDoc Agent'})).toBeVisible();await page.getByRole('button',{name:'关闭 Agent',exact:true}).click();await expect(launcher).toHaveCount(0);
 await page.getByRole('button',{name:'设置',exact:true}).click();dialog=page.getByRole('dialog',{name:'设置',exact:true});await dialog.getByRole('button',{name:'Agent',exact:true}).click();await expect(dialog.getByLabel('显示悬浮图标')).not.toBeChecked();await dialog.getByLabel('显示悬浮图标').check();await expect(dialog.getByLabel('显示悬浮图标')).toBeEnabled();await page.screenshot({path:'artifacts/agent-launcher-setting.png',fullPage:true});await dialog.getByRole('button',{name:'关闭弹窗'}).click();await expect(launcher).toBeVisible();
 await page.setViewportSize({width:390,height:844});await expect(launcher).toBeInViewport({ratio:1});await page.screenshot({path:'artifacts/agent-launcher-mobile.png'});
});
