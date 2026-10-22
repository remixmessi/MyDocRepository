import {createRequire} from 'node:module';
import {spawnSync} from 'node:child_process';
import {randomUUID} from 'node:crypto';
import {setTimeout as delay} from 'node:timers/promises';
const require=createRequire(process.cwd()+'/package.json'),{chromium,expect}=require('@playwright/test');
const id=randomUUID().slice(0,8),name='mydoc-ui-'+id,volume=name+'-data',rollback=name+'-rollback',password=randomUUID();let origin,port,browser;
const docker=args=>{const r=spawnSync('docker',args,{encoding:'utf8',maxBuffer:2e6});if(r.status)throw Error(r.stderr||r.stdout);return r.stdout.trim();};
// Only the first start on a fresh volume may initialize; later starts must reuse the existing database.
async function start(image,mounted=volume,init=false){docker(['run','-d','--name',name,'--platform','linux/amd64','-p','127.0.0.1:'+(port||'')+':8080','-v',mounted+':/app/data',...(init?['-e','MYDOC_INIT_DATABASE=1']:[]),'-e','MYDOC_ADMIN_EMAIL=upgrade@example.com','-e','MYDOC_ADMIN_PASSWORD='+password,image]);origin='http://'+docker(['port',name,'8080/tcp']);port=new URL(origin).port;for(let i=0;i<90;i++){try{if((await fetch(origin+'/api/hosting',{signal:AbortSignal.timeout(1000)})).ok)return;}catch{}await delay(300);}throw Error(docker(['logs',name]));}
function inventory(){return docker(['exec',name,'node','-e',`const db=new(require('better-sqlite3'))('/app/data/cloud.sqlite',{readonly:true});const hash=require('node:crypto').createHash('sha256');for(const [table,cols]of [['documents','scope,id,data'],['assets','scope,workspace,id,name,mime,data'],['preferences','scope,key,value'],['cloud_users','id,email,name,salt,password'],['cloud_workspaces','id,name,created'],['cloud_members','workspace,user,role']]){const rows=db.prepare('SELECT '+cols+' FROM '+table+' ORDER BY rowid').all();if(table==='assets'&&db.prepare("SELECT 1 FROM sqlite_master WHERE name='asset_files'").get())for(const row of rows){const file=db.prepare('SELECT path FROM asset_files WHERE scope=? AND workspace=? AND id=?').get(row.scope,row.workspace,row.id);if(file)row.data=require('node:fs').readFileSync('/app/data/attachments/'+file.path);}hash.update(JSON.stringify(rows));};console.log(hash.digest('hex'));db.close();`]);}
async function logged(){const context=await browser.newContext();const r=await context.request.post(origin+'/api/cloud/login',{data:{email:'upgrade@example.com',password}});expect(r.status()).toBe(200);const page=await context.newPage();page.on('pageerror',e=>console.log('PAGEERROR',String(e)));await page.goto(origin);await expect(page.getByRole('button',{name:/^新建(文档|文件)$/}).first()).toBeVisible();return {context,page};}
async function createDoc(page,title,body){await page.getByRole('button',{name:'新建文档',exact:true}).first().click();await page.getByLabel('文档标题',{exact:true}).fill(title);await page.getByRole('textbox',{name:'文档正文'}).fill(body);await expect(page.getByRole('status').filter({hasText:'正在保存'})).toHaveCount(0);await delay(1000);}
async function openManager(page){await page.getByLabel('工作区选项',{exact:true}).click();await page.getByRole('button',{name:'团队工作空间',exact:true}).click();return page.getByRole('dialog',{name:'团队工作空间'});}
try{
 browser=await chromium.launch({headless:true,channel:'chromium'});await start(process.env.MYDOC_OLD_IMAGE||'mydoc-web:latest',volume,true);let {context,page}=await logged();
 await createDoc(page,'升级前私人文档','私人正文不能丢失');const privateUrl=page.url();
 let manager=await openManager(page);await manager.getByLabel('新工作空间名称').fill('升级保留团队');await manager.getByRole('button',{name:/^创建(团队|工作|团队工作)空间$/}).click();await manager.getByRole('button',{name:'打开',exact:true}).click();
 await createDoc(page,'升级前团队文档','团队正文不能丢失');const teamUrl=page.url();await page.getByLabel('添加本地附件').setInputFiles({name:'保留附件.txt',mimeType:'text/plain',buffer:Buffer.from('升级前附件字节')});await expect(page.getByRole('button',{name:'下载 保留附件.txt'})).toBeVisible();await delay(1200);await context.close();
 let before=inventory();const original=before;console.log('PASS old UI created personal/team documents and asset');
 docker(['kill',name]);docker(['rm',name]);
 for(let i=0;i<2;i++){
  await start(process.env.MYDOC_UPGRADE_IMAGE||'mydoc-web:2026.10.15-1');expect(inventory()).toBe(before);({context,page}=await logged());
  await page.goto(teamUrl);await expect(page.getByLabel('文档标题',{exact:true})).toHaveValue('升级前团队文档');await expect(page.getByRole('textbox',{name:'文档正文'})).toContainText('团队正文不能丢失');await expect(page.getByRole('button',{name:'下载 保留附件.txt'})).toBeVisible();
  const downloading=page.waitForEvent('download');await page.getByRole('button',{name:'下载 保留附件.txt'}).click();const stream=await(await downloading).createReadStream(),chunks=[];for await(const chunk of stream)chunks.push(chunk);expect(Buffer.concat(chunks).toString()).toBe('升级前附件字节');
  await page.goto(privateUrl);await expect(page.getByLabel('文档标题',{exact:true})).toHaveValue('升级前私人文档');await expect(page.getByRole('textbox',{name:'文档正文'})).toContainText('私人正文不能丢失');
  await context.close();before=inventory();docker(['stop',name]);docker(['rm',name]);console.log('PASS upgraded UI restart '+(i+1)+' preserved account, team, personal/team content and asset bytes');
 }
 // Restore the pre-migration snapshot into a fresh volume and open it with the
 // old image, without mixing the snapshot with another database's WAL files.
 const target=process.env.MYDOC_UPGRADE_IMAGE||'mydoc-web:2026.10.15-1';
 docker(['run','--rm','--network','none','--entrypoint','node','-v',volume+':/source:ro','-v',rollback+':/target',target,'-e',`const fs=require('node:fs'),dir='/source/upgrade-backups/';const rank=x=>{const m=JSON.parse(fs.readFileSync(dir+x));return m.schemaVersion*1e4+(m.cloudSchemaVersion??0);};const entry=fs.readdirSync(dir).filter(x=>x.endsWith('.sqlite.json')).sort((a,b)=>rank(a)-rank(b))[0];if(!entry)throw Error('missing pre-migration snapshot');fs.copyFileSync(dir+entry.slice(0,-5),'/target/cloud.sqlite');`]);
 await start(process.env.MYDOC_OLD_IMAGE||'mydoc-web:latest',rollback);expect(inventory()).toBe(original);({context,page}=await logged());
 await page.goto(teamUrl);await expect(page.getByLabel('文档标题',{exact:true})).toHaveValue('升级前团队文档');await expect(page.getByRole('textbox',{name:'文档正文'})).toContainText('团队正文不能丢失');
 await page.goto(privateUrl);await expect(page.getByRole('textbox',{name:'文档正文'})).toContainText('私人正文不能丢失');await context.close();
 console.log('PASS pre-migration snapshot opens in original image with original business bytes and visible documents');
}finally{await browser?.close();spawnSync('docker',['rm','-f',name],{stdio:'ignore'});for(const v of [volume,rollback])spawnSync('docker',['volume','rm',v],{stdio:'ignore'});}
