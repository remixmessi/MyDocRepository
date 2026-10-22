import {createNote} from './helpers';
import {test,expect,_electron as electron} from '@playwright/test';
import Database from 'better-sqlite3';
import {mkdtemp,mkdir,readFile,rm,readdir} from 'node:fs/promises';
import {tmpdir} from 'node:os';
import {join,resolve} from 'node:path';
import {unzipSync,strFromU8} from 'fflate';
import {SCHEMA_MIGRATIONS} from '../../packages/common/nbstore/src/schema';

test('桌面旧库附件启动迁移、上传文件落盘、Markdown ZIP 导出与重启读取',async()=>{
 const dir=await mkdtemp(join(tmpdir(),'mydoc-attachment-export-')),path=join(dir,'data','mydoc.sqlite');await mkdir(join(dir,'data'));
 const raw=new Database(path);raw.exec([1,2,3].map(v=>SCHEMA_MIGRATIONS[v]).join('\n'));raw.pragma('user_version=3');raw.prepare('INSERT INTO assets VALUES (?,?,?,?,?,?)').run('main','local','legacy','历史.txt','text/plain',Buffer.from('历史附件原字节'));raw.close();
 const launch=()=>electron.launch({args:[resolve('packages/frontend/apps/desktop')],env:{...process.env,MYDOC_DESKTOP_DB_PATH:path,MYDOC_DESKTOP_PROFILE:join(dir,'profile')}});
 let app=await launch();try{
  const page=await app.firstWindow();await expect(page.getByRole('button',{name:'新建文件',exact:true}).first()).toBeVisible();
  expect((await readdir(join(dir,'data','attachments'))).length).toBeGreaterThan(0);
  const legacy=await page.evaluate(()=>window.mydocDesktop!.storage({path:'/asset?workspace=local&id=legacy',method:'GET'}));expect(JSON.parse(new TextDecoder().decode(legacy.body)).data).toBe(Buffer.from('历史附件原字节').toString('base64'));
  await createNote(page);await page.getByLabel('文档标题',{exact:true}).fill('附件导出');await page.getByRole('textbox',{name:'文档正文'}).fill('带附件正文');await page.getByLabel('添加本地附件').setInputFiles({name:'报告.txt',mimeType:'text/plain',buffer:Buffer.from('导出附件字节')});await expect(page.getByRole('button',{name:'下载 报告.txt'})).toBeVisible();
  const exported=join(dir,'export.zip');await app.evaluate(({BrowserWindow},path)=>{BrowserWindow.getAllWindows()[0].webContents.session.on('will-download',(_event,item)=>item.setSavePath(path));},exported);
  await page.getByLabel('文档选项',{exact:true}).click();await page.getByRole('button',{name:'导出 Markdown',exact:true}).click();await expect.poll(async()=>{try{return Object.keys(unzipSync(await readFile(exported))).length;}catch{return 0;}}).toBe(2);const files=unzipSync(await readFile(exported));expect(strFromU8(files['附件导出.md'])).toContain('./assets/1-');expect(strFromU8(files['assets/1-报告.txt'])).toBe('导出附件字节');
  await app.close();app=await launch();const restored=await app.firstWindow();await expect(restored.getByRole('button',{name:'下载 报告.txt'})).toBeVisible();
 }finally{await app.close();await rm(dir,{recursive:true,force:true});}
});
