import {it,expect,afterEach,vi} from 'vitest';
import {mkdtempSync,writeFileSync,readFileSync,existsSync,rmSync} from 'node:fs';
import {tmpdir} from 'node:os';
import {join} from 'node:path';
import {OfficeParserPool,runOfficeParserProcess} from '../packages/frontend/apps/web/server/office-parser-process';
import * as Y from 'yjs';

afterEach(()=>vi.unstubAllEnvs());
it('受限子进程生成可重开的协作模型，错误文件不影响下一次解析',async()=>{
 const pool=new OfficeParserPool();
 try{
  const update=await pool.parse('csv',new TextEncoder().encode('标题,编号\n中文,001\n'));
  const doc=new Y.Doc();try{Y.applyUpdate(doc,update);expect((doc.getArray('sheet').get(1) as Y.Array<string>).toArray()).toEqual(['中文','001']);}finally{doc.destroy();}
  await expect(pool.parse('docx',new Uint8Array([1,2,3]))).rejects.toMatchObject({status:400});
  await expect(pool.parse('csv',new TextEncoder().encode('ok'))).resolves.toBeInstanceOf(Uint8Array);
  expect(pool.activeCount).toBe(0);
 }finally{pool.close();}
});
it('禁止读写业务文件、启动子进程，并移除继承环境变量',async()=>{
 const dir=mkdtempSync(join(tmpdir(),'mydoc-parser-permissions-')),file=join(dir,'business.txt'),output=join(dir,'output.txt');writeFileSync(file,'private');
 vi.stubEnv('MYDOC_TEST_PARSER_SECRET','private-env');
 const source=`const fs=require('node:fs');const result={secret:process.env.MYDOC_TEST_PARSER_SECRET,argv:process.execArgv};
 for(const [key,fn]of Object.entries({read:()=>fs.readFileSync(${JSON.stringify(file)}),write:()=>fs.writeFileSync(${JSON.stringify(output)},'changed'),child:()=>require('node:child_process').spawnSync(process.execPath,['-e','']),worker:()=>new(require('node:worker_threads').Worker)('',{eval:true})})){try{fn();result[key]='ALLOWED';}catch(error){result[key]=error.code;}}
 process.stdout.end(Buffer.concat([Buffer.from([0]),Buffer.from(JSON.stringify(result))]));`;
 try{
  const bytes=await runOfficeParserProcess(source,'csv',new Uint8Array());const result=JSON.parse(Buffer.from(bytes).toString());
  expect(result).toMatchObject({read:'ERR_ACCESS_DENIED',write:'ERR_ACCESS_DENIED',child:'ERR_ACCESS_DENIED',worker:'ERR_ACCESS_DENIED'});expect(result.secret).toBeUndefined();expect(result.argv).toContain('--max-old-space-size=256');expect(readFileSync(file,'utf8')).toBe('private');expect(existsSync(output)).toBe(false);
 }finally{rmSync(dir,{recursive:true,force:true});}
});
it('终止超时、取消、崩溃和过大输出，不阻塞父事件循环',async()=>{
 let ticks=0;const timer=setInterval(()=>ticks++,10);
 try{
  await expect(runOfficeParserProcess('while(true){}','csv',new Uint8Array(),{timeoutMs:250})).rejects.toMatchObject({status:422});expect(ticks).toBeGreaterThan(0);
  const controller=new AbortController(),job=runOfficeParserProcess('while(true){}','csv',new Uint8Array(),{signal:controller.signal});controller.abort();await expect(job).rejects.toMatchObject({status:499});
  await expect(runOfficeParserProcess('process.exit(2)','csv',new Uint8Array())).rejects.toMatchObject({status:503});
  await expect(runOfficeParserProcess('process.stdout.end(Buffer.alloc(1000))','csv',new Uint8Array(),{outputBytes:10})).rejects.toMatchObject({status:413});
 }finally{clearInterval(timer);}
});
it('并发上限包含准备阶段，关闭服务取消尚未启动的任务',async()=>{
 const pool=new OfficeParserPool({source:()=>new Promise(()=>{})});
 const a=pool.parse('csv',new Uint8Array()),b=pool.parse('csv',new Uint8Array());
 const outcomes=Promise.allSettled([a,b]);expect(pool.activeCount).toBe(2);
 await expect(pool.parse('csv',new Uint8Array())).rejects.toMatchObject({status:503});pool.close();
 expect((await outcomes).map(result=>result.status==='rejected'?result.reason.status:0)).toEqual([499,499]);expect(pool.activeCount).toBe(0);
 await expect(pool.parse('csv',new Uint8Array())).rejects.toMatchObject({status:503});
});
it('解析程序准备失败或超时后释放名额',async()=>{
 const failed=new OfficeParserPool({source:async()=>{throw new Error('build failed');}});
 await expect(failed.parse('csv',new Uint8Array())).rejects.toMatchObject({status:503});expect(failed.activeCount).toBe(0);failed.close();
 const delayed=new OfficeParserPool({source:()=>new Promise(()=>{}),timeoutMs:30});await expect(delayed.parse('csv',new Uint8Array())).rejects.toMatchObject({status:422});expect(delayed.activeCount).toBe(0);delayed.close();
});

it('解析失败或解析后撤权时，附件和资源不会被写入',async()=>{
 const {CloudServer}=await import('../packages/frontend/apps/web/server/cloud');
 const {importOfficeOriginal}=await import('../packages/frontend/apps/web/server/office-import');
 const {OfficeParserError}=await import('../packages/frontend/apps/web/server/office-parser-process');
 const dir=mkdtempSync(join(tmpdir(),'mydoc-parser-no-commit-')),cloud=new CloudServer(join(dir,'cloud.sqlite'));
 const put=vi.spyOn(cloud.database.attachments,'put');
 try{
  await expect(importOfficeOriginal(cloud.database,'workspace','user','cells.csv',new Uint8Array(),async()=>{throw new OfficeParserError('超时',422);})).rejects.toMatchObject({status:422});
  let finish:(value:Uint8Array)=>void=()=>{};const parsed=new Promise<Uint8Array>(resolve=>finish=resolve);
  const recheck=vi.fn(()=>{throw new Error('已撤销写权限');});
  const importing=importOfficeOriginal(cloud.database,'workspace','user','cells.csv',new Uint8Array(),()=>parsed,recheck);
  expect(put).not.toHaveBeenCalled();expect(recheck).not.toHaveBeenCalled();finish(new Uint8Array([0,0]));
  await expect(importing).rejects.toThrow('已撤销写权限');expect(recheck).toHaveBeenCalledOnce();expect(put).not.toHaveBeenCalled();
  for(const table of ['cloud_office_resources','cloud_asset_owners','assets','documents'])expect(cloud.database.db.prepare('SELECT count(*) AS count FROM '+table).get()).toEqual({count:0});
 }finally{put.mockRestore();cloud.close();rmSync(dir,{recursive:true,force:true});}
});
