import {it,expect} from 'vitest';
import {createServer} from 'node:http';
import {build} from 'esbuild';
import {mkdtempSync,rmSync} from 'node:fs';
import {tmpdir} from 'node:os';
import {join} from 'node:path';
import {CloudServer} from '../packages/frontend/apps/web/server/cloud';
import {AgentService} from '../packages/frontend/apps/web/server/agent/service';
it('production service uses real cookie identity, fences requests and executes a prompt in a child',async()=>{
 const dir=mkdtempSync(join(tmpdir(),'mydoc-agent-api-'));const cloud=new CloudServer(join(dir,'db.sqlite'));
 const model=createServer(async(req,res)=>{for await(const _ of req){}res.writeHead(200,{'Content-Type':'text/event-stream'});res.end('data: '+JSON.stringify({choices:[{index:0,delta:{content:'模型答复'},finish_reason:'stop'}]})+'\n\ndata: [DONE]\n\n');});
 await new Promise<void>(resolve=>model.listen(0,'127.0.0.1',resolve));
 const worker=join(dir,'worker.cjs');await build({entryPoints:['packages/common/agent-host/src/worker-child.ts'],outfile:worker,bundle:true,platform:'node',format:'cjs'});
 const agents=new AgentService(cloud.database,worker,req=>{const account=cloud.account(req);return account?{id:account.id,scope:'user:'+account.id}:undefined;},{MYDOC_AGENT_ENDPOINT:'http://127.0.0.1:'+(model.address() as {port:number}).port+'/chat',MYDOC_AGENT_MODEL:'test'});
 const server=createServer((req,res)=>{if(req.url?.startsWith('/api/agent/'))void agents.handler(req,res);else void cloud.handler(req,res);});await new Promise<void>(resolve=>server.listen(0,'127.0.0.1',resolve));const root='http://127.0.0.1:'+(server.address() as {port:number}).port;
 try{
  const register=async(email:string)=>{const response=await fetch(root+'/register',{method:'POST',body:JSON.stringify({email,name:'测试',password:'test-password-123'})});return {cookie:response.headers.get('set-cookie')!.split(';')[0],id:(await response.json()).user.id};};
  const a=await register('a@example.com'),b=await register('b@example.com');
  const call=(path:string,method='GET',input?:unknown,user=a,headers:Record<string,string>={})=>fetch(root+'/api/agent'+path,{method,headers:{Cookie:user.cookie,'X-Mydoc-Account':user.id,'X-Mydoc-Generation':cloud.database.generation,...headers},body:input?JSON.stringify(input):undefined});
  expect((await fetch(root+'/api/agent/status')).status).toBe(401);
  expect((await call('/sessions','POST',{title:'x'},a,{Origin:'https://evil.example'})).status).toBe(403);
  expect((await call('/sessions','POST',{title:'x'},a,{'X-Mydoc-Account':b.id})).status).toBe(409);
  expect((await call('/sessions','POST',{title:'x'},a,{'X-Mydoc-Generation':'old'})).status).toBe(409);
  const created=await call('/sessions','POST',{title:'问答'});expect(created.status).toBe(201);const {id}=await created.json();
  expect(await (await call('/sessions','GET',undefined,b)).json()).toEqual([]);
  const runResponse=await call('/runs','POST',{session:id,requestKey:'k',prompt:'你好'});expect(runResponse.status).toBe(202);const run=await runResponse.json();
  let state=run;for(let i=0;i<100&&state.status!=='completed';i++){await new Promise(resolve=>setTimeout(resolve,10));state=await (await call('/runs?id='+run.runId)).json();}
  expect(state.status).toBe('completed');expect((await call('/runs?id='+run.runId,'GET',undefined,b)).status).toBe(404);
  const events=await (await call('/events?session='+id)).json();expect(events.some((event:{payload:string})=>event.payload.includes('模型答复'))).toBe(true);
  const retry=await (await call('/runs','POST',{session:id,requestKey:'k',prompt:'你好'})).json();expect(retry.runId).toBe(run.runId);
  await fetch(root+'/logout',{method:'POST',headers:{Cookie:a.cookie}});expect((await call('/sessions')).status).toBe(401);
 }finally{await agents.close();await new Promise<void>(resolve=>server.close(()=>resolve()));await new Promise<void>(resolve=>model.close(()=>resolve()));cloud.close();rmSync(dir,{recursive:true,force:true});}
});
it('desktop service requires its existing bearer token and binds the main local scope',async()=>{
 const {startDesktopServer}=await import('../packages/frontend/apps/desktop/src/server');
 const dir=mkdtempSync(join(tmpdir(),'mydoc-agent-desktop-api-')),cloud=new CloudServer(join(dir,'db.sqlite'));
 const desktop=await startDesktopServer(dir,cloud.database);
 try{
  expect((await fetch(desktop.origin+'/api/agent/status')).status).toBe(401);
  const headers={Authorization:'Bearer '+desktop.token,'X-Mydoc-Generation':cloud.database.generation,'X-Mydoc-Account':'main'};
  const created=await fetch(desktop.origin+'/api/agent/sessions',{method:'POST',headers,body:JSON.stringify({title:'本机'})});expect(created.status).toBe(201);
  expect((await fetch(desktop.origin+'/api/agent/sessions',{headers:{...headers,'X-Mydoc-Account':'other'}})).status).toBe(409);
  expect((await fetch(desktop.origin+'/api/agent/sessions',{headers:{...headers,Origin:'https://evil.example'}})).status).toBe(403);
 }finally{await desktop.close();cloud.close();rmSync(dir,{recursive:true,force:true});}
});
