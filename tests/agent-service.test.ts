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
  const stream=await call('/subscribe?session='+id);expect(stream.headers.get('content-type')).toContain('text/event-stream');const reader=stream.body!.getReader();let received='';while(!received.includes('data: ')){const part=await reader.read();received+=new TextDecoder().decode(part.value);}expect(received).toContain('id: ');await reader.cancel();
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
it('desktop saves its own model through the API with encrypted keys; managed env config locks it',async()=>{
 const {startDesktopServer}=await import('../packages/frontend/apps/desktop/src/server');
 const {EncryptedFileSecrets}=await import('../packages/common/agent-host/src/model-config');
 const dir=mkdtempSync(join(tmpdir(),'mydoc-agent-desktop-model-')),cloud=new CloudServer(join(dir,'db.sqlite'));
 const cipher={isEncryptionAvailable:()=>true,encryptString:(v:string)=>Buffer.from([...v].reverse().join('')),decryptString:(b:Buffer)=>[...b.toString()].reverse().join('')};
 const desktop=await startDesktopServer(dir,cloud.database,undefined,fetch,new EncryptedFileSecrets(join(dir,'secrets'),cipher));
 try{
  const headers={Authorization:'Bearer '+desktop.token,'X-Mydoc-Generation':cloud.database.generation,'X-Mydoc-Account':'main'};
  const status=async()=>(await fetch(desktop.origin+'/api/agent/status',{headers})).json();
  expect(await status()).toMatchObject({enabled:false,userConfig:true,secureStorage:true,disclosure:expect.stringContaining('发送')});
  const save=(body:unknown)=>fetch(desktop.origin+'/api/agent/model',{method:'POST',headers,body:JSON.stringify(body)});
  expect((await save({endpoint:'http://10.1.1.1/v1',model:'m',maxOutputTokens:10,contextWindow:100,supportsTools:false,includeUsage:false})).status).toBe(400);
  const saved=await save({endpoint:'https://models.example.com/v1/chat/completions',model:'m',maxOutputTokens:10,contextWindow:100,supportsTools:false,includeUsage:false,apiKey:'desktop-secret'});
  expect(saved.status).toBe(200);const text=JSON.stringify(await saved.json())+JSON.stringify(await status());expect(text).not.toContain('desktop-secret');expect(text).toContain('"hasKey":true');
  expect(JSON.stringify(cloud.database.db.prepare('SELECT * FROM preferences').all())).not.toContain('desktop-secret');
  expect((await save({clear:true})).status).toBe(200);expect(await status()).toMatchObject({enabled:false});
  const {SQLiteAgentStore}=await import('../packages/common/agent-host/src/sqlite-store');const store=new SQLiteAgentStore(cloud.database,'main');store.createSession('locked','personal','排队测试');store.enqueue('locked','queued','q','hi');expect((await save({clear:true})).status).toBe(409);
 }finally{await desktop.close();cloud.close();rmSync(dir,{recursive:true,force:true});}
});
it('shared server rejects user model configuration and validates the managed endpoint allow-list',async()=>{
 const dir=mkdtempSync(join(tmpdir(),'mydoc-agent-managed-')),cloud=new CloudServer(join(dir,'db.sqlite'));
 try{
  expect(()=>new AgentService(cloud.database,'w',()=>undefined,{MYDOC_AGENT_ENDPOINT:'https://evil.example.com/v1',MYDOC_AGENT_MODEL:'m',MYDOC_AGENT_ALLOWED_HOSTS:'models.example.com'})).toThrow('ALLOWED_HOSTS');
  const agents=new AgentService(cloud.database,'w',()=>({id:'u',scope:'user:u'}),{MYDOC_AGENT_ENDPOINT:'https://models.example.com/v1',MYDOC_AGENT_MODEL:'m',MYDOC_AGENT_API_KEY:'server-secret'});
  const server=createServer((req,res)=>void agents.handler(req,res));await new Promise<void>(resolve=>server.listen(0,'127.0.0.1',resolve));const root='http://127.0.0.1:'+(server.address() as {port:number}).port;
  try{
   const body=await (await fetch(root+'/api/agent/status')).text();expect(body).not.toContain('server-secret');expect(JSON.parse(body)).toMatchObject({enabled:true,managed:true,userConfig:false});
   expect((await fetch(root+'/api/agent/model',{method:'POST',headers:{'X-Mydoc-Generation':cloud.database.generation},body:JSON.stringify({clear:true})})).status).toBe(403);
  }finally{await agents.close();await new Promise<void>(resolve=>server.close(()=>resolve()));}
 }finally{cloud.close();rmSync(dir,{recursive:true,force:true});}
});
it('Vite development mounts the same authenticated Agent service with an isolated database',async()=>{
 const {createServer:devServer}=await import('vite');const {cloudPlugin}=await import('../packages/frontend/apps/web/cloud-plugin');
 const dir=mkdtempSync(join(tmpdir(),'mydoc-agent-dev-test-'));const saved={db:process.env.MYDOC_CLOUD_DB_PATH,config:process.env.MYDOC_CONFIG_PATH,test:process.env.MYDOC_TEST_STORAGE};
 process.env.MYDOC_CLOUD_DB_PATH=join(dir,'cloud.sqlite');process.env.MYDOC_CONFIG_PATH=join(dir,'absent.json');delete process.env.MYDOC_TEST_STORAGE;
 let server:Awaited<ReturnType<typeof devServer>>|undefined;
 try{server=await devServer({configFile:false,plugins:[cloudPlugin()],server:{host:'127.0.0.1',port:0}});await server.listen();const root='http://127.0.0.1:'+(server.httpServer!.address() as {port:number}).port;expect((await fetch(root+'/api/agent/status')).status).toBe(401);}
 finally{await server?.close();for(const [key,value] of [['MYDOC_CLOUD_DB_PATH',saved.db],['MYDOC_CONFIG_PATH',saved.config],['MYDOC_TEST_STORAGE',saved.test]])if(value===undefined)delete process.env[key!];else process.env[key!]=value;rmSync(dir,{recursive:true,force:true});}
});
