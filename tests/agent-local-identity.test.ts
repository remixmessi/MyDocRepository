import {it,expect} from 'vitest';
import {createServer} from 'node:http';
import {mkdtempSync,rmSync,readFileSync,readdirSync} from 'node:fs';
import {tmpdir} from 'node:os';
import {join} from 'node:path';
import {CloudServer} from '../packages/frontend/apps/web/server/cloud';
import {storageHandler} from '../packages/common/nbstore/server/http';
import {createLocalAgentService} from '../packages/frontend/apps/web/server/agent/local';
import {AgentService} from '../packages/frontend/apps/web/server/agent/service';
it('local and server identities never substitute for one another; local configs and histories are isolated',async()=>{
 const dir=mkdtempSync(join(tmpdir(),'mydoc-agent-identities-')),cloud=new CloudServer(join(dir,'db.sqlite'));
 const local=createLocalAgentService(cloud.database,join(dir,'unused-worker.cjs'),{}),team=new AgentService(cloud.database,join(dir,'unused-worker.cjs'),req=>{const user=cloud.account(req);return user?{id:user.id,scope:'user:'+user.id}:undefined;},{});
 const server=createServer((req,res)=>{if(req.url?.startsWith('/api/local-agent/'))void local.handler(req,res);else if(req.url?.startsWith('/api/agent/'))void team.handler(req,res);else if(req.url?.startsWith('/api/storage/')){req.url=req.url.slice('/api/storage'.length);void storageHandler(cloud.database)(req,res);}else void cloud.handler(req,res);});await new Promise<void>(resolve=>server.listen(0,'127.0.0.1',resolve));const root='http://127.0.0.1:'+(server.address() as {port:number}).port;
 const call=(path:string,input?:unknown,headers:Record<string,string>={})=>fetch(root+path,{method:input===undefined?'GET':'POST',headers:{'Content-Type':'application/json','X-Mydoc-Generation':cloud.database.generation,...headers},body:input===undefined?undefined:JSON.stringify(input)});
 try{
  const register=async(email:string)=>{const response=await call('/api/storage/local-account',{action:'register',email,password:'password-test-123',name:'本地用户'});return {...await response.json(),cookie:response.headers.get('set-cookie')!.split(';')[0]};};
  const a=await register('same@example.com'),b=await register('other@example.com');const login=await call('/register',{email:'same@example.com',password:'password-team-123',name:'服务器用户'});const cookie=login.headers.get('set-cookie')!.split(';')[0];const teamUser=(await login.json()).user;
  const aHeaders={Cookie:a.cookie,'X-Mydoc-Account':a.user.id},bHeaders={Cookie:b.cookie,'X-Mydoc-Account':b.user.id};
  expect((await call('/api/local-agent/status',undefined,{Cookie:cookie})).status).toBe(401);
  expect((await call('/api/agent/status',undefined,aHeaders)).status).toBe(401);
  expect((await call('/api/local-agent/status',undefined,aHeaders)).status).toBe(200);
  expect((await call('/api/local-agent/status',undefined,{Cookie:a.cookie+'; '+cookie})).status).toBe(200);
  expect((await call('/api/local-agent/status',undefined,{'X-Mydoc-Local-Token':a.token})).status).toBe(401);
  const restored=await (await call('/api/storage/local-account',{action:'session',token:'wrong'},{Cookie:a.cookie})).json();expect(restored.user.id).toBe(a.user.id);
  expect((await call('/api/local-agent/status',undefined,{...aHeaders,'X-Mydoc-Account':teamUser.id})).status).toBe(409);
  const session=await call('/api/local-agent/sessions',{title:'本地历史'},aHeaders);expect(session.status).toBe(201);const {id}=await session.json();
  expect(await (await call('/api/local-agent/sessions',undefined,bHeaders)).json()).toEqual([]);
  expect(await (await call('/api/agent/sessions',undefined,{Cookie:cookie})).json()).toEqual([]);
  expect((await call('/api/local-agent/events?session='+id,undefined,bHeaders)).status).toBe(400);
  const saved=await call('/api/local-agent/model',{endpoint:'https://example.com/v1/chat/completions',model:'local-choice',maxOutputTokens:2048,contextWindow:32000,supportsTools:false,includeUsage:false,apiKey:'test-secret-key'},aHeaders);expect(saved.status).toBe(200);
  const own=await (await call('/api/local-agent/status',undefined,aHeaders)).json();expect(own.model.model).toBe('local-choice');expect(own.model.hasKey).toBe(true);expect(own.model.keyAvailable).toBe(true);
  expect((await (await call('/api/local-agent/status',undefined,bHeaders)).json()).model).toBeUndefined();
  const prefs=cloud.database.db.prepare("SELECT value FROM preferences WHERE scope LIKE 'agent-config:%'").all();expect(JSON.stringify(prefs)).not.toContain('test-secret-key');
  const files=readdirSync(join(dir,'local-agent-secrets')).filter(file=>file.endsWith('.bin'));expect(files.length).toBe(1);expect(readFileSync(join(dir,'local-agent-secrets',files[0])).includes(Buffer.from('test-secret-key'))).toBe(false);
  const verified=await call('/api/storage/local-account',{action:'login',email:'same@example.com',password:'password-test-123',verifyOnly:true},{Cookie:a.cookie});expect(verified.headers.get('set-cookie')).toBeNull();expect((await call('/api/local-agent/status',undefined,aHeaders)).status).toBe(200);
  await call('/api/storage/local-account',{action:'logout',token:'wrong'},{Cookie:a.cookie});expect((await call('/api/local-agent/status',undefined,aHeaders)).status).toBe(401);expect((await call('/api/agent/status',undefined,{Cookie:cookie})).status).toBe(200);
  await call('/logout',{}, {Cookie:cookie});expect((await call('/api/local-agent/status',undefined,bHeaders)).status).toBe(200);
 }finally{await local.close();await team.close();await new Promise<void>(resolve=>server.close(()=>resolve()));cloud.close();rmSync(dir,{recursive:true,force:true});}
});

it('Web cookies stay private and exposed legacy preference tokens are revoked atomically at startup',async()=>{
 const dir=mkdtempSync(join(tmpdir(),'mydoc-agent-private-session-')),cloud=new CloudServer(join(dir,'db.sqlite'));
 let local:ReturnType<typeof createLocalAgentService>|undefined;
 const server=createServer((req,res)=>{if(req.url?.startsWith('/api/local-agent/')){if(local)void local.handler(req,res);else{res.statusCode=503;res.end();}}else{req.url=req.url?.slice('/api/storage'.length);void storageHandler(cloud.database,false,undefined,undefined,undefined,true)(req,res);}});
 await new Promise<void>(resolve=>server.listen(0,'127.0.0.1',resolve));const root='http://127.0.0.1:'+(server.address() as {port:number}).port;
 const account={email:'private@example.com',name:'私有会话',password:'private-password-123'};
 const call=(path:string,input?:unknown,cookie?:string)=>fetch(root+path,{method:input===undefined?'GET':'POST',headers:{'Content-Type':'application/json','X-Mydoc-Generation':cloud.database.generation,...(cookie?{Cookie:cookie}:{})},body:input===undefined?undefined:JSON.stringify(input)});
 try{
  const registered=await call('/api/storage/local-account',{...account,action:'register'}),legacyCookie=registered.headers.get('set-cookie')!.split(';')[0],legacyToken=legacyCookie.split('=')[1];
  expect(await registered.json()).not.toHaveProperty('token');
  cloud.database.preference('main','local-account-session',legacyToken);cloud.database.preference('user:legacy','local-account-session',legacyToken);cloud.database.preference('main','keep','data');
  cloud.database.db.exec("CREATE TEMP TRIGGER reject_session_cleanup BEFORE DELETE ON preferences WHEN OLD.key='local-account-session' BEGIN SELECT RAISE(ABORT,'cleanup failed'); END");
  expect(()=>createLocalAgentService(cloud.database,join(dir,'unused-worker.cjs'),{})).toThrow('cleanup failed');
  expect(cloud.database.db.prepare('SELECT count(*) n FROM local_account_sessions').get()).toEqual({n:1});
  expect(cloud.database.db.prepare("SELECT count(*) n FROM preferences WHERE key='local-account-session'").get()).toEqual({n:2});
  cloud.database.db.exec('DROP TRIGGER reject_session_cleanup');
  local=createLocalAgentService(cloud.database,join(dir,'unused-worker.cjs'),{});
  const exposed=await (await call('/api/storage/status')).json();expect(exposed.preferences).not.toHaveProperty('local-account-session');expect(exposed.preferences.keep).toBe('data');
  expect(cloud.database.db.prepare("SELECT 1 FROM preferences WHERE key='local-account-session'").get()).toBeUndefined();
  expect((await call('/api/local-agent/sessions',undefined,legacyCookie)).status).toBe(401);
  expect((await (await call('/api/storage/local-account',{action:'session',token:legacyToken})).json()).user).toBeNull();
  const login=await call('/api/storage/local-account',{...account,action:'login'}),cookie=login.headers.get('set-cookie')!.split(';')[0];expect(await login.json()).not.toHaveProperty('token');
  expect((await call('/api/local-agent/sessions',{title:'private history'},cookie)).status).toBe(201);
  expect((await call('/api/local-agent/sessions')).status).toBe(401);
  expect((await fetch(root+'/api/local-agent/sessions',{headers:{'X-Mydoc-Local-Token':cookie.split('=')[1]}})).status).toBe(401);
  const leaked=JSON.stringify(await (await call('/api/storage/status')).json());expect(leaked).not.toContain(cookie.split('=')[1]);expect(leaked).not.toContain(legacyToken);
  await local.close();local=createLocalAgentService(cloud.database,join(dir,'unused-worker.cjs'),{});
  expect(await (await call('/api/local-agent/sessions',undefined,cookie)).json()).toMatchObject([{title:'private history'}]);
 }finally{await local?.close();server.closeAllConnections();await new Promise<void>(resolve=>server.close(()=>resolve()));cloud.close();rmSync(dir,{recursive:true,force:true});}
});
