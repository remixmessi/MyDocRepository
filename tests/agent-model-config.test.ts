import {it,expect} from 'vitest';
import Database from 'better-sqlite3';
import {mkdtempSync,readdirSync,readFileSync,rmSync} from 'node:fs';
import {tmpdir} from 'node:os';
import {join} from 'node:path';
import {SQLiteProjectDatabase} from '../packages/common/nbstore/server/database';
import {EncryptedFileSecrets,EnvSecrets,ModelConfigStore,ModelError,ResilientModel,classifyModelError} from '../packages/common/agent-host/src/model-config';
import type {ModelGateway,ProviderEvent} from '@mydoc/agent-contracts';
const cipher=(available=true)=>({isEncryptionAvailable:()=>available,encryptString:(value:string)=>Buffer.from('enc:'+Buffer.from(value).toString('base64')),decryptString:(value:Buffer)=>Buffer.from(value.toString().slice(4),'base64').toString()});
function fixture(available=true,allowed?:string[]){const dir=mkdtempSync(join(tmpdir(),'mydoc-agent-model-'));const db=new SQLiteProjectDatabase(join(dir,'db.sqlite'),p=>new Database(p));const secrets=new EncryptedFileSecrets(join(dir,'secrets'),cipher(available));return {dir,db,secrets,store:new ModelConfigStore(db,secrets,allowed),close(){db.close();rmSync(dir,{recursive:true,force:true});}};}
const settings={endpoint:'https://models.example.com/v1/chat/completions',model:'m',maxOutputTokens:1024,contextWindow:32000,supportsTools:true,includeUsage:true};
it('stores only a secret reference; the key never reaches SQLite or the public view',()=>{
 const f=fixture();try{
  f.store.save({...settings,apiKey:'secret-key-123'});
  const row=f.db.db.prepare("SELECT value FROM preferences WHERE scope='agent-config'").get() as {value:string};
  expect(row.value).not.toContain('secret-key-123');expect(JSON.parse(row.value).secretRef).toMatch(/^safe:/);
  expect(readFileSync(join(f.dir,'secrets',readdirSync(join(f.dir,'secrets'))[0])).toString()).not.toContain('secret-key-123');
  expect(f.store.publicView()).toEqual({...settings,hasKey:true,keyAvailable:true});expect(JSON.stringify(f.store.publicView())).not.toContain('safe:');
  f.store.save({...settings,model:'m2',keepKey:true});expect(f.store.publicView()).toMatchObject({model:'m2',hasKey:true});expect(readdirSync(join(f.dir,'secrets'))).toHaveLength(1);
  expect(()=>f.store.save({...settings,endpoint:'https://other.example.com/v1/chat',keepKey:true})).toThrow('重新输入密钥');
  f.store.save({...settings,apiKey:'second-key'});expect(readdirSync(join(f.dir,'secrets'))).toHaveLength(1);
  f.store.clear();expect(f.store.publicView()).toBeUndefined();expect(readdirSync(join(f.dir,'secrets'))).toHaveLength(0);
 }finally{f.close();}
});
it('refuses to save keys without OS encryption and reports an unavailable stored key',()=>{
 const f=fixture(false);try{
  expect(()=>f.store.save({...settings,apiKey:'k'})).toThrow('系统加密不可用');expect(f.store.publicView()).toBeUndefined();
  f.store.save(settings);expect(f.store.publicView()).toMatchObject({hasKey:false,keyAvailable:true});
  f.db.preference('agent-config','agent.model',JSON.stringify({...settings,secretRef:'safe:00000000-0000-0000-0000-000000000000'}));
  expect(f.store.publicView()).toMatchObject({hasKey:true,keyAvailable:false});expect(()=>f.store.provider()).toThrow('密钥不可用');
 }finally{f.close();}
});
it('validates endpoints: HTTPS or loopback HTTP, configured allow-list, no credentials or query',()=>{
 const f=fixture(true,['models.example.com']);try{
  for(const endpoint of ['http://10.0.0.5/v1/chat','https://user:pw@models.example.com/x','https://models.example.com/x?key=1','https://evil.example.com/v1/chat','file:///etc/passwd'])expect(()=>f.store.save({...settings,endpoint})).toThrow();
  f.store.save(settings);expect(f.store.load()?.endpoint).toBe(settings.endpoint);
  expect(new EnvSecrets({K:'v'}).get('env:K')).toBe('v');expect(new EnvSecrets({K:'v'}).get('K')).toBeUndefined();expect(()=>new EnvSecrets().put()).toThrow();
 }finally{f.close();}
});
const gateway=(script:(attempt:number)=>AsyncIterable<ProviderEvent>):ModelGateway&{attempts:number}=>{const g={attempts:0,capabilities:{tools:true,contextWindow:1000,maxOutputTokens:100},stream(){return script(g.attempts++);}};return g;};
const collect=async(model:ModelGateway,signal=new AbortController().signal)=>{const out:ProviderEvent[]=[];for await(const event of model.stream([],[],signal))out.push(event);return out;};
const fast={retries:2,baseDelayMs:1,firstEventTimeoutMs:200};
it('retries only transient failures before the first event, then classifies errors',async()=>{
 const flaky=gateway(async function*(attempt){if(attempt<2)throw new Error('Model HTTP status 503');yield {type:'text',text:'ok'};yield {type:'done'};});
 expect(await collect(new ResilientModel(flaky,fast))).toHaveLength(2);expect(flaky.attempts).toBe(3);
 const auth=gateway(async function*(){throw new Error('Model HTTP status 401');});
 await expect(collect(new ResilientModel(auth,fast))).rejects.toMatchObject({code:'auth'});expect(auth.attempts).toBe(1);
 const partial=gateway(async function*(){yield {type:'text',text:'half'};throw new Error('Model HTTP status 503');});
 await expect(collect(new ResilientModel(partial,fast))).rejects.toMatchObject({code:'unavailable'});expect(partial.attempts).toBe(1);
 const down=gateway(async function*(){throw new TypeError('fetch failed');});
 await expect(collect(new ResilientModel(down,fast))).rejects.toMatchObject({code:'network'});expect(down.attempts).toBe(3);
 const slow=gateway(async function*(_attempt){await new Promise(resolve=>setTimeout(resolve,1000));yield {type:'done'};});
 await expect(collect(new ResilientModel(slow,{...fast,retries:0,firstEventTimeoutMs:20}))).rejects.toMatchObject({code:'timeout'});
 const controller=new AbortController();controller.abort();
 await expect(collect(new ResilientModel(flaky,fast),controller.signal)).rejects.toBeInstanceOf(ModelError);
 expect(classifyModelError(new Error('Model HTTP status 429'))).toMatchObject({code:'rate_limited',retryable:true});
 expect(classifyModelError(new Error('Model HTTP status 400 body secret'))).toMatchObject({code:'invalid_response'});expect(classifyModelError(new Error('x')).message).not.toContain('x');
});
it('preserves native API types, defaults legacy settings to chat, and rejects unknown API types',async()=>{
 const f=fixture();try{
  f.store.save(settings);expect(f.store.provider()).toBeDefined();
  for(const apiType of ['openai-responses','anthropic-messages'] as const){f.store.save({...settings,apiType});expect(f.store.publicView()?.apiType).toBe(apiType);expect(f.store.provider()).toBeDefined();
   const end=apiType==='openai-responses'?[{type:'response.completed',response:{status:'completed',output:[]}}]:[{type:'message_start',message:{usage:{input_tokens:0,output_tokens:0}}},{type:'message_delta',delta:{stop_reason:'end_turn'},usage:{output_tokens:0}},{type:'message_stop'}];
   let body:Record<string,unknown>={};const provider=f.store.provider(async(_url,init)=>{body=JSON.parse(init!.body as string);return new Response(end.map(event=>'data: '+JSON.stringify(event)+'\n\n').join(''),{headers:{'Content-Type':'text/event-stream'}});})!;
   expect((await collect(provider)).at(-1)).toEqual({type:'done'});expect(body).toHaveProperty(apiType==='openai-responses'?'input':'messages');}
  expect(()=>f.store.save({...settings,apiType:'unknown' as never})).toThrow('API type');expect(f.store.publicView()?.apiType).toBe('anthropic-messages');
 }finally{f.close();}
});
it('upgrades a legacy connection on first edit, keeps its key and routes the chosen default model',async()=>{
 const f=fixture();try{
  f.store.save({...settings,apiKey:'original-secret'});const ref=f.store.load()!.secretRef;
  expect(f.store.publicConnections()).toMatchObject({connections:[{id:'legacy',name:'原有模型连接',models:['m'],hasKey:true}],defaultModel:{connectionId:'legacy',model:'m'}});
  expect(f.db.db.prepare("SELECT 1 FROM preferences WHERE key='agent.connections'").get()).toBeUndefined();
  const second=f.store.saveConnection({...settings,name:'第二服务',models:['m','other'],enabled:true,endpoint:'https://other.example.com/v1/responses',apiType:'openai-responses',apiKey:'second-secret'});
  expect(f.store.load()!.secretRef).toBe(ref);expect(f.store.connections().connections).toHaveLength(2);
  const {secretRef:secondRef,...secondInput}=f.store.connections().connections.find(c=>c.id===second)!;
  f.store.saveConnection({...secondInput,name:'第二服务（编辑）',enabled:false,keepKey:true});expect(()=>f.store.setDefault(second,'other')).toThrow('不可用');
  f.store.saveConnection({...secondInput,enabled:true,keepKey:true});expect(f.store.connections().connections.find(c=>c.id===second)?.secretRef).toBe(secondRef);
  f.store.setDefault(second,'other');expect(f.store.publicView()).toMatchObject({model:'other',endpoint:'https://other.example.com/v1/responses',apiType:'openai-responses'});
  let sent:Record<string,unknown>={};const provider=f.store.provider(async(_url,init)=>{sent=JSON.parse(init!.body as string);return new Response('data: '+JSON.stringify({type:'response.completed',response:{status:'completed',output:[]}})+'\n\n',{headers:{'Content-Type':'text/event-stream'}});})!;
  await collect(provider);expect(sent.model).toBe('other');
  expect(JSON.stringify(f.store.publicConnections())).not.toMatch(/original-secret|second-secret|safe:/);
  const again=new ModelConfigStore(f.db,f.secrets);expect(again.publicConnections()).toEqual(f.store.publicConnections());
  expect(()=>f.store.deleteConnection(second)).toThrow('其他默认');
  expect(()=>f.store.setDefault(second,'missing')).toThrow('不可用');
  const connection=f.store.connections().connections.find(c=>c.id===second)!;const {secretRef,...input}=connection;
  expect(()=>f.store.saveConnection({...input,enabled:false,keepKey:true})).toThrow('其他默认');
  expect(()=>f.store.saveConnection({...input,models:['m'],keepKey:true})).toThrow('其他默认');
  f.store.setDefault('legacy','m');f.store.deleteConnection(second);expect(f.store.connections().connections).toHaveLength(1);expect(readdirSync(join(f.dir,'secrets'))).toHaveLength(1);
  f.store.save({...settings,model:'updated',keepKey:true});expect(f.store.publicView()?.model).toBe('updated');expect(f.store.connections().connections[0].models).toContain('updated');
 }finally{f.close();}
});
it('isolates collections by account and rolls back collection/default writes and new secrets on failure',()=>{
 const f=fixture();try{
  f.store.save({...settings,apiKey:'old-secret'});const before=f.store.publicConnections();
  const other=new ModelConfigStore(f.db,f.secrets,undefined,'other-account');expect(other.connections().connections).toEqual([]);
  other.saveConnection({...settings,name:'Other',enabled:true,models:['m']});expect(f.store.publicConnections()).toEqual(before);
  const original=f.db.preference.bind(f.db);f.db.preference=(scope,key,value)=>{if(key==='agent.model')throw new Error('disk failed');return original(scope,key,value);};
  expect(()=>f.store.saveConnection({...settings,name:'broken',enabled:true,models:['m'],apiKey:'new-secret'})).toThrow('disk failed');
  expect(f.store.publicConnections()).toEqual(before);expect(readdirSync(join(f.dir,'secrets'))).toHaveLength(1);
  f.db.preference=original;
  expect(()=>f.store.saveConnection({...settings,name:'bad',enabled:true,models:['m','m']})).toThrow('无效');
 }finally{f.close();}
});
